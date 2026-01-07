Attribute VB_Name = "Plugin_EZTwain"
'***************************************************************************
'Scanner Interface
'Copyright 2001-2026 by Tanner Helland
'Created: 1/10/01
'Last updated: 12/May/16
'Last update: clean up the interface a bit
'
'Module for handling all TWAIN32 acquisition features.  This module relies heavily
' upon the EZTW32.dll file, which is required because VB does not have native scanner support.
'
'The EZTW32 library is a free, public domain TWAIN32-compliant library.  You can learn more
' about it at http://eztwain.com/
'
'This project was designed against v1.19 of the EZTW32 library (2009.02.22).  It may not work with
' other versions of the library.  Additional documentation regarding the use of EZTW32 is
' available from the EZTW32 developers at http://eztwain.com/ezt1_download.htm
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function TWAIN_AcquireToFilename Lib "eztw32.dll" (ByVal hWndApp As Long, ByVal sFile As String) As Long
Private Declare Function TWAIN_SelectImageSource Lib "eztw32.dll" (ByVal hWndApp As Long) As Long
Private Declare Function TWAIN_IsAvailable Lib "eztw32.dll" () As Long
'Private Declare Function TWAIN_CloseSourceManager Lib "eztw32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function TWAIN_UnloadSourceManager Lib "eztw32.dll" () As Long
Private Declare Function TWAIN_EasyVersion Lib "eztw32.dll" () As Long

'Used to load and unload the EZTW32 dll from an arbitrary location (in our case, the \Data\Plugins subdirectory)
Private m_hLibScanner As Long

'Once the EnableScanner function has been called, its result will be cached here
Private m_ScanningAvailable As Boolean

'Initialize EZTwain.  Do not call this until you have verified EZTwain's existence (typically via the PluginManager module)
Public Function InitializeEZTwain() As Boolean
    
    If (m_hLibScanner = 0) Then
    
        'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
        Dim eztPath As String
        eztPath = PluginManager.GetPluginPath & "eztw32.dll"
        m_hLibScanner = VBHacks.LoadLib(eztPath)
        InitializeEZTwain = (m_hLibScanner <> 0)
        m_ScanningAvailable = InitializeEZTwain
        
        If (Not InitializeEZTwain) Then
            PDDebug.LogAction "WARNING!  LoadLibrary failed to load EZTwain.  Last DLL error: " & Err.LastDllError
            PDDebug.LogAction "(FYI, the attempted path was: " & eztPath & ")"
        End If
        
    Else
        InitializeEZTwain = True
    End If
    
End Function

'When PD closes, make sure to release our open handle!
Public Sub ReleaseEZTwain()
    VBHacks.FreeLib m_hLibScanner
End Sub

'Return the EZTwain version number, as a string
Public Function GetEZTwainVersion() As String
    
    InitializeEZTwain
    If (m_hLibScanner <> 0) Then
    
        'The TWAIN version is the version number * 100.  Modify the return string accordingly
        Dim ezVer As Long
        ezVer = TWAIN_EasyVersion()
        GetEZTwainVersion = (ezVer \ 100) & "." & (ezVer Mod 100) & ".0.0"
        
    Else
        RaiseInternalDebugEvent "LoadLibrary Failed", "GetEZTwainVersion() failed to load base library"
    End If
    
End Function

'This should be run before the scanner is accessed
Public Function EnableScanner() As Boolean
    
    InitializeEZTwain
    
    If (m_hLibScanner <> 0) Then
        EnableScanner = (TWAIN_IsAvailable() <> 0)
        If (Not EnableScanner) Then RaiseInternalDebugEvent "TWAIN_IsAvailable() Failed", "TWAIN_IsAvailable() returned 0"
    Else
        RaiseInternalDebugEvent "LoadLibrary Failed", "EnableScanner() failed to load base library"
        EnableScanner = False
    End If
    
End Function

Public Sub ForciblySetScannerAvailability(ByVal newState As Boolean)
    m_ScanningAvailable = newState
End Sub

Public Function IsScannerAvailable() As Boolean
    IsScannerAvailable = m_ScanningAvailable
End Function

'Allow the user to select which hardware PhotoDemon will use for scanning
Public Sub Twain32SelectScanner()
    
    If m_ScanningAvailable Then
        
        'Make sure the scanner is plugged in and powered up
        If EnableScanner Then
            TWAIN_SelectImageSource GetModalOwner().hWnd
            Message "Scanner successfully enabled "
        Else
            PDMsgBox "The selected scanner or digital camera isn't responding." & vbCrLf & vbCrLf & "Please make sure the device is turned on and ready for use.", vbExclamation Or vbOKOnly, "Error"
            Message "Scanner error"
        End If
    
    Else
        PDMsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into the plug-in directory and reload the program.", vbCritical Or vbOKOnly, "Error"
        Message "Scanner error"
    End If
    
End Sub

'Acquire an image from the currently selected scanner
Public Sub Twain32Scan()

    Message "Acquiring image..."
    
    'Make sure the EZTW32.dll file exists
    If IsScannerAvailable Then
        
        'Make sure the scanner is on and responsive
        If EnableScanner Then
            
            'This function has a fairly extensive error handling routine, since a lot can go wrong when scanning
            On Error GoTo ScanError
            
            Dim scannerCaptureFile As String, scanCheck As Long
            
            'ScanCheck is used to store the return values of the EZTW32.dll scanner functions.  We start by setting it
            ' to an arbitrary value that only we know; if an error occurs and this value is still present, it means an
            ' error occurred outside of the EZTW32 library.
            scanCheck = -5
            
            'A temporary file is required by the scanner; we will place it in the project folder, then delete it when finished
            scannerCaptureFile = UserPrefs.GetTempPath & "PDScanInterface.tmp"
                
            'This line uses the EZTW32.dll file to scan the image and send it to a temporary file
            scanCheck = TWAIN_AcquireToFilename(GetModalOwner().hWnd, scannerCaptureFile)
                
            'If the image was successfully scanned, load it
            If (scanCheck = 0) Then
                
                Dim sTitle As String
                sTitle = g_Language.TranslateMessage("Scanned Image")
                sTitle = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
                Loading.LoadFileAsNewImage scannerCaptureFile, sTitle, False
                
                'Be polite and remove the temporary file acquired from the scanner
                Files.FileDeleteIfExists scannerCaptureFile
                Message "Image acquired successfully "
                
                If FormMain.Enabled Then FormMain.MainCanvas(0).SetFocusToCanvasView
                
            Else
                'If the scan was unsuccessful, let the user know what happened
                GoTo ScanError
            End If
            
        Else
            PDMsgBox "The selected scanner or digital camera isn't responding." & vbCrLf & vbCrLf & "Please make sure the device is turned on and ready for use.", vbExclamation Or vbOKOnly, "Error"
            Message "Scanner error"
        End If
        
    Else
        PDMsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into the plug-in directory and reload the program.", vbCritical Or vbOKOnly, "Error"
        Message "Scanner error"
    End If
    
    Exit Sub

'Something went wrong
ScanError:
    
    Dim scanErrMessage As String
    
    Select Case scanCheck
        Case -5
            scanErrMessage = g_Language.TranslateMessage("Unknown error occurred.  Please make sure your scanner is turned on and ready for use.")
        Case -4
            scanErrMessage = g_Language.TranslateMessage("Scan successful, but temporary file save failed.  Is it possible that your hard drive is full (or almost full)?")
        Case -3
            scanErrMessage = g_Language.TranslateMessage("Unable to acquire DIB lock.  Please make sure no other programs are accessing the scanner.  If the problem persists, reboot and try again.")
        Case -2
            scanErrMessage = g_Language.TranslateMessage("Temporary file access error.  This can be caused when running on a system with limited access rights.  Please enable admin rights and try again.")
        Case -1
            scanErrMessage = g_Language.TranslateMessage("Scan canceled at the user's request.")
            Message "Scan canceled"
            Exit Sub
        Case Else
            scanErrMessage = g_Language.TranslateMessage("The scanner returned an error code that wasn't specified in the EZTW32.dll documentation (Error #%1).  Please visit http://www.eztwain.com for more information.", scanCheck)
    End Select
        
    PDMsgBox scanErrMessage, vbExclamation Or vbOKOnly, "Scan canceled"
    Message "Scan canceled"
    
End Sub

Private Sub RaiseInternalDebugEvent(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString)
    PDDebug.LogAction "WARNING!  EZTwain interface reported error """ & errName & """ - " & errDescription
End Sub
