Attribute VB_Name = "Scanner_Interface"
'***************************************************************************
'Scanner Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 1/10/01
'Last updated: 15/June/12
'Last update: automatically populate the default filename with "Scanned Image " & today's date.
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
'***************************************************************************

Option Explicit

Private Declare Function TWAIN_AcquireToFilename Lib "EZTW32.dll" (ByVal hwndApp As Long, ByVal sFile As String) As Long
Private Declare Function TWAIN_SelectImageSource Lib "EZTW32.dll" (ByVal hwndApp As Long) As Long
Private Declare Function TWAIN_IsAvailable Lib "EZTW32.dll" () As Long

Public Function EnableScanner() As Boolean
    'Quick hack to let me load the dll from anywhere I want
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "EZTW32.dll")
    If TWAIN_IsAvailable() = 0 Then EnableScanner = False Else EnableScanner = True
    FreeLibrary hLib
End Function

Public Sub Twain32SelectScanner()
    If ScanEnabled = True Then
        TWAIN_SelectImageSource (FormMain.HWnd)
    Else
    'If the EZTW32.dll file doesn't exist...
        MsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " Scanner Interface Error"
        Message "Scanning disabled "
        Exit Sub
    End If
    Message "Scanner successfully enabled "
End Sub

Public Sub Twain32Scan()

    Message "Acquiring image..."
    If ScanEnabled = False Then
        'If the EZTW32.dll file doesn't exist...
        MsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " Scanner Interface Error"
        Message "Scanner/digital camera import disabled "
        Exit Sub
    End If

    'This form has a fairly extensive error handling routine
    On Error GoTo ScanError
    
    Dim ScannerCaptureFile As String, ScanCheck As Long
    'ScanCheck is used to store the return values of the EZTW32.dll scanner functions.  We start by setting it
    ' to an arbitrary value that only we know; if an error occurs and this value is still present, it means an
    ' error occurred outside of the EZTW32 library.
    ScanCheck = -5
    
    'A temporary file is required by the scanner; we will place it in the project folder, then delete it when finished
    ScannerCaptureFile = TempPath & "PDScanInterface.tmp"
    
    'This line uses the EZTW32.dll file to scan the image and send it to a temporary file
    ScanCheck = TWAIN_AcquireToFilename(FormMain.HWnd, ScannerCaptureFile)
    
    'If the image was successfully scanned, load it
    If ScanCheck = 0 Then
        
        'Because PreLoadImage requires a string array, create one to send it
        Dim sFile(0) As String
        sFile(0) = ScannerCaptureFile
        
        PreLoadImage sFile, False, "Scanned Image", "Scanned Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
        'Be polite and remove the temporary file acquired from the scanner
        Kill ScannerCaptureFile
        
        Message "Image acquired successfully "
        
        FormMain.SetFocus
    Else
        'If the scan was unsuccessful, let the user know what happened
        GoTo ScanError
    End If
    
    Exit Sub

'Something went wrong
ScanError:
    
    Dim scanErrMessage As String
    
    Select Case ScanCheck
        Case -5
            scanErrMessage = "Unknown error occurred.  Please make sure your scanner is turned on and ready for use."
        Case -4
            scanErrMessage = "Scan successful, but temporary file save failed.  Is it possible that your hard drive is full (or almost full)?"
        Case -3
            scanErrMessage = "Unable to acquire DIB lock.  Please make sure no other programs are accessing the scanner.  If the problem persists, reboot and try again."
        Case -2
            scanErrMessage = "Temporary file access error.  This can be caused when running on a system with limited access rights.  Please enable admin rights and try again."
        Case -1
            scanErrMessage = "Scan canceled at the user's request."
            Message "Scan canceled."
            Exit Sub
        Case Else
            scanErrMessage = "The scanner returned an error code that wasn't specified in the EZTW32.dll documentation (Error # " & ScanCheck & ").  Please visit http://www.eztwain.com for more information."
    End Select
    
    MsgBox scanErrMessage, vbCritical + vbOKOnly + vbApplicationModal, "Scan Canceled"

    Message "Scan canceled. "
End Sub
