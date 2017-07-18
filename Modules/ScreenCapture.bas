Attribute VB_Name = "ScreenCapture"
'***************************************************************************
'Screen Capture Interface
'Copyright 1999-2017 by Tanner Helland
'Created: 12/June/99
'Last updated: 27/June/14
'Last update: sanitize window titles before converting them to filenames; otherwise, subsequent Save/Save As functions may fail
'
'Description: this module captures the screen.  The options are fairly minimal - it only captures
'             the entire screen, but it does give the user the option to minimize the form first.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Various API calls required for screen capturing
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, ByVal nFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Const PW_CLIENTONLY As Long = &H1
Private Const PW_RENDERFULLCONTENT As Long = &H2    'Win 8.1+ only

'Vista+ only
Private Declare Function DwmGetWindowAttribute Lib "dwmapi" (ByVal targetHwnd As Long, ByVal dwAttribute As Long, ByVal ptrToRecipient As Long, ByVal sizeOfRecipient As Long) As Long

'Helper functions for retrieving various window parameters
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long

Private Type WindowPlacement
    wpLength As Long
    wpFlags As Long
    wpShowCmd As Long
    ptMinPositionX As Long
    ptMinPositionY As Long
    ptMaxPositionX As Long
    ptMaxPositionY As Long
    rcNormalPosition As RECTL
End Type
 
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As WindowPlacement) As Long

'Constant used to determine window owner.
Private Const GWL_HWNDPARENT As Long = (-8)

'Local string stacks used to store open window names, and open window hWnds
Private m_WindowNames As pdStringStack
Private m_WindowHWnds As pdStringStack

'ShowWindow is used to minimize and restore the PhotoDemon window, if requested.  Using VB's internal .WindowState
' command doesn't notify the window manager (I have no idea why) so this necessary to prevent parts of the toolbar
' client areas from disappearing upon restoration.
Private Const SW_SHOWMINIMIZED As Long = &H2
Private Const SW_MINIMIZE As Long = 6&
Private Const SW_RESTORE As Long = 9&
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As Long) As Long

'Simple routine for capturing the screen and loading it as an image
Public Sub CaptureScreen(ByVal screenCaptureParams As String)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString screenCaptureParams
    
    Dim captureFullDesktop As Boolean, minimizePD As Boolean, alternateWindowHwnd As Long, includeChrome As Boolean, alternateWindowName As String
    
    With cParams
        captureFullDesktop = .GetBool("wholescreen", True)
        minimizePD = .GetBool("minimizefirst", False)
        alternateWindowHwnd = .GetLong("targethwnd", 0&)
        includeChrome = .GetBool("chrome", True)
        alternateWindowName = .GetString("targetwindowname", vbNullString)
    End With
    
    Message "Capturing screen..."
    
    'If the user wants us to minimize the form, obey their orders
    If (captureFullDesktop And minimizePD) Then ShowWindow FormMain.hWnd, SW_MINIMIZE
    
    'The capture happens so quickly that the message box prompting the capture will be caught in the snapshot.  Sleep for 1/2 of a second
    ' to give the message box time to disappear
    Sleep 500
    
    'Use the getDesktopAsDIB function to copy the requested screen contents into a temporary DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    If captureFullDesktop Then
        GetDesktopAsDIB tmpDIB
    Else
        If Not GetHwndContentsAsDIB(tmpDIB, alternateWindowHwnd, includeChrome) Then
            Message "Could not retrieve program window - the program appears to have been unloaded."
            Exit Sub
        End If
    End If
    
    'If we minimized the main window, now's the time to return it to normal size
    If (captureFullDesktop And minimizePD) Then ShowWindow FormMain.hWnd, SW_RESTORE
        
    'TODO: confirm that the previous step is okay on XP.  Previously, we had to forcibly invoke a full refresh via
    ' the window manager, but since switching to the new, lightweight toolbox manager in v7.0, I haven't re-checked this.
    
    'Set the picture of the form to equal its image
    Dim tmpFilename As String
    tmpFilename = g_UserPreferences.GetTempPath & PROGRAMNAME & " Screen Capture.tmpdib"
    
    'Ask the DIB to write out its data to file in PD's internal temporary DIB format
    tmpDIB.WriteToFile tmpFilename, PD_CE_Lz4
        
    'We are now done with the temporary DIB, so free it up
    tmpDIB.EraseDIB
    Set tmpDIB = Nothing
        
    'Once the capture is saved, load it up like any other bitmap
    Dim sTitle As String
    If captureFullDesktop Then sTitle = g_Language.TranslateMessage("Screen Capture") Else sTitle = alternateWindowName
    
    'Sanitize the calculated string to remove any potentially invalid characters
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    Dim sTitlePlusDate As String
    sTitlePlusDate = cFile.MakeValidWindowsFilename(sTitle) & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
    
    LoadFileAsNewImage tmpFilename, sTitlePlusDate, False
    
    'Erase the temp file
    Files.FileDeleteIfExists tmpFilename
    
    Message "Screen capture complete."
    
End Sub

'Use this function to return a copy of the current desktop in DIB format
Public Sub GetDesktopAsDIB(ByRef dstDIB As pdDIB)

    'Use the g_Displays object to detect VIRTUAL screen size.  This will capture all monitors on a multimonitor arrangement,
    ' not just the primary one.
    Dim screenLeft As Long, screenTop As Long
    Dim screenWidth As Long, screenHeight As Long
    screenLeft = g_Displays.GetDesktopLeft
    screenTop = g_Displays.GetDesktopTop
    screenWidth = g_Displays.GetDesktopWidth
    screenHeight = g_Displays.GetDesktopHeight
    
    'Prepare the target DIB
    dstDIB.CreateBlank screenWidth, screenHeight, 32
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Preparing to capture screen using rect (" & screenLeft & ", " & screenTop & ")x(" & screenWidth & ", " & screenHeight & ")"
    #End If
    
    'Copy the image directly from the screen's DC to the target DIB's DC
    Dim screenHwnd As Long, desktopDC As Long
    screenHwnd = GetDesktopWindow()
    desktopDC = GetDC(screenHwnd)
    BitBlt dstDIB.GetDIBDC, 0, 0, screenWidth, screenHeight, desktopDC, screenLeft, screenTop, vbSrcCopy
    ReleaseDC screenHwnd, desktopDC
    
    'Enforce correct alpha on the result
    dstDIB.ForceNewAlpha 255
    
End Sub

'Use this function to return a subsection of the current desktop in DIB format.
' IMPORTANT NOTE: the source rect should be in *desktop coordinates*, which may not be zero-based on a multimonitor system.
Public Sub GetPartialDesktopAsDIB(ByRef dstDIB As pdDIB, ByRef srcRect As RECTL)
    
    'Make sure the target DIB is the correct size
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, 32
    
    'BitBlt the relevant portion of the screen directly from the screen DC to the specified DIB
    Dim screenHwnd As Long, desktopDC As Long
    screenHwnd = GetDesktopWindow()
    desktopDC = GetDC(screenHwnd)
    BitBlt dstDIB.GetDIBDC, 0, 0, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, desktopDC, srcRect.Left, srcRect.Top, vbSrcCopy
    ReleaseDC screenHwnd, desktopDC
    
    'Enforce normal alpha on the result
    dstDIB.ForceNewAlpha 255
    
End Sub

'Copy the visual contents of any hWnd into a DIB; window chrome can be optionally included, if desired
Public Function GetHwndContentsAsDIB(ByRef dstDIB As pdDIB, ByVal targetHwnd As Long, Optional ByVal includeChrome As Boolean = True, Optional ByRef isWindowMinimized As Boolean = False) As Boolean

    'Vista+ defines window boundaries differently, so we have to use a special API to retrieve correct boundaries.
    ' (NOTE: this behavior is currently disabled, as it doesn't actually help that much, and it introduces some
    '        unwanted complexities under Win 10.  Further comments are given below.)
    'Dim hLib As Long
    'If OS.IsVistaOrLater Then hLib = LoadLibraryA("dwmapi.dll")
    
    'Start by retrieving the necessary dimensions from the target window
    Dim wpSuccess As Boolean, tmpWinPlacement As WindowPlacement
    tmpWinPlacement.wpLength = LenB(tmpWinPlacement)
    wpSuccess = CBool(GetWindowPlacement(targetHwnd, tmpWinPlacement) <> 0)
    
    'See if the window is currently minimized; the caller may want to use this information to recognize that the capture
    ' isn't going to look right.
    If wpSuccess Then
        isWindowMinimized = CBool(tmpWinPlacement.wpShowCmd = SW_SHOWMINIMIZED) Or CBool(tmpWinPlacement.wpShowCmd = SW_MINIMIZE)
    Else
        isWindowMinimized = False
    End If
    
    Dim targetRect As winRect
    If includeChrome Then
        
        'On Vista+, window border dimensions are reported differently, due to the way Aero handles window coords.
        ' We can retrieve those boundaries using the code below, but it doesn't always play nicely with the way
        ' individual applications respond to PrintWindow, so it's a crapshoot as to which retrieval method is better.
        ' Short of applying some kind of AutoCrop to the final image (ugh), we run a lower risk of damage by simply
        ' using the old, backward-compatible GDI measurement.
        
        'If OS.IsVistaOrLater And (hLib <> 0) Then
        '    Const DWMWA_EXTENDED_FRAME_BOUNDS As Long = 9&
        '    DwmGetWindowAttribute targetHwnd, DWMWA_EXTENDED_FRAME_BOUNDS, VarPtr(targetRect), 16&
        '    FreeLibrary hLib
        'Else
            GetWindowRect targetHwnd, targetRect
        'End If
        
    Else
        GetClientRect targetHwnd, targetRect
    End If
    
    'Check to make sure the window hasn't been destroyed
    If (targetRect.x2 - targetRect.x1 <= 0) Or (targetRect.y2 - targetRect.y1 <= 0) Then
        GetHwndContentsAsDIB = False
        Exit Function
    End If
    
    'Prepare the DIB at the proper size
    If OS.IsVistaOrLater Then
        dstDIB.CreateBlank targetRect.x2 - targetRect.x1, targetRect.y2 - targetRect.y1, 32
    Else
        dstDIB.CreateBlank targetRect.x2 - targetRect.x1, targetRect.y2 - targetRect.y1, 24
    End If
    
    'Ask the window in question to paint itself into our DIB
    Dim printFlags As Long
    printFlags = 0&
    If (Not includeChrome) Then printFlags = printFlags Or PW_CLIENTONLY
    If OS.IsWin81OrLater Then printFlags = printFlags Or PW_RENDERFULLCONTENT
    
    GetHwndContentsAsDIB = CBool(PrintWindow(targetHwnd, dstDIB.GetDIBDC, printFlags) <> 0)
    
    'DWM-rendered windows have the (bizarre) side-effect of alpha values being set to 0 in some regions of the image.
    ' To circumvent this, we forcibly set all alpha values to opaque, which makes the resulting image okay.
    If ((dstDIB.GetDIBColorDepth = 32) And GetHwndContentsAsDIB) Then dstDIB.ForceNewAlpha 255
    
End Function

'After calling EnumWindowsProc, you can call this function to get a copy of the window title and hWnd string stacks.
' IMPORTANT NOTE: by design, this function will clear the local copies of window and hWnd names.
Public Sub GetAllWindowNamesAndHWnds(ByRef dstNameStack As pdStringStack, ByRef dstHWndStack As pdStringStack)
    
    If (dstNameStack Is Nothing) Then Set dstNameStack = New pdStringStack
    dstNameStack.CloneStack m_WindowNames
    Set m_WindowNames = Nothing
    
    If (dstHWndStack Is Nothing) Then Set dstHWndStack = New pdStringStack
    dstHWndStack.CloneStack m_WindowHWnds
    Set m_WindowHWnds = Nothing
    
End Sub

'The EnumWindows API call will call this function repeatedly until it exhausts the full list of open windows.
' We apply additional checks to the windows it returns to make sure there are no unwanted additions (hidden windows, etc).
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

    Static WindowText As String
    Static nRet As Long
    
    'Only return visible windows
    If IsWindowVisible(hWnd) Then
    
        'Only return windows without parents (to exclude toolbars, etc)
        If GetParent(hWnd) = 0& Then
            
            'Only return windows with a size larger than 0
            Dim tmpRect As winRect
            GetWindowRect hWnd, tmpRect
            
            If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
            
                'Only return windows with a client size larger than 0
                GetClientRect hWnd, tmpRect
                
                If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
                    
                    'Retrieve the window's caption
                    WindowText = Space$(256)
                    nRet = GetWindowText(hWnd, StrPtr(WindowText), Len(WindowText))
                    
                    'If window text was obtained, trim it and add this entry to the list
                    If (nRet <> 0) Then
                    
                        WindowText = Left$(WindowText, nRet)
                        
                        If (m_WindowNames Is Nothing) Then Set m_WindowNames = New pdStringStack
                        If (m_WindowHWnds Is Nothing) Then Set m_WindowHWnds = New pdStringStack
                        m_WindowNames.AddString WindowText
                        m_WindowHWnds.AddString CStr(hWnd)
                        
                    End If
                    
                End If
            End If
        End If
    End If
    
    'Return True, which instructs the function to continue enumerating window entries.
    EnumWindowsProc = 1

End Function
