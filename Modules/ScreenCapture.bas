Attribute VB_Name = "ScreenCapture"
'***************************************************************************
'Screen Capture Interface
'Copyright 1999-2026 by Tanner Helland
'Created: 12/June/99
'Last updated: 13/April/22
'Last update: replace lingering picture box with pdPictureBox
'
'Minimal code for capturing the screen.  Because this code has to work from XP through Win 11, it doesn't
' attempt anything especially fancy - but note that myriad workarounds are *still* required for quirks
' in various OS versions.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Capturing the cursor as part of a screen capture is non-trivial; extra APIs are required
Private Enum W32_CursorState
    CursorHidden = 0
    CursorShowing = 1
    CursorSuppressed = 2    'Win8+ only
End Enum

#If False Then
    Private Const CursorHidden = 0, CursorShowing = 1, CursorSuppressed = 2
#End If

Private Type W32_IconInfo
    fIconBool As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type W32_CursorInfo
    cbSize As Long
    wFlags As W32_CursorState
    hCursor As Long
    ptX As Long
    ptY As Long
End Type

'Various API calls required for screen capturing and cursor rendering
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetCursorInfo Lib "user32" (ByVal ptrToCursorInfo As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByVal ptrToIconInfo As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, ByVal nFlags As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Const PW_CLIENTONLY As Long = &H1
Private Const PW_RENDERFULLCONTENT As Long = &H2    'Win 8.1+ only

Private Type WindowPlacement
    wpLength As Long
    wpFlags As Long
    wpShowCmd As Long
    ptMinPositionX As Long
    ptMinPositionY As Long
    ptMaxPositionX As Long
    ptMaxPositionY As Long
    rcNormalPosition As RectL
End Type
 
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpWndPl As WindowPlacement) As Long

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
Public Sub CaptureScreen(ByRef screenCaptureParams As String)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
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
    VBHacks.SleepAPI 500
    
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
    
    'Set the picture of the form to equal its image
    Dim tmpFilename As String
    tmpFilename = UserPrefs.GetTempPath & "PhotoDemon Screen Capture.tmpdib"
    
    'Ask the DIB to write out its data to file in PD's internal temporary DIB format
    tmpDIB.WriteToFile tmpFilename, cf_Lz4
        
    'We are now done with the temporary DIB, so free it up
    tmpDIB.EraseDIB
    Set tmpDIB = Nothing
        
    'Once the capture is saved, load it up like any other bitmap
    Dim sTitle As String
    If captureFullDesktop Then sTitle = g_Language.TranslateMessage("Screen capture") Else sTitle = alternateWindowName
    
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
    PDDebug.LogAction "Preparing to capture screen using rect (" & screenLeft & ", " & screenTop & ")x(" & screenWidth & ", " & screenHeight & ")"
    
    'Copy the image directly from the screen's DC to the target DIB's DC
    Dim screenHwnd As Long, desktopDC As Long
    screenHwnd = GetDesktopWindow()
    desktopDC = GetDC(screenHwnd)
    GDI.BitBltWrapper dstDIB.GetDIBDC, 0, 0, screenWidth, screenHeight, desktopDC, screenLeft, screenTop, vbSrcCopy
    ReleaseDC screenHwnd, desktopDC
    
    'Enforce correct alpha on the result
    dstDIB.ForceNewAlpha 255
    
End Sub

'Use this function to return a subsection of the current desktop in DIB format.
' IMPORTANT NOTE: the source rect should be in *desktop coordinates*, which may not be zero-based on a multimonitor system.
Public Sub GetPartialDesktopAsDIB(ByRef dstDIB As pdDIB, ByRef srcRect As RectL, Optional ByVal includeCursor As Boolean = False, Optional ByVal includeClicks As Boolean = False)
    
    'Make sure the target DIB is the correct size
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    Dim capWidth As Long, capHeight As Long
    capWidth = srcRect.Right - srcRect.Left
    capHeight = srcRect.Bottom - srcRect.Top
    If (dstDIB.GetDIBWidth <> capWidth) Or (dstDIB.GetDIBHeight <> capHeight) Then dstDIB.CreateBlank capWidth, capHeight, 32, 0, 0
    
    'BitBlt the relevant portion of the screen directly from the screen DC to the specified DIB
    Dim screenHwnd As Long, desktopDC As Long
    screenHwnd = GetDesktopWindow()
    desktopDC = GetDC(screenHwnd)
    GDI.BitBltWrapper dstDIB.GetDIBDC, 0, 0, capWidth, capHeight, desktopDC, srcRect.Left, srcRect.Top, vbSrcCopy
    ReleaseDC screenHwnd, desktopDC
    
    'Enforce normal alpha on the result
    If (dstDIB.GetDIBColorDepth = 32) Then dstDIB.ForceNewAlpha 255

    'Retrieve cursor info from the system
    Dim ci As W32_CursorInfo
    ci.cbSize = LenB(ci)
    If (GetCursorInfo(VarPtr(ci)) <> 0) Then
        
        'If the caller wants mouse clicks included, handle those first
        ' (so that our little overlay appears *above* the screenshot but *beneath* the cursor)
        If includeClicks Then
            
            'Note that we query both left and right mouse buttons, because GetAsyncKeyState
            ' doesn't differentiate between these for left-handed mouse users and we don't
            ' want to query additional APIs for that kind of low-level data - so said another
            ' way, *either* button down gets an identical render.
            Const VK_LBUTTON As Long = &H1, VK_RBUTTON As Long = &H2
            If (IsVirtualKeyDown(VK_LBUTTON, True) Or IsVirtualKeyDown(VK_RBUTTON, True)) Then
                
                'pd2D handles rendering duties
                Dim cBrush As pd2DBrush
                Drawing2D.QuickCreateSolidBrush cBrush, RGB(255, 255, 0), 50!
                
                Dim cSurface As pd2DSurface
                Set cSurface = New pd2DSurface
                cSurface.WrapSurfaceAroundPDDIB dstDIB
                cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
                
                Dim cRadius As Long
                cRadius = Interface.FixDPI(15)
                PD2D.FillCircleI cSurface, cBrush, (ci.ptX - srcRect.Left), (ci.ptY - srcRect.Top), cRadius
                Set cSurface = Nothing
                
            End If
            
        End If
        
        'If the caller wants the cursor displayed in the capture, we have extra work to do
        If includeCursor Then
    
            'Ensure cursor is visible
            If (ci.wFlags = CursorShowing) Then
            
                'Cursor is visible; use the cursor handle to retrieve a matching icon struct.
                ' (We need this so we can determine cursor hotspots, which may not be [0, 0].)
                Const DI_NORMAL As Long = 3
                Dim ii As W32_IconInfo
                If (GetIconInfo(ci.hCursor, VarPtr(ii)) <> 0) Then
                
                    'Mask rendering for cursors can be surprisingly complicated, especially when
                    ' considering mono vs color (color masks are optional for mono and may not exist!)
                    ' Rather than use a manual renderer like MaskBlt, just use the system icon renderer
                    ' and let it sort out any messy details for us.  (Note that this solution has *NOT*
                    ' been tested on high-DPI displays.)
                    DrawIconEx dstDIB.GetDIBDC, (ci.ptX - srcRect.Left) - ii.xHotspot, (ci.ptY - srcRect.Top) - ii.yHotspot, ci.hCursor, 0&, 0&, 0&, 0&, DI_NORMAL
                    
                    'GetIconInfo allocates up to two bitmaps; these *must* be deleted before exiting!
                    ' (Note, however, that these *can* be 0, particularly for hbmColor - it is not
                    '  required for monochrome cursors like the traditional I-beam text cursor, so we
                    '  must check for existence before blindly deleting 'em.)
                    If (ii.hbmColor <> 0) Then DeleteObject ii.hbmColor
                    If (ii.hbmMask <> 0) Then DeleteObject ii.hbmMask
                
                'If GetIconInfo failed (this outcome is *not* expected), assume a hotspot of [0, 0]
                Else
                    DrawIconEx dstDIB.GetDIBDC, (ci.ptX - srcRect.Left), (ci.ptY - srcRect.Top), ci.hCursor, 0&, 0&, 0&, 0&, DI_NORMAL
                End If
            
            '/end cursor is visible check
            End If
        '/end caller wants cursor displayed check
        End If
        
    '/end cursor info retrieved successfully check
    End If
    
End Sub

'Copy the visual contents of any hWnd into a DIB; window chrome can be optionally included, if desired
Public Function GetHwndContentsAsDIB(ByRef dstDIB As pdDIB, ByVal targetHWnd As Long, Optional ByVal includeChrome As Boolean = True, Optional ByRef isWindowMinimized As Boolean = False) As Boolean

    'Vista+ defines window boundaries differently, so we have to use a special API to retrieve correct boundaries.
    ' (NOTE: this behavior is currently disabled, as it doesn't actually help that much, and it introduces some
    '        unwanted complexities under Win 10.  Further comments are given below.)
    'Dim hLib As Long
    'If OS.IsVistaOrLater Then hLib = LoadLibraryA("dwmapi.dll")
    
    'Start by retrieving the necessary dimensions from the target window
    Dim wpSuccess As Boolean, tmpWinPlacement As WindowPlacement
    tmpWinPlacement.wpLength = LenB(tmpWinPlacement)
    wpSuccess = (GetWindowPlacement(targetHWnd, tmpWinPlacement) <> 0)
    
    'See if the window is currently minimized; the caller may want to use this information to recognize that the capture
    ' isn't going to look right.
    If wpSuccess Then
        isWindowMinimized = (tmpWinPlacement.wpShowCmd = SW_SHOWMINIMIZED) Or (tmpWinPlacement.wpShowCmd = SW_MINIMIZE)
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
            GetWindowRect targetHWnd, targetRect
        'End If
        
    Else
        GetClientRect targetHWnd, targetRect
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
    
    GetHwndContentsAsDIB = (PrintWindow(targetHWnd, dstDIB.GetDIBDC, printFlags) <> 0)
    
    'DWM-rendered windows have the (bizarre) side-effect of alpha values being set to 0 in some regions of the image.
    ' To circumvent this, we forcibly set all alpha values to opaque, which makes the resulting image okay.
    If ((dstDIB.GetDIBColorDepth = 32) And GetHwndContentsAsDIB) Then dstDIB.ForceNewAlpha 255
    
    'Before returning, we now have to deal with an extremely obnoxious side-effect of PrintWindows' implementation.
    ' PrintWindow requires the client to respond to WM_PRINT messages; if a client doesn't choose to implement this,
    ' we're SOL.
    
    'To avoid this scenario, look for all-black images and resort to an alternate strategy if found.
    ' (Note that we can't rely on the return of PrintWindow for this; many applications return TRUE even if they
    '  don't process the message!)
    Dim captureIsAllBlack As Boolean: captureIsAllBlack = False
    captureIsAllBlack = DIBs.IsDIBSolidColor(dstDIB)
    
    If captureIsAllBlack Or (Not GetHwndContentsAsDIB) Then
        
        'New, alternate strategy: BitBlt directly from the source window.  This is not guaranteed to work
        ' (especially on Win 10+) but it's better than doing nothing.
        Dim tmpDC As Long
        tmpDC = GetWindowDC(targetHWnd)
        GDI.BitBltWrapper dstDIB.GetDIBDC, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, tmpDC, 0, 0, vbSrcCopy
        ReleaseDC targetHWnd, tmpDC
        
        'Deal with problematic alpha values
        dstDIB.ForceNewAlpha 255
        
        GetHwndContentsAsDIB = True
        
    End If
    
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
    
    'Only return visible windows
    If (IsWindowVisible(hWnd) <> 0) Then
    
        'Only return windows without parents (to exclude toolbars, etc)
        If (GetParent(hWnd) = 0&) Then
            
            'Only return windows with a size larger than 0
            Dim tmpRect As winRect
            GetWindowRect hWnd, tmpRect
            
            If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
            
                'Only return windows with a client size larger than 0
                GetClientRect hWnd, tmpRect
                
                If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
                    
                    'Retrieve the window's caption
                    Dim curWindowText As String, sizeOfName As Long
                    curWindowText = Space$(256)
                    sizeOfName = GetWindowText(hWnd, StrPtr(curWindowText), Len(curWindowText))
                    
                    'If window text was obtained, trim it and add this entry to the list
                    If (sizeOfName <> 0) Then
                    
                        curWindowText = Left$(curWindowText, sizeOfName)
                        
                        If (m_WindowNames Is Nothing) Then Set m_WindowNames = New pdStringStack
                        If (m_WindowHWnds Is Nothing) Then Set m_WindowHWnds = New pdStringStack
                        
                        'Perform one final check for protected or known-bad window types.
                        Dim okayToAdd As Boolean: okayToAdd = True
                        If (Not OS.IsVistaOrLater) Then
                            okayToAdd = Strings.StringsNotEqual(curWindowText, "Program Manager", True)
                        End If
                        
                        If OS.IsWin10OrLater Then
                            okayToAdd = Strings.StringsNotEqual(curWindowText, "Windows Shell Experience Host", True)
                        End If
                        
                        If okayToAdd Then okayToAdd = Strings.StringsNotEqual(Trim$(curWindowText), Trim$(g_Language.TranslateMessage("Screenshot options")), True)
                        
                        If okayToAdd Then
                            m_WindowNames.AddString curWindowText
                            m_WindowHWnds.AddString CStr(hWnd)
                        End If
                        
                    End If
                    
                End If
            End If
        End If
    End If
    
    'Return True, which instructs the function to continue enumerating window entries.
    EnumWindowsProc = 1&

End Function
