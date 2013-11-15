Attribute VB_Name = "Screen_Capture"
'***************************************************************************
'Screen Capture Interface
'Copyright ©1999-2013 by Tanner Helland
'Created: 12/June/99
'Last updated: 04/September/12
'Last update: use the Sleep API call to prevent the capture message box from being caught in the capture.
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
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, ByVal nFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Const PW_CLIENTONLY As Long = &H1
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Constant used to determine window owner.
Private Const GWL_HWNDPARENT = (-8)

'Listbox messages
Private Const LB_ADDSTRING = &H180
Private Const LB_SETITEMDATA = &H19A


'Simple routine for capturing the screen and loading it as an image
Public Sub CaptureScreen(ByVal captureFullDesktop As Boolean, ByVal minimizePD As Boolean, ByVal alternateWindowHwnd As Long, ByVal includeChrome As Boolean, Optional ByVal windowName As String)
    
    Message "Capturing screen..."
        
    'If the user wants us to minimize the form, obey their orders
    If captureFullDesktop And minimizePD Then FormMain.WindowState = vbMinimized

    'The capture happens so quickly that the message box prompting the capture will be caught in the snapshot.  Sleep for 1/4 of a second
    ' to give the message box time to disappear
    Sleep 250
    
    'Use the getDesktopAsLayer function to copy the requested screen contents into a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    
    If captureFullDesktop Then
        getDesktopAsLayer tmpLayer
    Else
        If Not getHwndContentsAsLayer(tmpLayer, alternateWindowHwnd, includeChrome) Then
            Message "Could not retrieve program window - the program appears to have been unloaded."
            Exit Sub
        End If
    End If
    
    'If we minimized the main window, now's the time to return it to normal size
    If captureFullDesktop And minimizePD Then FormMain.WindowState = vbNormal
    
    'Set the picture of the form to equal its image
    Dim tmpFilename As String
    tmpFilename = g_UserPreferences.getTempPath & PROGRAMNAME & " Screen Capture.tmp"
    
    'Ask the layer to write out its data to file in BMP format
    tmpLayer.writeToBitmapFile tmpFilename
        
    'We are now done with the temporary layer, so free it up
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
        
    'Once the capture is saved, load it up like any other bitmap
    ' NOTE: Because PreLoadImage requires an array of strings, create an array to send to it
    Dim sFile(0) As String
    sFile(0) = tmpFilename
    
    Dim sTitle As String
    If captureFullDesktop Then
        sTitle = g_Language.TranslateMessage("Screen Capture")
    Else
        sTitle = windowName
    End If
    
    Dim sTitlePlusDate As String
    sTitlePlusDate = sTitle & " (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
    
    PreLoadImage sFile, False, sTitle, sTitlePlusDate
    
    'Erase the temp file
    If FileExist(tmpFilename) Then Kill tmpFilename
    
    Message "Screen capture complete."
    
End Sub

'Use this function to return a copy of the current desktop in layer format
Public Sub getDesktopAsLayer(ByRef dstLayer As pdLayer)

    'Get the window handle of the screen
    Dim scrHwnd As Long
    scrHwnd = GetDesktopWindow()
    
    'Use the GetDC call to generate a device context for the screen's hWnd
    Dim scrhDC As Long
    scrhDC = GetDC(scrHwnd)

    'Get the screen dimensions in pixels and set the picture box size to that
    Dim screenLeft As Long, screenTop As Long
    Dim screenWidth As Long, screenHeight As Long
    
    'UPDATE 12 November '12: use our new g_cMonitors object to detect VIRTUAL screen size.  This will capture all monitors
    ' on a multimonitor arrangement, not just the primary one.
    screenLeft = g_cMonitors.DesktopLeft
    screenTop = g_cMonitors.DesktopTop
    screenWidth = g_cMonitors.DesktopWidth
    screenHeight = g_cMonitors.DesktopHeight
    
    'Convert the hDC into the appropriate bitmap format
    CreateCompatibleBitmap scrhDC, screenWidth, screenHeight
    
    'Copy the bitmap into the specified layer
    dstLayer.createBlank screenWidth, screenHeight
    BitBlt dstLayer.getLayerDC, 0, 0, screenWidth, screenHeight, scrhDC, screenLeft, screenTop, vbSrcCopy
    
    'Release the object and handle we generated for the capture, then exit
    ReleaseDC scrHwnd, scrhDC
    DeleteObject scrhDC

End Sub

'Copy the visual contents of any hWnd into a layer; window chrome can be optionally included, if desired
Public Function getHwndContentsAsLayer(ByRef dstLayer As pdLayer, ByVal targetHwnd As Long, Optional ByVal includeChrome As Boolean = True) As Boolean

    'Start by retrieving the necessary dimensions from the target window
    Dim targetRect As winRect
    
    If includeChrome Then
        GetWindowRect targetHwnd, targetRect
    Else
        GetClientRect targetHwnd, targetRect
    End If
    
    'Check to make sure the window hasn't been unloaded
    If (targetRect.x2 - targetRect.x1 = 0) Or (targetRect.y2 - targetRect.y1 = 0) Then
        getHwndContentsAsLayer = False
        Exit Function
    End If
    
    'Prepare the layer at the proper size
    dstLayer.createBlank targetRect.x2 - targetRect.x1, targetRect.y2 - targetRect.y1
    
    'Ask the window in question to paint itself into our layer
    If includeChrome Then
        PrintWindow targetHwnd, dstLayer.getLayerDC, 0
    Else
        PrintWindow targetHwnd, dstLayer.getLayerDC, PW_CLIENTONLY
    End If
    
    getHwndContentsAsLayer = True
    
End Function

'The EnumWindows API call will call this function repeatedly until it exhausts the full list of open windows.
' We apply additional checks to the windows it returns to make sure there are no unwanted additions (hidden windows, etc).
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

    Static WindowText As String
    Static nRet As Long
    
    'Only return visible windows
    If IsWindowVisible(hWnd) Then
    
        'Only return windows without parents (to exclude toolbars, etc)
        If GetParent(hWnd) = 0 Then
            
            'Only return windows with a size larger than 0
            Dim tmpRect As winRect
            GetWindowRect hWnd, tmpRect
            
            If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
            
                'Only return windows with a client size larger than 0
                GetClientRect hWnd, tmpRect
                
                If ((tmpRect.x2 - tmpRect.x1) > 0) And ((tmpRect.y2 - tmpRect.y1) > 0) Then
                    
                    'Retrieve the window's caption
                    WindowText = Space$(256)
                    nRet = GetWindowText(hWnd, WindowText, Len(WindowText))
                    
                    'If window text was obtained, trim it and add this entry to the list
                    If nRet Then
                    
                        WindowText = Left$(WindowText, nRet)
                        nRet = SendMessage(lParam, LB_ADDSTRING, 0, ByVal WindowText)
                        Call SendMessage(lParam, LB_SETITEMDATA, nRet, ByVal hWnd)
                    
                    End If
                    
                End If
            End If
        End If
    End If
    
    'Return True, which instructs the function to continue enumerating window entries.
    EnumWindowsProc = True

End Function
