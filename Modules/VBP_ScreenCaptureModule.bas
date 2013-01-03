Attribute VB_Name = "Screen_Capture"
'***************************************************************************
'Screen Capture Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/June/99
'Last updated: 04/September/12
'Last update: use the Sleep API call to prevent the capture message box from being caught in the capture.
'
'Description: this module captures the screen.  The options are fairly minimal - it only captures
'             the entire screen, but it does give the user the option to minimize the form first.
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
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


'Simple routine for capturing the screen and loading it as an image
Public Sub CaptureScreen()
    
    Message "Waiting for capture..."
    
    'See if we should minimize the form before the capture
    Dim CaptureMethod As Long
    CaptureMethod = MsgBox("Would you like to minimize " & PROGRAMNAME & " before capturing the screen?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "Screen Capture")
    
    'Check for cancel
    If CaptureMethod = vbCancel Then
        Message "Screen capture canceled. "
        Exit Sub
    End If
    
    'If the user wants us to minimize the form, obey their orders
    If CaptureMethod = vbYes Then FormMain.WindowState = vbMinimized

    'The capture happens so quickly that the message box prompting the capture will be caught in the snapshot.  Sleep for 1/4 of a second
    ' to give the msgbox time to disappear
    Sleep 250
            
    'Get the window handle of the screen
    Dim scrHwnd As Long
    scrHwnd = GetDesktopWindow()
    
    'Use the GetDC call to generate a device context for the screen's hWnd
    Dim scrhDC As Long
    scrhDC = GetDC(scrHwnd)

    'Get the screen dimensions in pixels and set the picture box size to that
    Dim screenLeft As Long, screenTop As Long
    Dim screenWidth As Long, screenHeight As Long
    
    'UPDATE 12 November '12: use our new cMonitors object to detect VIRTUAL screen size.  This will capture all monitors
    ' on a multimonitor arrangement, not just the primary one.
    screenLeft = cMonitors.DesktopLeft
    screenTop = cMonitors.DesktopTop
    screenWidth = cMonitors.DesktopWidth
    screenHeight = cMonitors.DesktopHeight
    
    'screenWidth = Screen.Width \ Screen.TwipsPerPixelX
    'screenHeight = Screen.Height \ Screen.TwipsPerPixelY
    
    'Convert the hDC into the appropriate bitmap format
    CreateCompatibleBitmap scrhDC, screenWidth, screenHeight
    
    'BitBlt from the new bitmap-compatible hDC to a temporary layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createBlank screenWidth, screenHeight
    BitBlt tmpLayer.getLayerDC, 0, 0, screenWidth, screenHeight, scrhDC, screenLeft, screenTop, vbSrcCopy
    
    'Release the object and handle we generated for the capture
    ReleaseDC scrHwnd, scrhDC
    DeleteObject scrhDC
    
    'If we minimized the main window, now's the time to return it to normal size
    If CaptureMethod = vbYes Then
        FormMain.WindowState = vbNormal
        DoEvents
    End If
    
    'Set the picture of the form to equal its image
    Dim tmpFilename As String
    tmpFilename = userPreferences.getTempPath & PROGRAMNAME & " Screen Capture.tmp"
    
    'Ask the layer to write out its data to file in BMP format
    tmpLayer.writeToBitmapFile tmpFilename
        
    'We are now done with the temporary layer, so free it up
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
        
    'Once the capture is saved, load it up like any other bitmap
    ' NOTE: Because PreLoadImage requires an array of strings, create an array to send to it
    Dim sFile(0) As String
    sFile(0) = tmpFilename
        
    PreLoadImage sFile, False, "Screen Capture", "Screen capture (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
    
    'Erase the temp file
    If FileExist(tmpFilename) Then Kill tmpFilename
    
    Message "Screen capture complete."
    
End Sub
