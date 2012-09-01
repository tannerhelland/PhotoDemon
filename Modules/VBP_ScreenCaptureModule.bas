Attribute VB_Name = "Screen_Capture"
'***************************************************************************
'Screen Capture Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/99
'Last updated: 15/June/12
'Last update: append today's date to the default screen capture filename
'
'Description: this module captures the screen.  The options are fairly minimal - it only captures
'             the entire screen, but it does give the user the option to minimize the form first.
'
'***************************************************************************

Option Explicit

'Various API calls required for screen capturing
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long

'Simple routine for capturing the screen and loading it as an image
Public Sub CaptureScreen()
    
    Message "Waiting for capture..."
    
    'See if we should minimize the form before the capture
    Dim CaptureMethod As Long
    CaptureMethod = MsgBox("Would you like to minimize " & PROGRAMNAME & " before capturing the screen?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "Screen Capture")
    
    'Check for cancel
    If CaptureMethod = vbCancel Then
        Message "Screen capture canceled. "
        Unload FormMain.ActiveForm
        Exit Sub
    End If
    
    'If the user wants us to minimize the form, obey their orders
    If CaptureMethod = vbYes Then
        FormMain.WindowState = vbMinimized
    End If
    
    'Temporarily disable scrolling (to prevent strange scroll bar effects)
    FixScrolling = False
    
    'Create a new, blank form
    CreateNewImageForm True
    
    'Get the window handle of the screen
    Dim scrHwnd As Long
    scrHwnd = GetDesktopWindow()
    
    'Use the GetDC call to generate a device context for the screen's hWnd
    Dim scrhDC As Long
    scrhDC = GetDC(scrHwnd)

    'Get the screen dimensions in pixels and set the picture box size to that
    Dim screenWidth As Long, screenHeight As Long
    screenWidth = Screen.Width \ Screen.TwipsPerPixelX
    screenHeight = Screen.Height \ Screen.TwipsPerPixelY
    FormMain.ActiveForm.BackBuffer.Width = screenWidth + 2
    FormMain.ActiveForm.BackBuffer.Height = screenHeight + 2
    
    'Convert the hDC into the appropriate bitmap format
    CreateCompatibleBitmap scrhDC, screenWidth, screenHeight
    
    'BitBlt from the new bitmap-compatible hDC to the form
    BitBlt FormMain.ActiveForm.BackBuffer.hDC, 0, 0, screenWidth, screenHeight, scrhDC, 0, 0, vbSrcCopy
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    DoEvents  'Random fact - without DoEvents right here, this method fails.  Go figure.
    
    'Release the object and handle we generated for the capture
    ReleaseDC scrHwnd, scrhDC
    DeleteObject scrhDC
    
    'If we minimized the main window, now's the time to return it to normal size
    If CaptureMethod = vbYes Then
        FormMain.WindowState = vbNormal
        DoEvents
    End If
    
    'Set the picture of the form to equal its image
    Dim tmpFileName As String
    tmpFileName = TempPath & PROGRAMNAME & " Screen Capture.tmp"
    SavePicture FormMain.ActiveForm.BackBuffer.Picture, tmpFileName
    
    'Kill the temporary form
    Unload FormMain.ActiveForm
    DoEvents
    
    'Once the capture is saved, load it up like any other bitmap
    ' NOTE: Because PreLoadImage requires an array of strings, create an array to send to it
    Dim sFile(0) As String
    sFile(0) = tmpFileName
    
    FixScrolling = True
    
    PreLoadImage sFile, False, "Screen Capture", "Screen capture (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
    
    'Erase the temp file
    If FileExist(tmpFileName) Then Kill tmpFileName
    
    Message "Screen capture complete."
    
End Sub
