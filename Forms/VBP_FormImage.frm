VERSION 5.00
Begin VB.Form FormImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   Caption         =   "Image Window"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   13275
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   885
   Visible         =   0   'False
End
Attribute VB_Name = "FormImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'The Activate event (which is handled by subclassing in the pdWindowManager class) wraps this public ActivateWorkaround function.
' This function can be called externally when any activation-related event (including peripheral things like the Next/Previous
' Image menus) requires a change in focus between images windows.
Public Sub ActivateWorkaround()
    
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    
End Sub

'LOAD form
Private Sub Form_Load()
    
    
End Sub

'Track which mouse buttons are pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    
End Sub

'Track which mouse buttons are released
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    
End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    
End Sub


Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    

End Sub

'In VB6, _QueryUnload fires before _Unload. We check for unsaved images here.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    
    
End Sub

Private Sub Form_Resize()
    
    If pdImages(Me.Tag) Is Nothing Then Exit Sub
    
    'Redraw this form if certain criteria are met (image loaded, form visible, viewport adjustments allowed)
    If (pdImages(Me.Tag).Width > 0) And (pdImages(Me.Tag).Height > 0) And Me.Visible And (FormMain.WindowState <> vbMinimized) And (g_WindowManager.getClientWidth(Me.hWnd) > 0) Then
        
        'Additionally, do not attempt to draw the image until it has been marked as "loaded successfully"; otherwise it will
        ' attempt to draw mid-load, causing unsightly flickering.
        If pdImages(Me.Tag).loadedSuccessfully Then
        
            'New test as of 16 Oct '13 - do not redraw the viewport unless it is the active one.
            If g_CurrentImage = CLng(Me.Tag) Then PrepareViewport Me, "Form_Resize(" & Me.ScaleWidth & "," & Me.ScaleHeight & ")"
            
            'Reflow any image-window-specific chrome (status bar, rulers, etc)
            fixChromeLayout
            
        End If
        
    End If
    
    'The height of a newly created form is automatically set to 1. This is normally changed when the image is
    ' resized to fit on screen, but if an image is loaded into a maximized window, the height value will remain
    ' at 1. If the user ever un-maximized the window, it will leave a bare title bar behind, which looks
    ' terrible. Thus, let's check for a height of 1, and if found resize the form to a larger (arbitrary) value.
    'If (Me.WindowState = vbNormal) And (Me.ScaleHeight <= 1) Then
    '    Me.Height = 6000
    '    Me.Width = 8000
    'End If
    
    'Remember this window state in the relevant pdImages object
    pdImages(Me.Tag).WindowState = Me.WindowState
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
End Sub
