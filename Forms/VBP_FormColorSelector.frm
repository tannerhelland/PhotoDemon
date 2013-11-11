VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change color"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   8
      Top             =   4560
      Width           =   3735
   End
   Begin VB.PictureBox picCurrent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   7
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   5430
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5550
      TabIndex        =   2
      Top             =   5430
      Width           =   1365
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "original:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   4650
      Width           =   885
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "current:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   4170
      Width           =   840
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5280
      Width           =   7095
   End
End
Attribute VB_Name = "dialog_ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Selection Dialog
'Copyright ©2012-2013 by Tanner Helland
'Created: 11/November/13
'Last updated: 11/November/13
'Last update: initial build
'
'Basic color selection dialog.  I've modeled this after the comparable color selector in GIMP; of all the color
' selector's I've used (and there have been many!), I find it is the most intuitive.
'
'Unlike other dialogs in the program, I wanted this one to be fully resizable.  A bit of extra code is required
' to accomplish this, but I believe it's worth the effort.  Dialog size is not currently cached, but it could
' be in the future.
'
'More features are certainly possible in the future, but for now, the dialog is pretty minimalist.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The original color when the dialog was first loaded
Private oldColor As Long

'The new color selected by the user, if any
Private newUserColor As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Private m_ToolTip As clsToolTip

'pdLayer for the primary color box (luminance/saturation) on the left
Private primaryBox As pdLayer

'Currently selected color, including RGB and HSL attributes
Private curColor As Long
Private curRed As Long, curGreen As Long, curBlue As Long
Private curHue As Double, curSaturation As Double, curValue As Double

'Left/right arrows for the hue box; these are 7x13 and loaded from the resource file at run-time
Private leftSideArrow As pdLayer, rightSideArrow As pdLayer

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected color (if any) is returned via this property
Public Property Get newColor() As Long
    newColor = newUserColor
End Property

'CANCEL button
Private Sub CmdCancel_Click()
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    newUserColor = RGB(curRed, curGreen, curBlue)
    
    userAnswer = vbOK
    Me.Hide
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub showDialog(ByVal initialColor As Long)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Load the left/right side hue box arrow images from the resource file
    Set leftSideArrow = New pdLayer
    Set rightSideArrow = New pdLayer
    
    loadResourceToLayer "CLR_ARROW_L", leftSideArrow
    loadResourceToLayer "CLR_ARROW_R", rightSideArrow
    
    'Cache the currentColor parameter so we can access it elsewhere
    oldColor = initialColor
    picOriginal.BackColor = oldColor
    
    'Sync all current color values to the initial color
    curColor = initialColor
    curRed = ExtractR(initialColor)
    curGreen = ExtractG(initialColor)
    curBlue = ExtractB(initialColor)
    
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Synchronize the interface to this new color
    syncInterfaceToCurrentColor
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
        
    Message "Waiting for user to select color..."
        
    'Render the vertical hue box
    drawHueBox
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the dialog
    showPDDialog vbModal, Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The hue box only needs to be drawn once, when the dialog is first created
Private Sub drawHueBox()
    
    Dim hVal As Double
    Dim r As Long, g As Long, b As Long
    
    'Simple gradient-ish code implementation of drawing hue
    Dim y As Long
    For y = 0 To picHue.ScaleHeight
    
        'Based on our x-position, gradient a value between -1 and 5
        hVal = y / picHue.ScaleHeight
        
        'Generate a hue for this position (the 1 and 0.5 correspond to full saturation and half luminance, respectively)
        HSVtoRGB hVal, 1, 1, r, g, b
        
        'Draw the color
        picHue.Line (0, y)-(picHue.ScaleWidth, y), RGB(r, g, b)
        
    Next y
    
    turnOnDefaultColorManagement picHue.hDC, picHue.hWnd
    picHue.Picture = picHue.Image
    
End Sub

'When *all* current color values are updated and valid, use this function to synchronize the interface to match
' their appearance.
Private Sub syncInterfaceToCurrentColor()

    'Turn on color management for all relevant picture boxes
    turnOnDefaultColorManagement picColor.hDC, picColor.hWnd
    
    turnOnDefaultColorManagement picCurrent.hDC, picCurrent.hWnd
    turnOnDefaultColorManagement picOriginal.hDC, picOriginal.hWnd

    'Start by drawing the primary box (luminance/saturation) using the current values
    Set primaryBox = New pdLayer
    
    primaryBox.createBlank picColor.ScaleWidth, picColor.ScaleHeight
    
    Dim pImageData() As Byte
    Dim pSA As SAFEARRAY2D
    prepSafeArray pSA, primaryBox
    CopyMemory ByVal VarPtrArray(pImageData()), VarPtr(pSA), 4
    
    Dim x As Long, y As Long, quickX As Long
    
    Dim tmpR As Long, tmpG As Long, tmpB As Long
    Dim tmpSat As Double, tmpLum As Double
    
    Dim loopWidth As Long, loopHeight As Long
    loopWidth = primaryBox.getLayerWidth - 1
    loopHeight = primaryBox.getLayerHeight - 1
    
    For x = 0 To loopWidth
        quickX = x * 3
    For y = 0 To loopHeight
    
        'The x-axis position determines value (0 -> 1)
        'The y-axis position determines saturation (1 -> 0)
        HSVtoRGB curHue, (loopHeight - y) / loopHeight, x / loopWidth, tmpR, tmpG, tmpB
        
        pImageData(quickX + 2, y) = tmpR
        pImageData(quickX + 1, y) = tmpG
        pImageData(quickX, y) = tmpB
    
    Next y
    Next x
    
    'With our work complete, point the ImageData() array away from the DIBs and deallocate it
    CopyMemory ByVal VarPtrArray(pImageData), 0&, 4
    Erase pImageData
    
    'We now want to draw a circle around the point where the user's current color resides
    GDIPlusDrawCanvasCircle primaryBox.getLayerDC, curValue * loopWidth, (1 - curSaturation) * loopHeight, fixDPI(7), 192
        
    'Render the primary color box
    BitBlt picColor.hDC, 0, 0, primaryBox.getLayerWidth, primaryBox.getLayerHeight, primaryBox.getLayerDC, 0, 0, vbSrcCopy
    picColor.Picture = picColor.Image
    picColor.Refresh
        
    'Position the arrows along the hue box properly according to the current hue
    Dim hueY As Long
    hueY = picHue.Top + 1 + (curHue * picHue.ScaleHeight)
    
    Me.Picture = LoadPicture("")
    leftSideArrow.alphaBlendToDC Me.hDC, , picHue.Left - leftSideArrow.getLayerWidth, hueY - (leftSideArrow.getLayerHeight \ 2)
    rightSideArrow.alphaBlendToDC Me.hDC, , picHue.Left + picHue.Width, hueY - (rightSideArrow.getLayerHeight \ 2)
    Me.Picture = Me.Image
    Me.Refresh
    
    'Synchronize the "current color" picture box with the current color
    picCurrent.BackColor = RGB(curRed, curGreen, curBlue)

End Sub

'When the user clicks the hue box (or moves with the mouse button down), this function is called.  It uses the y-value
' of the click to determine new image colors, then refreshes the interface.
Private Sub hueBoxClicked(ByVal clickY As Long)

    'Calculate a new hue using the mouse's y-position as our guide
    curHue = clickY / picHue.ScaleHeight
    trimHSV curHue
    
    'Rebuild our RGB variables to match
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw any necessary interface elements
    syncInterfaceToCurrentColor

End Sub

'When the user clicks the primary box (or moves with the mouse button down), this function is called.  It uses the coordinates
' of the click to determine new image colors, then refreshes the interface.
Private Sub primaryBoxClicked(ByVal clickX As Long, ByVal clickY As Long)

    'Calculate a new value using the mouse's x-position as our guide
    curValue = clickX / picColor.ScaleWidth
    trimHSV curValue
    
    'Calculate a new saturation using the mouse's y-position as our guide
    curSaturation = clickY / picColor.ScaleHeight
    trimHSV curSaturation
    curSaturation = 1 - curSaturation
    
    'Rebuild our RGB variables to match
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw any necessary interface elements
    syncInterfaceToCurrentColor

End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    primaryBoxClicked x, y
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then primaryBoxClicked x, y
End Sub

Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    hueBoxClicked y
End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then hueBoxClicked y
End Sub

Private Sub trimHSV(ByRef hsvValue As Double)
    If hsvValue > 1 Then hsvValue = 1
    If hsvValue < 0 Then hsvValue = 0
End Sub
