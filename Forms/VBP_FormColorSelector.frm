VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change color"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11535
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
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   28
      Text            =   "abcdef"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   25
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   23
      Top             =   720
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   21
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   20
      Top             =   3120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   18
      Top             =   2520
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   16
      Top             =   1920
      Width           =   3735
   End
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
      Left            =   8520
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
      Left            =   9990
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
   Begin PhotoDemon.textUpDown tudRGB 
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudRGB 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudRGB 
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHSV 
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   22
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   359
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHSV 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   24
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHSV 
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   26
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HTML / CSS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   5310
      TabIndex        =   27
      Top             =   3765
      Width           =   1110
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "blue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   5970
      TabIndex        =   14
      Top             =   3180
      Width           =   435
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "green:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   5835
      TabIndex        =   13
      Top             =   2580
      Width           =   570
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "red:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   6045
      TabIndex        =   12
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "value:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   5880
      TabIndex        =   11
      Top             =   1380
      Width           =   525
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "saturation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   5475
      TabIndex        =   10
      Top             =   780
      Width           =   930
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "hue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   6015
      TabIndex        =   9
      Top             =   180
      Width           =   390
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
      Width           =   11535
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
' selectors I've used (and there have been many!), I find GIMP's the most intuitive... strange, I know, considering
' what a mess the rest of their interface is.
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

'pdLayer for the hue box on the rihgt
Private hueBox As pdLayer

'Currently selected color, including RGB and HSL attributes
Private curColor As Long
Private curRed As Long, curGreen As Long, curBlue As Long
Private curHue As Double, curSaturation As Double, curValue As Double

'One layer for each of the individual color sample boxes
Private sRed As pdLayer, sGreen As pdLayer, sBlue As pdLayer
Private sHue As pdLayer, sSaturation As pdLayer, sValue As pdLayer

'Left/right/up arrows for the hue and color boxes; these are 7x13 (or 13x7) and loaded from the resource at run-time
Private leftSideArrow As pdLayer, rightSideArrow As pdLayer, upArrow As pdLayer

'Changing the various text boxes resyncs the dialog, unless this parameter is set.  (We use it to prevent
' infinite resyncs.)
Private suspendTextResync As Boolean

Private Enum colorCheckType
    ccRed = 0
    ccGreen = 1
    ccBlue = 2
    ccHue = 3
    ccSaturation = 4
    ccValue = 5
End Enum

#If False Then
    Private Const ccRed = 0, ccGreen = 1, ccBlue = 2, ccHue = 3, ccSaturation = 4, ccValue = 5
#End If

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected color (if any) is returned via this property
Public Property Get newColor() As Long
    newColor = newUserColor
End Property

'CANCEL button
Private Sub cmdCancel_Click()
    
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
    Set upArrow = New pdLayer
    
    loadResourceToLayer "CLR_ARROW_L", leftSideArrow
    loadResourceToLayer "CLR_ARROW_R", rightSideArrow
    loadResourceToLayer "CLR_ARROW_U", upArrow
        
    'Cache the currentColor parameter so we can access it elsewhere
    oldColor = initialColor
    
    'Render the old color to the screen.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color managed.
    Dim tmpLayer As New pdLayer
    tmpLayer.createBlank picOriginal.ScaleWidth, picOriginal.ScaleHeight, 24, oldColor
    tmpLayer.renderToPictureBox picOriginal
    
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
    
    'Manually assign a hand cursor to the main picture box.  Because cursors are assigned on a class basis, this will also assign
    ' a hand cursor to all other picture boxes on the form.  I'm okay with that.
    setHandCursor picColor
    
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
    
    'Because we want the hue box to be color-managed, we must create it as a DIB, then render it to the screen later
    Set hueBox = New pdLayer
    hueBox.createBlank picHue.ScaleWidth, picHue.ScaleHeight
    
    'Simple gradient-ish code implementation of drawing hue
    Dim y As Long
    For y = 0 To hueBox.getLayerHeight
    
        'Based on our x-position, gradient a value between -1 and 5
        hVal = y / hueBox.getLayerHeight
        
        'Generate a hue for this position (the 1 and 0.5 correspond to full saturation and half luminance, respectively)
        HSVtoRGB hVal, 1, 1, r, g, b
        
        'Draw the color
        drawLineToDC hueBox.getLayerDC, 0, y, picHue.ScaleWidth, y, RGB(r, g, b)
        
    Next y
    
    'Render the hue to the picture box, which will also activate color management
    hueBox.renderToPictureBox picHue
    
End Sub

'When *all* current color values are updated and valid, use this function to synchronize the interface to match
' their appearance.
Private Sub syncInterfaceToCurrentColor()
    
    Me.Picture = LoadPicture("")
    
    'Start by drawing the primary box (luminance/saturation) using the current values
    Set primaryBox = New pdLayer
    
    primaryBox.createBlank picColor.ScaleWidth, picColor.ScaleHeight
    
    Dim pImageData() As Byte
    Dim psa As SAFEARRAY2D
    prepSafeArray psa, primaryBox
    CopyMemory ByVal VarPtrArray(pImageData()), VarPtr(psa), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    Dim tmpR As Long, tmpG As Long, tmpB As Long
    Dim tmpSat As Double, tmpLum As Double
    
    Dim loopWidth As Long, loopHeight As Long
    loopWidth = primaryBox.getLayerWidth - 1
    loopHeight = primaryBox.getLayerHeight - 1
    
    For x = 0 To loopWidth
        QuickX = x * 3
    For y = 0 To loopHeight
    
        'The x-axis position determines value (0 -> 1)
        'The y-axis position determines saturation (1 -> 0)
        HSVtoRGB curHue, (loopHeight - y) / loopHeight, x / loopWidth, tmpR, tmpG, tmpB
        
        pImageData(QuickX + 2, y) = tmpR
        pImageData(QuickX + 1, y) = tmpG
        pImageData(QuickX, y) = tmpB
    
    Next y
    Next x
    
    'With our work complete, point the ImageData() array away from the DIBs and deallocate it
    CopyMemory ByVal VarPtrArray(pImageData), 0&, 4
    Erase pImageData
    
    'We now want to draw a circle around the point where the user's current color resides
    GDIPlusDrawCanvasCircle primaryBox.getLayerDC, curValue * loopWidth, (1 - curSaturation) * loopHeight, fixDPI(7), 192
        
    'Render the primary color box
    primaryBox.renderToPictureBox picColor
    
    'Render the current color box.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color managed.
    Dim tmpLayer As New pdLayer
    tmpLayer.createBlank picCurrent.ScaleWidth, picCurrent.ScaleHeight, 24, RGB(curRed, curGreen, curBlue)
    tmpLayer.renderToPictureBox picCurrent
    
    'Synchronize all text boxes to their current values
    redrawAllTextBoxes
    
    'Position the arrows along the hue box properly according to the current hue
    Dim hueY As Long
    hueY = picHue.Top + 1 + (curHue * picHue.ScaleHeight)
    
    leftSideArrow.alphaBlendToDC Me.hDC, , picHue.Left - leftSideArrow.getLayerWidth, hueY - (leftSideArrow.getLayerHeight \ 2)
    rightSideArrow.alphaBlendToDC Me.hDC, , picHue.Left + picHue.Width, hueY - (rightSideArrow.getLayerHeight \ 2)
    Me.Picture = Me.Image
    Me.Refresh
    
End Sub

'Use this sub to resync all text boxes to the current RGB/HSV values
Private Sub redrawAllTextBoxes()

    'We don't want the _Change events for the text boxes firing while we resync them, so we disable any resyncing in advance
    suspendTextResync = True
    
    'Start by matching up the text values themselves
    tudRGB(0) = curRed
    tudRGB(1) = curGreen
    tudRGB(2) = curBlue
    
    tudHSV(0) = curHue * 359
    tudHSV(1) = curSaturation * 100
    tudHSV(2) = curValue * 100
    
    'Next, prepare some universal values for the arrow image offsets
    Dim arrowOffset As Long
    arrowOffset = (upArrow.getLayerWidth \ 2) - 1
    
    Dim leftOffset As Long
    leftOffset = picSampleRGB(0).Left
    
    Dim widthCheck As Long
    widthCheck = picSampleRGB(0).ScaleWidth - 1
    
    'Next, redraw all marker arrows
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + ((curRed / 255) * widthCheck) - arrowOffset, picSampleRGB(0).Top + picSampleRGB(0).Height
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + ((curGreen / 255) * widthCheck) - arrowOffset, picSampleRGB(1).Top + picSampleRGB(1).Height
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + ((curBlue / 255) * widthCheck) - arrowOffset, picSampleRGB(2).Top + picSampleRGB(2).Height
    
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + (curHue * widthCheck) - arrowOffset, picSampleHSV(0).Top + picSampleHSV(0).Height
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + (curSaturation * widthCheck) - arrowOffset, picSampleHSV(1).Top + picSampleHSV(1).Height
    upArrow.alphaBlendToDC Me.hDC, , leftOffset + (curValue * widthCheck) - arrowOffset, picSampleHSV(2).Top + picSampleHSV(2).Height
    
    'Next, we need to prep all our color bar layers
    renderSampleLayer sRed, ccRed
    renderSampleLayer sGreen, ccGreen
    renderSampleLayer sBlue, ccBlue
    
    renderSampleLayer sHue, ccHue
    renderSampleLayer sSaturation, ccSaturation
    renderSampleLayer sValue, ccValue
    
    'Now we can render the bars to screen
    sRed.renderToPictureBox picSampleRGB(0)
    sGreen.renderToPictureBox picSampleRGB(1)
    sBlue.renderToPictureBox picSampleRGB(2)
    
    sHue.renderToPictureBox picSampleHSV(0)
    sSaturation.renderToPictureBox picSampleHSV(1)
    sValue.renderToPictureBox picSampleHSV(2)
    
    'Update the hex representation box
    txtHex = getHexStringFromRGB
    
    'Re-enable syncing
    suspendTextResync = False
    
End Sub

'When the user clicks the hue box (or moves with the mouse button down), this function is called.  It uses the y-value
' of the click to determine new image colors, then refreshes the interface.
Private Sub hueBoxClicked(ByVal clickY As Long)

    'Restrict mouse clicks to the picture box area
    If clickY < 0 Then clickY = 0
    If clickY > picHue.ScaleHeight Then clickY = picHue.ScaleHeight

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

'This sub handles the preparation of the individual color sample boxes (one each for R/G/B/H/S/V)
' (Because we want these boxes to be color-managed, we must create them as DIBs.)
Private Sub renderSampleLayer(ByRef dstLayer As pdLayer, ByVal layerColorType As colorCheckType)

    Set dstLayer = New pdLayer
    dstLayer.createBlank picSampleRGB(0).ScaleWidth, picSampleRGB(0).ScaleHeight
    
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double
    
    'Initialize each component to its default type; only one parameter will be changed per layerColorType
    r = curRed
    g = curGreen
    b = curBlue
    h = curHue
    s = curSaturation
    v = curValue
    
    Dim gradientValue As Double, gradientMax As Double
    gradientMax = dstLayer.getLayerWidth
    
    'Simple gradient-ish code implementation of drawing any individual color component
    Dim x As Long
    For x = 0 To dstLayer.getLayerWidth
    
        gradientValue = x / gradientMax
    
        'We handle RGB separately from HSV
        If layerColorType <= ccBlue Then
            
            Select Case layerColorType
            
                Case ccRed
                    r = gradientValue * 255
                    
                Case ccGreen
                    g = gradientValue * 255
                    
                Case Else
                    b = gradientValue * 255
                    
            End Select
            
        Else
        
            Select Case layerColorType
            
                Case ccHue
                    h = gradientValue
                
                Case ccSaturation
                    s = gradientValue
                
                Case ccValue
                    v = gradientValue
            
            End Select
            
            HSVtoRGB h, s, v, r, g, b
        
        End If
        
        'Draw the color
        drawLineToDC dstLayer.getLayerDC, x, 0, x, dstLayer.getLayerHeight, RGB(r, g, b)
        
    Next x
    
End Sub

Private Sub picSampleHSV_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    hsvBoxClicked Index, x
End Sub

Private Sub picSampleHSV_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then hsvBoxClicked Index, x
End Sub

'Whenever one of the HSV sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub hsvBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case (boxIndex + 3)
    
        Case ccHue
            curHue = (xPos / boxWidth)
        
        Case ccSaturation
            curSaturation = (xPos / boxWidth)
        
        Case ccValue
            curValue = (xPos / boxWidth)
    
    End Select
    
    'Recalculate RGB based on the new HSV values
    HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
    
    'Redraw the interface
    syncInterfaceToCurrentColor

End Sub

Private Sub picSampleRGB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    rgbBoxClicked Index, x
End Sub

Private Sub picSampleRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then rgbBoxClicked Index, x
End Sub

'Whenever one of the RGB sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub rgbBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case boxIndex
    
        Case ccRed
            curRed = (xPos / boxWidth) * 255
        
        Case ccGreen
            curGreen = (xPos / boxWidth) * 255
        
        Case ccBlue
            curBlue = (xPos / boxWidth) * 255
    
    End Select
    
    'Recalculate HSV based on the new RGB values
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Redraw the interface
    syncInterfaceToCurrentColor

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudHSV_Change(Index As Integer)

    If Not suspendTextResync Then

        Select Case (Index + 3)
        
            Case ccHue
                If tudHSV(Index).IsValid Then curHue = tudHSV(Index) / 359
            
            Case ccSaturation
                If tudHSV(Index).IsValid Then curSaturation = tudHSV(Index) / 100
            
            Case ccValue
                If tudHSV(Index).IsValid Then curValue = tudHSV(Index) / 100
        
        End Select
        
        'Recalculate RGB based on the new HSV values
        HSVtoRGB curHue, curSaturation, curValue, curRed, curGreen, curBlue
        
        'Redraw the interface
        syncInterfaceToCurrentColor
        
    End If

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudRGB_Change(Index As Integer)

    If Not suspendTextResync Then

        Select Case Index
        
            Case ccRed
                If tudRGB(Index).IsValid Then curRed = tudRGB(Index)
            
            Case ccGreen
                If tudRGB(Index).IsValid Then curGreen = tudRGB(Index)
        
            Case ccBlue
                If tudRGB(Index).IsValid Then curBlue = tudRGB(Index)
        
        End Select
        
        'Recalculate HSV values based on the new RGB values
        RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
        
        'Redraw the interface
        syncInterfaceToCurrentColor
        
    End If

End Sub

'Assuming the curRed/Green/Blue values are valid, this function can be used to retrieve a matching hex representation.
Private Function getHexStringFromRGB() As String
    getHexStringFromRGB = getHexFromByte(curRed) & getHexFromByte(curGreen) & getHexFromByte(curBlue)
End Function

'HTML hex requires each RGB entry to be two characters wide, but the VB Hex$ function won't add a leading 0.  We do this manually.
Private Function getHexFromByte(ByVal srcByte As Byte) As String
    If srcByte < 16 Then
        getHexFromByte = "0" & LCase(Hex$(srcByte))
    Else
        getHexFromByte = LCase(Hex$(srcByte))
    End If
End Function

Private Sub txtHex_Validate(Cancel As Boolean)

    'Before doing anything else, remove all invalid characters from the text box
    Dim validChars As String
    validChars = "0123456789abcdef"
    
    Dim curText As String
    curText = Trim$(txtHex)
    
    Dim newText As String
    newText = ""
    
    Dim curChar As String
    
    Dim i As Long
    For i = 1 To Len(curText)
        curChar = Mid$(curText, i, 1)
        If InStr(1, validChars, curChar, vbTextCompare) > 0 Then newText = newText & curChar
    Next i
        
    newText = LCase(newText)
    
    'newString now contains the contents of the text box, but limited to lowercase hex chars only.
    
    'Make sure the length is 1, 3, or 6.  Each case is handled specially.
    Select Case Len(newText)
    
        'One character is treated as a shade of gray; extend it to six characters.  (I don't know if this is actually
        ' valid CSS, but it doesn't hurt to support it... right?)
        Case 1
            newText = String$(6, newText)
        
        'Three characters is standard shorthand hex; expand each character as a pair
        Case 3
            newText = Left$(newText, 1) & Left$(newText, 1) & Mid$(newText, 2, 1) & Mid$(newText, 2, 1) & Right$(newText, 1) & Right$(newText, 1)
        
        'Six characters is already valid, so no need to screw with it further.
        Case 6
        
        Case Else
            'We can't handle this character string, so reset it
            newText = getHexStringFromRGB()
    
    End Select
    
    'Change the text box to match our properly formatted string
    txtHex = newText
    
    'Parse the string to calculate actual numeric values; we can use VB's Val() function for this!
    curRed = Val("&H" & Left$(newText, 2))
    curGreen = Val("&H" & Mid$(newText, 3, 2))
    curBlue = Val("&H" & Right$(newText, 2))
    
    'Calculate new HSV values to match
    RGBtoHSV curRed, curGreen, curBlue, curHue, curSaturation, curValue
    
    'Resync the interface to match the new value!
    syncInterfaceToCurrentColor

End Sub
