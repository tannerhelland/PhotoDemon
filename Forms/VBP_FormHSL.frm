VERSION 5.00
Begin VB.Form FormHSL 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Hue / Saturation / Lightness"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
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
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLuminance 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   5400
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "0"
      Top             =   5475
      Width           =   615
   End
   Begin VB.HScrollBar hsLuminance 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   -100
      TabIndex        =   12
      Top             =   5520
      Width           =   4935
   End
   Begin VB.TextBox txtSaturation 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   5400
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "0"
      Top             =   4635
      Width           =   615
   End
   Begin VB.HScrollBar hsSaturation 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   -100
      TabIndex        =   9
      Top             =   4680
      Width           =   4935
   End
   Begin VB.TextBox txtHue 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   5400
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "0"
      Top             =   3795
      Width           =   615
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsHue 
      Height          =   255
      Left            =   360
      Max             =   180
      Min             =   -180
      TabIndex        =   2
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   6360
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Label lblLuminance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "luminance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   1170
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "saturation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "FormHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'HSL Adjustment Form
'Copyright ©2011-2012 by Tanner Helland
'Created: 05/October/12
'Last updated: 05/October/12
'Last update: initial build
'
'Fairly simple and standard HSL adjustment form.  Layout and feature set derived from comparable tools
' in GIMP and Paint.NET.
'
'***************************************************************************

Option Explicit


'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    Me.Visible = False
    Process AdjustHSL, CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value)
    Unload Me
    
End Sub

'Colorize an image using a hue defined between -1 and 5
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub AdjustImageHSL(ByVal hModifier As Single, ByVal sModifier As Single, ByVal lModifier As Single, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Adjusting hue, saturation, and luminance values..."
    
    'Convert the modifiers to be on the same scale as the HSL translation routine
    
    hModifier = hModifier / 60
    sModifier = sModifier / 100
    lModifier = lModifier / 100
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Single, s As Single, l As Single
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Get the hue and saturation
        tRGBToHSL r, g, b, h, s, l
        
        'Apply the modifiers
        h = h + hModifier
        If h > 5 Then h = h - 6
        If h < -1 Then h = h + 6
        
        s = s + sModifier
        If s < 0 Then s = 0
        If s > 1 Then s = 1
        
        l = l + lModifier
        If l < 0 Then l = 0
        If l > 1 Then l = 1
        
        'Convert back to RGB using our artificial hue value
        tHSLToRGB h, s, l, r, g, b
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Activate()
    
    'Create a copy of the image on the preview window
    DrawPreviewImage picPreview
    
    'Display the previewed effect in the neighboring window
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'When the hue scroll bar is changed, redraw the preview
Private Sub hsHue_Change()
    copyToTextBoxI txtHue, hsHue.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub hsHue_Scroll()
    copyToTextBoxI txtHue, hsHue.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub hsLuminance_Change()
    copyToTextBoxI txtLuminance, hsLuminance.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub hsLuminance_Scroll()
    copyToTextBoxI txtLuminance, hsLuminance.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub hsSaturation_Change()
    copyToTextBoxI txtSaturation, hsSaturation.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub hsSaturation_Scroll()
    copyToTextBoxI txtSaturation, hsSaturation.Value
    AdjustImageHSL CSng(hsHue.Value), CSng(hsSaturation.Value), CSng(hsLuminance.Value), True, picEffect
End Sub

Private Sub txtHue_GotFocus()
    AutoSelectText txtHue
End Sub

Private Sub txtHue_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtHue, True
    If EntryValid(txtHue, hsHue.Min, hsHue.Max, False, False) Then hsHue.Value = Val(txtHue)
End Sub

Private Sub txtLuminance_GotFocus()
    AutoSelectText txtLuminance
End Sub

Private Sub txtLuminance_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtLuminance, True
    If EntryValid(txtLuminance, hsLuminance.Min, hsLuminance.Max, False, False) Then hsLuminance.Value = Val(txtLuminance)
End Sub

Private Sub txtSaturation_GotFocus()
    AutoSelectText txtSaturation
End Sub

Private Sub txtSaturation_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtSaturation, True
    If EntryValid(txtSaturation, hsSaturation.Min, hsSaturation.Max, False, False) Then hsSaturation.Value = Val(txtSaturation)
End Sub
