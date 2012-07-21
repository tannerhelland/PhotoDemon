VERSION 5.00
Begin VB.Form FormColorize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Colorize Options"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4965
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
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsHue 
      Height          =   255
      Left            =   240
      Max             =   359
      Min             =   1
      MouseIcon       =   "VBP_FormColorize.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Value           =   180
      Width           =   4575
   End
   Begin VB.PictureBox picHueDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   465
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   6
      Top             =   3360
      Width           =   4125
   End
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      MouseIcon       =   "VBP_FormColorize.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4080
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      MouseIcon       =   "VBP_FormColorize.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2340
      Width           =   4575
   End
End
Attribute VB_Name = "FormColorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Colorize Form
'Copyright ©2006-2012 by Tanner Helland
'Created: 12/January/07
'Last updated: 19/June/12
'Last update: minor modifications and optimizations
'
'Fairly simple and standard routine - look in the Miscellaneous Filters module
' for the HSL transformation code
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
    Process Colorize, CSng((CSng(hsHue.Value) - 60) / 60)
    Unload Me
End Sub

'Colorize an image using a hue defined between -1 and 5 (force saturation to 0.5)
Public Sub ColorizeImage(ByVal hToUse As Single)
    
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    
    Message "Colorizing image..."
    
    GetImageData
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        
        'Convert back to RGB using our artificial hue value
        tHSLToRGB hToUse, 0.5, LL, r, g, b
        
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
    
End Sub

Private Sub Form_Load()

    'This short routine is for drawing the picture box below the hue slider
    Dim hVal As Single
    Dim r As Long, g As Long, b As Long
    
    'Simple gradient-ish code implementation of drawing hue
    For x = 0 To picHueDemo.ScaleWidth
        'Based on our x-position, gradient a value between -1 and 5
        hVal = x / picHueDemo.ScaleWidth
        hVal = hVal * 360
        hVal = (hVal - 60) / 60
        
        'Generate a hue for this position (the one and 0.5 correspond to full saturation and half luminance, respectively)
        tHSLToRGB hVal, 1, 0.5, r, g, b
        
        'Draw the color
        picHueDemo.Line (x, 0)-(x, picHueDemo.ScaleHeight), RGB(r, g, b)
        
    Next x
    
    picHueDemo.Picture = picHueDemo.Image
    
    'Create a copy of the image on the preview window
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    'Actually do the effect
    DrawPreview
    
End Sub

'Preview a colorize effect
Private Sub DrawPreview()
    
    Dim r As Long, g As Long, b As Long
    
    GetPreviewData PicPreview
    Dim QuickVal As Long
    Dim hToUse As Single
    hToUse = (CSng(hsHue.Value) - 60) / 60
    Dim HH As Single, SS As Single, LL As Single
    
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial saturation value
        tHSLToRGB hToUse, 0.5, LL, r, g, b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
    Next x
        
    SetPreviewData PicEffect
End Sub

Private Sub hsHue_Change()
    DrawPreview
End Sub

Private Sub hsHue_Scroll()
    DrawPreview
End Sub
