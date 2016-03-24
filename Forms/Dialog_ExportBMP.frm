VERSION 5.00
Begin VB.Form dialog_ExportBMP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " BMP export options"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdColorSelector clsBackground 
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   1980
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1720
      Caption         =   "background color"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   3720
      Top             =   4860
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "unique colors"
   End
   Begin PhotoDemon.pdSlider sldColorCount 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Min             =   2
      Max             =   256
      Value           =   256
      NotchPosition   =   2
      NotchValueCustom=   256
   End
   Begin PhotoDemon.pdButtonStrip btsDepthRGB 
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdButtonStrip btsColorModel 
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "color model"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdCheckBox chkRLE 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      Caption         =   "use RLE compression"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chk16555 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "use legacy 15-bit encoding (X1-R5-G5-B5)"
      Value           =   0
   End
   Begin PhotoDemon.pdButtonStrip btsDepthGrayscale 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdCheckBox chkColorCount 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      Caption         =   "restrict palette size"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chkPremultiplyAlpha 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1860
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "premultiply alpha"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chkFlipRows 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "flip row order (top-down)"
      Value           =   0
   End
End
Attribute VB_Name = "dialog_ExportBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bitmap export dialog
'Copyright 2012-2016 by Tanner Helland
'Created: 11/December/12
'Last updated: 16/March/16
'Last update: repurpose old color-depth dialog into a BMP-specific one
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'Output parameters (stored as an XML string)
Public m_OutputParamString As String

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "color + transparency", 1
    btsColorModel.AddItem "color only", 2
    btsColorModel.AddItem "grayscale", 3
    
    btsDepthRGB.AddItem "32-bpp XRGB (X8-R8-G8-B8)", 0
    btsDepthRGB.AddItem "24-bpp RGB (R8-G8-B8)", 1
    btsDepthRGB.AddItem "16-bpp (R5-G6-B5)", 2
    btsDepthRGB.AddItem "8-bpp (indexed)", 3
    
    btsDepthGrayscale.AddItem "8-bpp (256 shades)", 0
    btsDepthGrayscale.AddItem "4-bpp (16 shades)", 1
    btsDepthGrayscale.AddItem "1-bpp (monochrome)", 2
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsColorModel_Click(ByVal buttonIndex As Long)
    UpdateAllVisibility
End Sub

Private Sub UpdateAllVisibility()

    Select Case btsColorModel.ListIndex
    
        'Auto
        Case 0
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = False
            chkFlipRows.Visible = False
            
        'RGBA
        Case 1
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = True
            clsBackground.Visible = False
            chkFlipRows.Visible = True
        
        'RGB
        Case 2
            btsDepthRGB.Visible = True
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = True
            chkFlipRows.Visible = True
        
        'Grayscale
        Case 3
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = True
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = True
            chkFlipRows.Visible = True
    
    End Select
    
    EvaluateDepthRGBVisibility

End Sub

Private Sub EvaluateDepthRGBVisibility()
    If (Not btsDepthRGB.Visible) Then
        chk16555.Visible = False
        setGroupVisibility_IndexedColor False
    Else
        Select Case btsDepthRGB.ListIndex
        
            '32-bpp XRGB
            Case 0
                chk16555.Visible = False
                setGroupVisibility_IndexedColor False
                
            '24-bpp
            Case 1
                chk16555.Visible = False
                setGroupVisibility_IndexedColor False
            
            '16-bpp
            Case 2
                chk16555.Visible = True
                setGroupVisibility_IndexedColor False
            
            '8-bpp
            Case 3
                chk16555.Visible = False
                setGroupVisibility_IndexedColor True
        
        End Select
    End If
End Sub

Private Sub setGroupVisibility_IndexedColor(ByVal vState As Boolean)
    chkRLE.Visible = vState
    chkColorCount.Visible = vState
    sldColorCount.Visible = vState
    lblTitle(0).Visible = vState
End Sub

Private Sub btsDepthRGB_Click(ByVal buttonIndex As Long)
    EvaluateDepthRGBVisibility
End Sub

Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    m_OutputParamString = GetExportParamString
    userAnswer = vbOK
    Me.Hide
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    UpdateAllVisibility
End Sub

Private Sub cmdBar_ResetClick()
    chkPremultiplyAlpha.Value = vbUnchecked
    chk16555.Value = vbUnchecked
    chkColorCount.Value = vbUnchecked
    chkRLE = vbUnchecked
    chkFlipRows.Value = vbUnchecked
    sldColorCount.Value = 256
    btsDepthGrayscale.ListIndex = 0
    btsDepthRGB.ListIndex = 1
    btsColorModel.ListIndex = 0
    clsBackground.Color = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    'Convert the color depth option buttons into a usable numeric value
    Dim outputDepth As String
    
    Select Case btsColorModel.ListIndex
        
        'Auto
        Case 0
            outputDepth = "Auto"
        
        'RGBA
        Case 1
            outputDepth = "32"
            cParams.AddParam "BMPUseXRGB", False
            cParams.AddParam "BMPPremultiplyAlpha", CBool(chkPremultiplyAlpha.Value)
        
        'RGB
        Case 2
            Select Case btsDepthRGB.ListIndex
                
                '32-bpp XRGB
                Case 0
                    outputDepth = "32"
                    cParams.AddParam "BMPUseXRGB", True
                    cParams.AddParam "BMPPremultiplyAlpha", False
                
                '24-bpp
                Case 1
                    outputDepth = "24"
                
                '16-bpp
                Case 2
                    outputDepth = "16"
                
                '8-bpp
                Case 3
                    outputDepth = "8"
                
            End Select
        
        'Grayscale
        Case 3
            Select Case btsDepthGrayscale.ListIndex
                
                '8-bpp
                Case 0
                    outputDepth = "8"
                
                '4-bpp
                Case 1
                    outputDepth = "4"
                
                '1-bpp
                Case 2
                    outputDepth = "1"
                
            End Select
    
    End Select
    
    cParams.AddParam "BMPColorDepth", outputDepth
    cParams.AddParam "BMPRLECompression", CBool(chkRLE.Value)
    cParams.AddParam "BMPForceGrayscale", CBool(btsColorModel.ListIndex = 3)
    cParams.AddParam "BMP16bpp555", CBool(chk16555.Value)
    If CBool(chkColorCount.Value) Then cParams.AddParam "BMPIndexedColorCount", sldColorCount.Value
    cParams.AddParam "BMPBackgroundColor", clsBackground.Color
    cParams.AddParam "BMPFlipRowOrder", CBool(chkFlipRows.Value)
    
    GetExportParamString = cParams.GetParamString
    
End Function

