VERSION 5.00
Begin VB.Form dialog_ExportBMP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " BMP export options"
   ClientHeight    =   5160
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
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   5520
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "unique colors"
   End
   Begin PhotoDemon.pdSlider sldColorCount 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3060
      Width           =   4575
      _ExtentX        =   8070
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
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "color depth"
   End
   Begin PhotoDemon.pdButtonStrip btsColorModel 
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1720
      Caption         =   "color model"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4665
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdCheckBox chkRLE 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3660
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Caption         =   "use RLE compression"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chk16555 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Caption         =   "use legacy 15-bit encoding (5-5-5)"
      Value           =   0
   End
   Begin PhotoDemon.pdButtonStrip btsDepthGrayscale 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdCheckBox chkColorCount 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Caption         =   "restrict palette size"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chkPremultiplyAlpha 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Caption         =   "premultiply alpha"
      Value           =   0
   End
   Begin PhotoDemon.pdButtonStrip btsDepthRGBA 
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "alpha encoding"
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
    
    btsColorModel.AddItem "Auto", 0
    btsColorModel.AddItem "Color + Transparency (RGBA)", 1
    btsColorModel.AddItem "Color only (RGB)", 2
    btsColorModel.AddItem "Grayscale", 3
    
    btsDepthRGBA.AddItem "ARGB (store alpha values)", 0
    btsDepthRGBA.AddItem "XRGB (erase alpha values)", 1
    
    btsDepthRGB.AddItem "16 million colors (24-bpp)", 0
    btsDepthRGB.AddItem "65,536 colors (16-bpp)", 1
    btsDepthRGB.AddItem "256 colors (8-bpp)", 2
    
    btsDepthGrayscale.AddItem "256 shades (8-bpp)", 0
    btsDepthGrayscale.AddItem "16 shades (4-bpp)", 1
    btsDepthGrayscale.AddItem "Monochrome (1-bpp)", 2
    
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
            btsDepthRGBA.Visible = False
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
            
        'RGBA
        Case 1
            btsDepthRGBA.Visible = True
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = True
        
        'RGB
        Case 2
            btsDepthRGBA.Visible = False
            btsDepthRGB.Visible = True
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
        
        'Grayscale
        Case 3
            btsDepthRGBA.Visible = False
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = True
            chkPremultiplyAlpha.Visible = False
    
    End Select
    
    EvaluateDepthRGBAVisibility
    EvaluateDepthRGBVisibility

End Sub

Private Sub EvaluateDepthRGBAVisibility()
    If (Not btsDepthRGBA.Visible) Then
        chkPremultiplyAlpha.Visible = False
    Else
        Select Case btsDepthRGBA.ListIndex
        
            'ARGB
            Case 0
                chkPremultiplyAlpha.Visible = True
            
            'XRGB
            Case 1
                chkPremultiplyAlpha.Visible = False
            
        End Select
    End If
End Sub

Private Sub EvaluateDepthRGBVisibility()
    If (Not btsDepthRGB.Visible) Then
        chk16555.Visible = False
        setGroupVisibility_IndexedColor False
    Else
        Select Case btsDepthRGB.ListIndex
        
            '24-bpp
            Case 0
                chk16555.Visible = False
                setGroupVisibility_IndexedColor False
            
            '16-bpp
            Case 1
                chk16555.Visible = True
                setGroupVisibility_IndexedColor False
            
            '8-bpp
            Case 2
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

Private Sub btsDepthRGBA_Click(ByVal buttonIndex As Long)
    EvaluateDepthRGBAVisibility
End Sub

Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()

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
        
        'RGB
        Case 2
            Select Case btsDepthRGB.ListIndex
                
                '24-bpp
                Case 0
                    outputDepth = "24"
                
                '16-bpp
                Case 1
                    outputDepth = "16"
                
                '8-bpp
                Case 2
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
    cParams.AddParam "BMPUseXRGB", CBool(btsDepthRGBA.ListIndex = 1)
    cParams.AddParam "BMPPremultiplyAlpha", CBool(chkPremultiplyAlpha.Value)
    
    m_OutputParamString = cParams.GetParamString
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
    sldColorCount.Value = 256
    btsDepthGrayscale.ListIndex = 0
    btsDepthRGB.ListIndex = 0
    btsDepthRGBA.ListIndex = 0
    btsColorModel.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
