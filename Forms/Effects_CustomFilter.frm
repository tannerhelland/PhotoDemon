VERSION 5.00
Begin VB.Form FormCustomFilter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Custom filter"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   DrawStyle       =   5  'Transparent
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdCheckBox chkNormalize 
      Height          =   330
      Left            =   6000
      TabIndex        =   26
      Top             =   3480
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   582
      Caption         =   "automatically normalize divisor and offset"
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   1
      Left            =   7320
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   2
      Left            =   8640
      TabIndex        =   3
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   3
      Left            =   9960
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   4
      Left            =   11280
      TabIndex        =   5
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   6
      Left            =   7320
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   7
      Left            =   8640
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   8
      Left            =   9960
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   9
      Left            =   11280
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   10
      Left            =   6000
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   11
      Left            =   7320
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   12
      Left            =   8640
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      DefaultValue    =   1
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
      Value           =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   13
      Left            =   9960
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   14
      Left            =   11280
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   15
      Left            =   6000
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   16
      Left            =   7320
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   17
      Left            =   8640
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   18
      Left            =   9960
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   19
      Left            =   11280
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   20
      Left            =   6000
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   21
      Left            =   7320
      TabIndex        =   22
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   22
      Left            =   8640
      TabIndex        =   23
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   23
      Left            =   9960
      TabIndex        =   24
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudF 
      Height          =   345
      Index           =   24
      Left            =   11280
      TabIndex        =   25
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Min             =   -1000
      Max             =   1000
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSpinner tudDivisor 
      Height          =   345
      Left            =   7560
      TabIndex        =   27
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      DefaultValue    =   1
      Min             =   1
      Max             =   1000
      SigDigits       =   1
      Value           =   1
   End
   Begin PhotoDemon.pdSpinner tudOffset 
      Height          =   345
      Left            =   9600
      TabIndex        =   28
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Min             =   -255
      Max             =   255
   End
   Begin PhotoDemon.pdLabel lblOffset 
      Height          =   285
      Left            =   9480
      Top             =   4080
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      Caption         =   "offset"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblDivisor 
      Height          =   285
      Left            =   7320
      Top             =   4095
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      Caption         =   "divisor"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblConvolution 
      Height          =   285
      Left            =   6000
      Top             =   600
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   503
      Caption         =   "convolution matrix"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Custom Filter Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 15/April/01
'Last updated: 31/July/17
'Last update: performance improvements, migrate to XML params
'
'This dialog allows the user to create custom convolution filters.  The actual processing of the convolution filter happens in
' a separate "ApplyConvolutionFilter" function; this dialog simply serves as a user-facing interface to that.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Normalizing automatically computes divisor and offset for the user
Private Sub chkNormalize_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Custom filter", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the filter is changed, update the preview to match
Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed Then
        
        'Disable additional previews (as we will be changing text box values)
        cmdBar.SetPreviewStatus False

        'If normalization has been requested, apply it before updating the preview
        tudDivisor.Enabled = Not chkNormalize.Value
        tudOffset.Enabled = Not chkNormalize.Value
        
        If chkNormalize.Value Then
        
            'Sum up the total of all filter boxes
            Dim filterSum As Double
            filterSum = 0
            
            Dim i As Long
            For i = 0 To 24
                filterSum = filterSum + TextSupport.CDblCustom(tudF(i))
            Next i
            
            'Generate automatic divisor and offset values based on the total.
            If (filterSum = 0) Then
                tudDivisor.Value = 1
                tudOffset.Value = 127
            ElseIf (filterSum > 0) Then
                tudDivisor.Value = filterSum
                tudOffset.Value = 0
            Else
                tudDivisor.Value = Abs(filterSum)
                tudOffset.Value = 255
            End If
        
        End If
            
        'Apply the preview
        Filters_Area.ApplyConvolutionFilter_XML GetLocalParamString(), True, pdFxPreview
    
        'Reenable previews
        cmdBar.SetPreviewStatus True
        
    End If
    
End Sub

Private Sub tudDivisor_Change()
    UpdatePreview
End Sub

Private Sub tudF_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub tudOffset_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'Stick all the current filter values into a parameter string, which can then be passed to the ApplyConvolutionFilter function
Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
    
        .AddParam "name", g_Language.TranslateMessage("custom")
        .AddParam "invert", False
        If (tudDivisor.Value = 0#) Then .AddParam "weight", 1# Else .AddParam "weight", tudDivisor.Value
        .AddParam "bias", tudOffset.Value
        
        'And finally, the convolution array itself.  This is just a pipe-delimited string with a 5x5 array
        ' of weights.
        Dim tmpString As String, i As Long
        For i = 0 To 24
            tmpString = tmpString & Trim$(Str$(tudF(i).Value))
            If (i < 24) Then tmpString = tmpString & "|"
        Next i
        
        .AddParam "matrix", tmpString
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

