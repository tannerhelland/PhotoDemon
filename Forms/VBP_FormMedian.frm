VERSION 5.00
Begin VB.Form FormMedian 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Median Filter"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   200
      Value           =   5
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
   Begin PhotoDemon.sliderTextCombo sltPercent 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   50
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
   Begin VB.Label lblPercentile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "percentile:"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label lblRadius 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius:"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblIDEWarning 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormMedian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Median Filter Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 08/Feb/13
'Last updated: 23/August/13
'Last update: added a mode-tracking variable to help with the new command bar addition
'
'This is a heavily optimized median filter function.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' median filter of a large radius (> 20).
'
'An optional percentile option is available.  At minimum value, this performs identically to an erode (minimum) filter.
' Similarly, at max value it performs identically to a dilate (maximum) filter.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Because this tool can be used for multiple actions (median, dilate, erode), we need to track which mode is currently active.
' When the reset or randomize buttons are pressed, we will automatically adjust our behavior to match.
Private Enum MedianToolMode
    MEDIAN_DEFAULT = 0
    MEDIAN_DILATE = 1
    MEDIAN_ERODE = 2
End Enum
Private curMode As MedianToolMode

'Apply a median filter to the image (heavily optimized accumulation implementation!)
'Input: radius of the median (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyMedianFilter(ByVal mRadius As Long, ByVal mPercent As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then
        If mPercent = 1 Then
            Message "Applying erode (minimum rank) filter..."
        ElseIf mPercent = 100 Then
            Message "Applying dilate (maximum rank) filter..."
        Else
            Message "Applying median filter..."
        End If
    End If
            
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curLayerValues.previewModifier
        If mRadius < 1 Then mRadius = 1
    End If
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    CreateMedianLayer mRadius, mPercent, srcLayer, workingLayer, toPreview
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Median", , buildParams(sltRadius.Value, sltPercent.Value)
End Sub

'Because this dialog can be used for multiple tools, we need to clarify some behavior when resetting and randomizing
Private Sub cmdBar_RandomizeClick()

    Select Case curMode
    
        Case MEDIAN_DEFAULT
            
        Case MEDIAN_DILATE
            sltPercent.Value = 100
        
        Case MEDIAN_ERODE
            sltPercent.Value = 1
    
    End Select

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()

    Select Case curMode
    
        Case MEDIAN_DEFAULT
            sltPercent.Value = 50
            
        Case MEDIAN_DILATE
            sltPercent.Value = 100
        
        Case MEDIAN_ERODE
            sltPercent.Value = 1
    
    End Select
    
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING! This tool is very slow when used inside the IDE. Please compile for best results.")
        lblIDEWarning.Visible = True
    End If
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews while we get everything initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The median dialog is reused for several tools: minimum, median, maximum.
Public Sub showMedianDialog(ByVal initPercentage As Long)

    If initPercentage = 1 Then
        Me.Caption = g_Language.TranslateMessage("Erode (Minimum rank filter)")
        sltPercent.Value = 1
        sltPercent.Visible = False
        lblPercentile.Visible = False
        cmdBar.setToolName "Erode"
        curMode = MEDIAN_ERODE
    ElseIf initPercentage = 100 Then
        Me.Caption = g_Language.TranslateMessage("Dilate (Maximum rank filter)")
        sltPercent.Value = 100
        sltPercent.Visible = False
        lblPercentile.Visible = False
        cmdBar.setToolName "Dilate"
        curMode = MEDIAN_DILATE
    Else
        Me.Caption = g_Language.TranslateMessage("Median filter")
        sltPercent.Value = initPercentage
        sltPercent.Visible = True
        lblPercentile.Visible = True
        curMode = MEDIAN_DEFAULT
    End If
    
    Me.Show vbModal, FormMain

End Sub

Private Sub sltPercent_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyMedianFilter sltRadius.Value, sltPercent.Value, True, fxPreview
End Sub
