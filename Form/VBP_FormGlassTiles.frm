VERSION 5.00
Begin VB.Form FormGlassTiles 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Glass Tiles"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4035
      Width           =   5700
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _extentx        =   21325
      _extenty        =   1323
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   720
      Width           =   5895
      _extentx        =   10398
      _extenty        =   873
      min             =   -45
      max             =   45
      value           =   45
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _extentx        =   9922
      _extenty        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltSize 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   1800
      Width           =   5895
      _extentx        =   10398
      _extenty        =   873
      min             =   2
      max             =   200
      value           =   40
   End
   Begin PhotoDemon.sliderTextCombo sltCurvature 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   2880
      Width           =   5895
      _extentx        =   10398
      _extenty        =   873
      min             =   -20
      max             =   20
      value           =   8
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   9
      Top             =   4980
      Width           =   1005
      _extentx        =   1773
      _extenty        =   635
      caption         =   "quality"
      value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   10
      Top             =   4980
      Width           =   975
      _extentx        =   1720
      _extenty        =   635
      caption         =   "speed"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      Index           =   5
      Left            =   6000
      TabIndex        =   12
      Top             =   3600
      Width           =   3315
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   11
      Top             =   4590
      Width           =   1845
   End
   Begin VB.Label lblLuminance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "curvature:"
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
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label lblSaturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "size:"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblHue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "angle:"
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
      TabIndex        =   1
      Top             =   360
      Width           =   660
   End
End
Attribute VB_Name = "FormGlassTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Glass Tiles Filter Dialog
'Copyright ©2014 by Audioglider
'Created: 23/May/14
'Last updated: 23/May/14
'Last update: Initial build
'
'An image distortion filter that divides the image into clear glass blocks.
'The curvature parameter generates a convex surface for positives values
'and a concave surface for negative values.
'
'***************************************************************************

Option Explicit

'Number of sample points to calculate
Private Const AA_SAMPLES As Long = 17


'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip


Public Sub GlassTiles(ByVal lSquareSize As Long, ByVal lCurvature As Long, ByVal lAngle As Long, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Generating glass tiles..."
    
    Dim i As Long
    Dim m_Sin As Double, m_Cos As Double
    Dim m_Scale As Double, m_Curvature As Double
    Dim m_aaPT(0 To AA_SAMPLES - 1) As POINTAPI
    
    'Convert angles to radians
    m_Sin = Sin(lAngle * (180 / PI))
    m_Cos = Cos(lAngle * (180 / PI))
    
    m_Scale = PI / lSquareSize
    
    If lCurvature = 0 Then lCurvature = 1
    m_Curvature = lCurvature * lCurvature / 10# * (Abs(lCurvature) / lCurvature)
    
    Dim j As Double, k As Double
    For i = 0 To AA_SAMPLES - 1
        j = (i * 4) / CDbl(AA_SAMPLES)
        k = i / CDbl(AA_SAMPLES)
        
        j = j - CLng(j)
        m_aaPT(i).x = m_Cos * j + m_Sin * k
        m_aaPT(i).y = m_Cos * k - m_Sin * j
    Next i
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    
    'Filter algorithm variables
    Dim hw As Double, hh As Double
    Dim mm As Long
    Dim xSample As Double, ySample As Double
    Dim u As Double, v As Double, s As Double, t As Double
    
    hw = finalX / 2
    hh = finalY / 2
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        j = x - hw
        k = y - hh
        
        newR = newG = newB = newA = 0
        For mm = 0 To AA_SAMPLES - 1
            u = j + m_aaPT(mm).x
            v = k - m_aaPT(mm).y
            
            s = (m_Cos * u) + (m_Sin * v)
            t = (-m_Sin * u) + (m_Cos * v)
            
            s = s + m_Curvature * Tan(s * m_Scale)
            t = t + m_Curvature * Tan(t * m_Scale)
            u = (m_Cos * s) - (m_Sin * t)
            v = (m_Sin * s) + (m_Cos * t)
            
            xSample = CLng(hw + u)
            ySample = CLng(hh + v)
            
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            fSupport.getColorsFromSource r, g, b, a, xSample, ySample, srcImageData
            
            'Add the retrieved values to our running average
            newR = newR + r
            newG = newG + g
            newB = newB + b
            If qvDepth = 4 Then newA = newA + a
            
        Next mm
        
        newR = newR / AA_SAMPLES
        newG = newG / AA_SAMPLES
        newB = newB / AA_SAMPLES
            
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        'If the image has an alpha channel, repeat the calculation there too
        If qvDepth = 4 Then
            newA = newA \ AA_SAMPLES
            dstImageData(QuickVal + 3, y) = newA
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
       'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmbEdges_Change()
    updatePreview
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltAngle.value = 45
    sltSize.value = 40
    sltCurvature.value = 8
    cmbEdges.ListIndex = EDGE_CLAMP
    OptInterpolate(1).value = True
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Disable previewing until the form has been fully initialized
    cmdBar.markPreviewStatus False
    
    popDistortEdgeBox cmbEdges, EDGE_CLAMP
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltCurvature_Change()
    updatePreview
End Sub

Private Sub sltSize_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then GlassTiles sltSize.value, sltCurvature.value, sltAngle.value, CLng(cmbEdges.ListIndex), OptInterpolate(0).value, True, fxPreview
End Sub
'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
