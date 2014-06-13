VERSION 5.00
Begin VB.Form FormLensFlare 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Lens flare"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
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
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltXCenter 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
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
   Begin PhotoDemon.sliderTextCombo sltYCenter 
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "center position (x, y)"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: you can also set a center position by clicking the preview window."
      ForeColor       =   &H00404040&
      Height          =   435
      Index           =   0
      Left            =   6120
      TabIndex        =   4
      Top             =   1050
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormLensFlare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Lens Flare Form
'Copyright ©2014 by Audioglider
'Created: 12/June/14
'Last updated: 12/June/14
'Last update: Initial build
'
'This filter generates a "true" lens flare, giving the
' impression that the sun hit the objective. You can
' relocate the flare using x,y coordinates.
' This is a simple model of how the lens flare works:
'
'        -----------------------------------------------
'        |               \ | / |                       |
'        |                \|/  |                       |
'        |               --0-- |                       |
'        |                /|\  |                       |
'        |               / | \ |                       |
'        |                    1|                       |
'        |---------------------C-----------------------|
'        |                     |\                      |
'        |                     |(2)                    |
'        |                     |  \___                 |
'        |                     |  (\  )                |
'        |                     | (  n  )               |
'        |                     |  (__\)                |
'        -----------------------------------------------
'
'       C : origin on the image (center)
'       0 : center of the 'sun' with bright lines come out of it
'       1 : First flare
'       2 : Second flare
'
'       n : n'th flare
'
'***************************************************************************

Option Explicit

Private Type fRGB
    r As Double
    g As Double
    b As Double
End Type

Private Type tFlare
    RGBColor As fRGB  '[0..1 range]
    fScale As Double
    ptX As Long
    ptY As Long
    wType As Long
End Type

Dim m_Flares() As tFlare
Dim m_numFlares As Long
Dim m_sColor As Double, m_sGlow As Double, m_sInner As Double
Dim m_sOuter As Double, m_sHalo As Double
Dim m_Color As fRGB, m_cGlow As fRGB, m_cInner As fRGB
Dim m_cOuter As fRGB, m_cHalo As fRGB

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Helper function to initialize a flare
Private Function SetFlare(ByVal wType As Long, ByVal fScale As Double, ByVal x As Long, ByVal y As Long, _
                          ByVal r As Double, ByVal g As Double, ByVal b As Double) As tFlare
    Dim ret As tFlare
    
    With ret
        .wType = wType
        .fScale = fScale
        .ptX = x
        .ptY = y
        .RGBColor.r = r
        .RGBColor.b = b
        .RGBColor.g = g
    End With
    SetFlare = ret
    
End Function

Private Sub initFlares(ByVal startX As Long, ByVal startY As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal matt As Long)

    Dim xh As Long, yh As Long, dx As Long, dy As Long
    
    xh = imgWidth / 2
    yh = imgHeight / 2
    dx = xh - startX
    dy = yh - startY
    
    m_numFlares = 19
    ReDim m_Flares(0 To m_numFlares - 1)
    
    'The array describes the position, size, color, and type flare material.
    m_Flares(0) = SetFlare(1, matt * 0.027, CLng(0.6699 * dx + xh), CLng(0.6699 * dy + yh), 0, 14 / 255, 113 / 255)
    m_Flares(1) = SetFlare(1, matt * 0.01, CLng(0.2692 * dx + xh), CLng(0.2692 * dy + yh), 90 / 255, 181 / 255, 142 / 255)
    m_Flares(2) = SetFlare(1, matt * 0.005, CLng(-0.0112 * dx + xh), CLng(-0.0112 * dy + yh), 56 / 255, 140 / 255, 106 / 255)
    m_Flares(3) = SetFlare(2, matt * 0.031, CLng(0.649 * dx + xh), CLng(0.649 * dy + yh), 9 / 255, 29 / 255, 19 / 255)
    m_Flares(4) = SetFlare(2, matt * 0.015, CLng(0.4696 * dx + xh), CLng(0.4696 * dy + yh), 24 / 255, 14 / 255, 0)
    m_Flares(5) = SetFlare(2, matt * 0.037, CLng(0.4087 * dx + xh), CLng(0.4087 * dy + yh), 24 / 255, 14 / 255, 0)
    m_Flares(6) = SetFlare(2, matt * 0.022, CLng(-0.2003 * dx + xh), CLng(-0.2003 * dy + yh), 42 / 255, 19 / 255, 0)
    m_Flares(7) = SetFlare(2, matt * 0.025, CLng(-0.4103 * dx + xh), CLng(-0.4103 * dy + yh), 0, 9 / 255, 17 / 255)
    m_Flares(8) = SetFlare(2, matt * 0.058, CLng(-0.4503 * dx + xh), CLng(-0.4503 * dy + yh), 0, 4 / 255, 10 / 255)
    m_Flares(9) = SetFlare(2, matt * 0.017, CLng(-0.5112 * dx + xh), CLng(-0.5112 * dy + yh), 5 / 255, 5 / 255, 14 / 255)
    m_Flares(10) = SetFlare(2, matt * 0.2, CLng(-1.496 * dx + xh), CLng(-1.496 * dy + yh), 9 / 255, 4 / 255, 0)
    m_Flares(11) = SetFlare(2, matt * 0.5, CLng(-1.496 * dx + xh), CLng(-1.496 * dy + yh), 9 / 255, 4 / 255, 0)
    m_Flares(12) = SetFlare(3, matt * 0.075, CLng(0.4487 * dx + xh), CLng(0.4487 * dy + yh), 34 / 255, 19 / 255, 0)
    m_Flares(13) = SetFlare(3, matt * 0.1, CLng(dx + xh), CLng(dy + yh), 14 / 255, 26 / 255, 0)
    m_Flares(14) = SetFlare(3, matt * 0.039, CLng(-1.301 * dx + xh), CLng(-1.301 * dy + yh), 10 / 255, 25 / 255, 13 / 255)
    m_Flares(15) = SetFlare(4, matt * 0.19, CLng(1.309 * dx + xh), CLng(1.309 * dy + yh), 9 / 255, 0, 17 / 255)
    m_Flares(16) = SetFlare(4, matt * 0.195, CLng(1.309 * dx + xh), CLng(1.309 * dy + yh), 9 / 255, 16 / 255, 5 / 255)
    m_Flares(17) = SetFlare(4, matt * 0.2, CLng(1.309 * dx + xh), CLng(1.309 * dy + yh), 17 / 255, 4 / 255, 0)
    m_Flares(18) = SetFlare(4, matt * 0.038, CLng(-1.301 * dx + xh), CLng(-1.301 * dy + yh), 17 / 255, 4 / 255, 0)
    
End Sub

Private Sub AdjustPixel(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal percent As Double, ByRef colPro As fRGB)
    
    srcR = CLng(srcR + (255 - srcR) * percent * colPro.r)
    srcG = CLng(srcG + (255 - srcG) * percent * colPro.g)
    srcB = CLng(srcB + (255 - srcB) * percent * colPro.b)

End Sub

Private Sub mColor(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sColor
    If percent > 0 Then
        percent = percent * percent
        AdjustPixel srcR, srcG, srcB, percent, m_Color
    End If
    
End Sub

'Glow portion of the main flare
Private Sub mGlow(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sGlow
    If percent > 0 Then
        percent = percent * percent
        AdjustPixel srcR, srcG, srcB, percent, m_cGlow
    End If
    
End Sub

Private Sub mInner(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sInner
    If percent > 0 Then
        percent = percent * percent
        AdjustPixel srcR, srcG, srcB, percent, m_cInner
    End If
    
End Sub

Private Sub mOuter(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sOuter
    If percent > 0 Then
        AdjustPixel srcR, srcG, srcB, percent, m_cOuter
    End If
    
End Sub

Private Sub mHalo(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = Abs((h - m_sHalo) / (m_sHalo * 0.07))
    If percent < 1 Then
        AdjustPixel srcR, srcG, srcB, 1 - percent, m_cHalo
    End If
    
End Sub

'Returns a fixed point hypotenuse (used to calculate the angle of the flares)
Private Function FHypot(ByVal x As Double, y As Double) As Double
    FHypot = Sqr(x * x + y * y)
End Function

Private Sub mRt1(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = 1 - FHypot(ref.ptX - x, ref.ptY - y) / ref.fScale
    If percent > 0 Then
        percent = percent * percent
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
    End If
    
End Sub

Private Sub mRt2(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = ref.fScale - FHypot(ref.ptX - x, ref.ptY - y)
    percent = percent / (ref.fScale * 0.15)
    If percent > 0 Then
        If percent > 1 Then
            percent = 1
        End If
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
    End If
    
End Sub

Private Sub mRt3(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = ref.fScale - FHypot(ref.ptX - x, ref.ptY - y)
    percent = percent / (ref.fScale * 0.12)
    If percent > 0 Then
        If percent > 1 Then
            percent = 1 - (percent * 0.12)
        End If
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
    End If
    
End Sub

Private Sub mRt4(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = FHypot(ref.ptX - x, ref.ptY - y) - ref.fScale
    percent = percent / (ref.fScale * 0.04)
    percent = Abs(percent)
    If percent < 1 Then
        AdjustPixel srcR, srcG, srcB, 1 - percent, ref.RGBColor
    End If
    
End Sub

Public Sub LensFlare(Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying lens flare..."
    
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
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    Dim imgWidth As Long, i As Long
    Dim hyp As Double
    
    'Calculate width of drawing area available and adjust size of lens flare artifacts
    imgWidth = finalX - initX
    m_sColor = imgWidth * 0.0375
    m_sGlow = imgWidth * 0.078125
    m_sInner = imgWidth * 0.1796875
    m_sOuter = imgWidth * 0.3359375
    m_sHalo = imgWidth * 0.084375
    
    'Setup our default colors for the flares
    m_Color.r = 239 / 255: m_Color.g = 239 / 255: m_Color.b = 239 / 255
    m_cGlow.r = 245 / 255: m_cGlow.g = 245 / 255: m_cGlow.b = 245 / 255
    m_cInner.r = 255 / 255: m_cInner.g = 38 / 255:  m_cInner.b = 43 / 255
    m_cOuter.r = 69 / 255:  m_cOuter.g = 59 / 255:  m_cOuter.b = 64 / 255
    m_cHalo.r = 80 / 255:   m_cHalo.g = 15 / 255:   m_cHalo.b = 4 / 255
    
    'Initialize array of lens flares
    initFlares midX, midY, imgWidth, finalY - initY, imgWidth
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        
        hyp = FHypot(x - midX, y - midY)
        
        mColor r, g, b, hyp    ' Create color
        mGlow r, g, b, hyp     ' Create glow
        mInner r, g, b, hyp    ' Create inner
        mOuter r, g, b, hyp    ' Create outer
        mHalo r, g, b, hyp     ' Create halo
        
        For i = 0 To m_numFlares - 1
            Select Case m_Flares(i).wType
                Case 1: mRt1 r, g, b, m_Flares(i), x, y
                Case 2: mRt2 r, g, b, m_Flares(i), x, y
                Case 3: mRt3 r, g, b, m_Flares(i), x, y
                Case 4: mRt4 r, g, b, m_Flares(i), x, y
            End Select
        Next i
        
        'Assign the new values to each color channel
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
   
End Sub

Private Sub cmdBar_OKClick()
    Process "Lens flare", , buildParams(sltXCenter.value, sltYCenter.value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub
Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then LensFlare sltXCenter.value, sltYCenter.value, True, fxPreview
End Sub

Private Sub fxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.markPreviewStatus False
    sltXCenter.value = xRatio
    sltYCenter.value = yRatio
    cmdBar.markPreviewStatus True
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
