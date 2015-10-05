VERSION 5.00
Begin VB.Form FormLensFlare 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Lens flare"
   ClientHeight    =   6750
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
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.buttonStrip btsOptions 
      Height          =   600
      Left            =   6240
      TabIndex        =   15
      Top             =   5160
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   1058
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   5880
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin PhotoDemon.sliderTextCombo sltXCenter 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.sliderTextCombo sltYCenter 
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.sliderTextCombo sltIntensity 
         Height          =   720
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "intensity"
         Min             =   0.01
         Max             =   3
         SigDigits       =   2
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.sliderTextCombo sltRadius 
         Height          =   720
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "radius"
         Min             =   1
         Max             =   200
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.sliderTextCombo sltHue 
         Height          =   720
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "hue"
         Max             =   359
         SliderTrackStyle=   4
      End
      Begin VB.Label lblExplanation 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: you can also set a position by clicking the preview window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5895
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "position (x, y)"
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
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   5880
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin PhotoDemon.sliderTextCombo sltIntensity 
         Height          =   720
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "secondary intensity"
         Min             =   0.01
         Max             =   3
         SigDigits       =   2
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.sliderTextCombo sltIntensity 
         Height          =   720
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "tertiary intensity"
         Min             =   0.01
         Max             =   3
         SigDigits       =   2
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.buttonStrip btsSyncIntensity 
         Height          =   600
         Left            =   330
         TabIndex        =   13
         Top             =   1140
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1058
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "synchronize intensity"
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
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2205
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "options"
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
      Index           =   6
      Left            =   6000
      TabIndex        =   14
      Top             =   4800
      Width           =   780
   End
End
Attribute VB_Name = "FormLensFlare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Lens Flare Form
'Copyright 2014 by Audioglider
'Created: 12/June/14
'Last updated: 12/June/14
'Last update: expand options exposed to the user.  Flare radius, intensity, and hue offset are now settable.
'              Advanced options also allow the user to specify intensity by flare group, if desired.
'              (By default, intensities are synced across all groups.)
'
'This filter generates a "true" lens flare, giving the impression that the sun hit the objective. The flare can
' be relocated by specifying custom (x, y) coordinates.
'
'Here's a simplified model of how lens flare works:
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
'Many thanks to expert contributor Audioglider for contributing this tool to PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'The main flare has its own collection of parameters, due to its specialized rendering
Private m_sColor As Double, m_sGlow As Double, m_sInner As Double
Private m_sOuter As Double, m_sHalo As Double
Private m_Color As fRGB, m_cGlow As fRGB, m_cInner As fRGB
Private m_cOuter As fRGB, m_cHalo As fRGB

'Secondary and tertiary flares are handled in a larger array
Private m_Flares() As tFlare
Private m_numFlares As Long

'To reduce the number of parameters passed mid-loop, some values are cached at module level
Private m_PrimaryIntensity As Double, m_SecondaryIntensity As Double, m_TertiaryIntensity As Double

'Helper function to initialize a flare
Private Function SetFlare(ByVal wType As Long, ByVal fScale As Double, ByVal x As Long, ByVal y As Long, _
                          ByVal r As Double, ByVal g As Double, ByVal b As Double) As tFlare
    
    With SetFlare
        .wType = wType
        .fScale = fScale
        .ptX = x
        .ptY = y
        .RGBColor.r = r
        .RGBColor.g = g
        .RGBColor.b = b
    End With
    
End Function

'Initialize all flare objects (besides the primary flare, which is handled specially).
' These values could be modified to produce different types of flares, but it's not a project for the faint of heart!
Private Sub initFlares(ByVal startX As Long, ByVal startY As Long, ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal matt As Long, ByVal hueOffset As Double)

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
    
    'Apply the requested hue offset, if any
    If hueOffset > 0 Then
    
        Dim i As Long
        
        For i = 0 To m_numFlares - 1
            rotateFlareHue m_Flares(i).RGBColor, hueOffset
        Next i
        
    End If
    
End Sub

'Given a flare's floating-point RGB triplet, and some hue adjustment value, rotate RGB values accordingly
Private Sub rotateFlareHue(ByRef flareRGB As fRGB, ByRef hueOffset As Double)

    Dim tmpH As Double, tmpS As Double, tmpV As Double
    
    With flareRGB
        
        fRGBtoHSV .r, .g, .b, tmpH, tmpS, tmpV
        
        'Rotate the returned hue by the specified amount
        tmpH = tmpH + hueOffset
        If tmpH > 1 Then tmpH = tmpH - 1
        
        fHSVtoRGB tmpH, tmpS, tmpV, .r, .g, .b
        
    End With

End Sub

'This central adjustment function is called by each flare handler, to apply the effect of a given flare point to a pixel
Private Sub AdjustPixel(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal percent As Double, ByRef colPro As fRGB)
    
    srcR = CLng(srcR + (255 - srcR) * percent * colPro.r)
    srcG = CLng(srcG + (255 - srcG) * percent * colPro.g)
    srcB = CLng(srcB + (255 - srcB) * percent * colPro.b)

End Sub

'Main color block of the primary flare
Private Sub mColor(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sColor
    
    percent = percent * m_PrimaryIntensity
    
    If percent > 0 Then
    
        If percent > 1 Then
            percent = 1
        Else
            percent = percent * percent
        End If
        
        AdjustPixel srcR, srcG, srcB, percent, m_Color
        
    End If
    
End Sub

'Glow portion of the primary flare
Private Sub mGlow(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sGlow
    
    percent = percent * m_PrimaryIntensity
    
    If percent > 0 Then
        
        If percent > 1 Then
            percent = 1
        Else
            percent = percent * percent
        End If
        
        AdjustPixel srcR, srcG, srcB, percent, m_cGlow
        
    End If
    
End Sub

'Inner glow of the primary flare
Private Sub mInner(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sInner
    
    percent = percent * m_PrimaryIntensity
    
    If percent > 0 Then
        
        If percent > 1 Then
            percent = 1
        Else
            percent = percent * percent
        End If
        
        AdjustPixel srcR, srcG, srcB, percent, m_cInner
        
    End If
    
End Sub

'Outer glow of the primary flare
Private Sub mOuter(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = 1 - h / m_sOuter
    
    percent = percent * m_PrimaryIntensity
    If percent > 1 Then percent = 1
    
    If percent > 0 Then AdjustPixel srcR, srcG, srcB, percent, m_cOuter
    
End Sub

'Thin halo of the primary flare
Private Sub mHalo(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByVal h As Double)
    
    Dim percent As Double
    percent = Abs((h - m_sHalo) / (m_sHalo * 0.07))
        
    If percent < 1 Then AdjustPixel srcR, srcG, srcB, (1 - percent) * m_PrimaryIntensity, m_cHalo
    
End Sub

'Returns a fixed point hypotenuse (used to calculate the angle of the flares, relative to the center of the image)
Private Function FHypot(ByVal x As Double, y As Double) As Double
    FHypot = Sqr(x * x + y * y)
End Function

'Secondary flare type one
Private Sub mRt1(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = 1 - FHypot(ref.ptX - x, ref.ptY - y) / ref.fScale
    
    If percent > 0 Then
        
        percent = percent * m_SecondaryIntensity
        percent = percent * percent
        
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
        
    End If
    
End Sub

'Secondary flare type two
Private Sub mRt2(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = ref.fScale - FHypot(ref.ptX - x, ref.ptY - y)
    percent = percent / (ref.fScale * 0.15)
    
    If percent > 0 Then
        percent = percent * m_SecondaryIntensity
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
    End If
    
End Sub

'Tertiary flare type one
Private Sub mRt3(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = ref.fScale - FHypot(ref.ptX - x, ref.ptY - y)
    percent = percent / (ref.fScale * 0.12)
    
    If percent > 0 Then
        
        If percent > 1 Then percent = 1 - (percent * 0.12)
        percent = percent * m_TertiaryIntensity
        
        AdjustPixel srcR, srcG, srcB, percent, ref.RGBColor
        
    End If
    
End Sub

'Tertiary flare type two
Private Sub mRt4(ByRef srcR As Long, ByRef srcG As Long, ByRef srcB As Long, ByRef ref As tFlare, ByVal x As Long, ByVal y As Long)
    
    Dim percent As Double
    percent = FHypot(ref.ptX - x, ref.ptY - y) - ref.fScale
    
    percent = percent / (ref.fScale * 0.04)
    percent = percent / m_TertiaryIntensity
    percent = Abs(percent)
    
    If percent < 1 Then AdjustPixel srcR, srcG, srcB, 1 - percent, ref.RGBColor
    
End Sub

'Apply a lens flare filter to an image
Public Sub LensFlare(Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal flareRadius As Double = 100, Optional ByVal primaryIntensity As Double = 1#, Optional ByVal secondaryIntensity As Double = 1#, Optional ByVal tertiaryIntensity As Double = 1#, Optional ByVal hueOffset As Double = 0#, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying lens flare..."
    
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
    
    Dim imgWidth As Long, imgHeight As Long, imgDiagonal As Long, i As Long
    Dim hyp As Double
    
    'Calculate width of drawing area available and adjust size of lens flare artifacts
    imgWidth = finalX - initX
    imgHeight = finalY - initY
    
    'Calculate a diagonal for the image, which will be used as the "base" radius
    imgDiagonal = Sqr(imgWidth * imgWidth + imgHeight * imgHeight)
    
    'The user's "Flare radius" input is also used to modify various inputs
    flareRadius = flareRadius / 100
    
    m_sColor = imgDiagonal * 0.0375 * flareRadius
    m_sGlow = imgDiagonal * 0.078125 * flareRadius
    m_sInner = imgDiagonal * 0.1796875 * flareRadius
    m_sOuter = imgDiagonal * 0.3359375 * flareRadius
    m_sHalo = imgDiagonal * 0.084375 * flareRadius
    
    m_PrimaryIntensity = primaryIntensity
    m_SecondaryIntensity = secondaryIntensity
    m_TertiaryIntensity = tertiaryIntensity
    
    'Setup our default colors for the flares
    m_Color.r = 239 / 255: m_Color.g = 239 / 255: m_Color.b = 239 / 255
    m_cGlow.r = 245 / 255: m_cGlow.g = 245 / 255: m_cGlow.b = 245 / 255
    m_cInner.r = 255 / 255: m_cInner.g = 38 / 255:  m_cInner.b = 43 / 255
    m_cOuter.r = 69 / 255:  m_cOuter.g = 59 / 255:  m_cOuter.b = 64 / 255
    m_cHalo.r = 80 / 255:   m_cHalo.g = 15 / 255:   m_cHalo.b = 4 / 255
    
    'Convert the hue modifier to the [0, 6] range
    hueOffset = hueOffset / 360
    
    'Rotate the hue of the primary flare object, if the user has requested it
    If hueOffset > 0 Then
        rotateFlareHue m_Color, hueOffset
        rotateFlareHue m_cGlow, hueOffset
        rotateFlareHue m_cInner, hueOffset
        rotateFlareHue m_cOuter, hueOffset
        rotateFlareHue m_cHalo, hueOffset
    End If
    
    'Initialize array of lens flare objects
    initFlares midX, midY, imgWidth, imgHeight, imgDiagonal * flareRadius, hueOffset
    
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
        
        'Iterate through each flare, rendering its effect (as relevant to this pixel)
        For i = 0 To m_numFlares - 1
            Select Case m_Flares(i).wType
                Case 1: mRt1 r, g, b, m_Flares(i), x, y
                Case 2: mRt2 r, g, b, m_Flares(i), x, y
                Case 3: mRt3 r, g, b, m_Flares(i), x, y
                Case 4: mRt4 r, g, b, m_Flares(i), x, y
            End Select
        Next i
        
        'The addition of variable intensity requires some additional failsafe checks, as the user's intensity values
        ' may cause HDR-style light blooming
        If r < 0 Then r = 0
        If g < 0 Then g = 0
        If b < 0 Then b = 0
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
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

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    picContainer(buttonIndex).Visible = True
    picContainer(1 - buttonIndex).Visible = False
End Sub

Private Sub btsSyncIntensity_Click(ByVal buttonIndex As Long)

    'Synchronize secondary and tertiary intensities to primary intensity
    If buttonIndex = 0 Then
    
        cmdBar.markPreviewStatus False
        sltIntensity(1).Value = sltIntensity(0).Value
        sltIntensity(2).Value = sltIntensity(0).Value
        cmdBar.markPreviewStatus True
        
        sltIntensity(1).Enabled = False
        sltIntensity(2).Enabled = False
    
    'Do NOT synchronize intensities
    Else
    
        sltIntensity(1).Enabled = True
        sltIntensity(2).Enabled = True
    
    End If
    
    updatePreview

End Sub

Private Sub cmdBar_OKClick()
    Process "Lens flare", , buildParams(sltXCenter, sltYCenter, sltRadius, sltIntensity(0), sltIntensity(1), sltIntensity(2), sltHue), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    sltXCenter.Value = 0.75
    sltYCenter.Value = 0.25
    
    sltIntensity(0).Value = 1
    sltIntensity(1).Value = 1
    sltIntensity(2).Value = 1
    
    sltRadius.Value = 100
    
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Set up the intensity sync option
    btsSyncIntensity.AddItem "yes", 0
    btsSyncIntensity.AddItem "no", 1
    btsSyncIntensity.ListIndex = 0
    btsSyncIntensity_Click 0
    
    'Set up the basic/advanced panels
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then LensFlare sltXCenter.Value, sltYCenter.Value, sltRadius, sltIntensity(0), sltIntensity(1), sltIntensity(2), sltHue, True, fxPreview
End Sub

Private Sub fxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.markPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.markPreviewStatus True
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltHue_Change()
    updatePreview
End Sub

Private Sub sltIntensity_Change(Index As Integer)
    
    'Synchronize secondary and tertiary options as necessary
    If (Index = 0) And (btsSyncIntensity.ListIndex = 0) Then
        
        'We disable previews before changing the other two intensity sliders; otherwise, their value changes
        ' will cause additional preview events to fire, harming performance.
        cmdBar.markPreviewStatus False
        sltIntensity(1).Value = sltIntensity(0).Value
        sltIntensity(2).Value = sltIntensity(0).Value
        cmdBar.markPreviewStatus True
        
    End If
    
    updatePreview
    
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub sltXCenter_Change()
    updatePreview
End Sub

Private Sub sltYCenter_Change()
    updatePreview
End Sub
