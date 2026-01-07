VERSION 5.00
Begin VB.Form FormKaleidoscope 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Kaleidoscope"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
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
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5925
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
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
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   1080
      Left            =   6000
      TabIndex        =   3
      Top             =   4080
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   1905
      Caption         =   "options"
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3495
      Index           =   0
      Left            =   5880
      Top             =   360
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltMirrors 
         Height          =   705
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "number of mirrors"
         Min             =   1
         Max             =   16
         Value           =   6
         DefaultValue    =   6
      End
      Begin PhotoDemon.pdSlider sltAngle 
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "primary angle"
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.pdSlider sltXCenter 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.pdSlider sltYCenter 
         Height          =   405
         Left            =   3120
         TabIndex        =   9
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.pdLabel lblExplanation 
         Height          =   435
         Index           =   0
         Left            =   240
         Top             =   1170
         Width           =   5655
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "you can also set a center position by clicking the preview window"
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   5
         Left            =   120
         Top             =   240
         Width           =   5805
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "center position (x, y)"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3495
      Index           =   1
      Left            =   5880
      Top             =   360
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltAngle2 
         Height          =   705
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "secondary angle"
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.pdSlider sltRadius 
         Height          =   705
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "radius (percentage)"
         Min             =   1
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdButtonStrip btsQuality 
         Height          =   1080
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1905
         Caption         =   "render emphasis"
      End
   End
End
Attribute VB_Name = "FormKaleidoscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Kaleiodoscope" Distortion
'Copyright 2013-2026 by Tanner Helland
'Created: 14/January/13
'Last updated: 20/February/18
'Last update: large performance improvements
'
'This tool allows the user to apply a simulated kaleidoscope distort to the image.  A number of variables can be
' set as part of the transformation; simply playing with the sliders should give a good indication of how they
' all work.
'
'As of January '14, the user can now select any center point for the effect.
'
'Finally, the transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 14 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "kaleidoscope" effect to an image
Public Sub KaleidoscopeImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Peering at image through imaginary kaleidoscope..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim numMirrors As Long, primaryAngle As Double, secondaryAngle As Double, effectRadius As Double
    Dim useBilinear As Boolean, centerX As Double, centerY As Double
    
    With cParams
        numMirrors = .GetLong("mirrors", sltMirrors.Value)
        primaryAngle = .GetDouble("angle", sltAngle.Value)
        secondaryAngle = .GetDouble("secondaryangle", sltAngle2.Value)
        effectRadius = .GetDouble("radius", sltRadius.Value)
        useBilinear = .GetBool("quality", True)
        centerX = .GetDouble("x", 0.5)
        centerY = .GetDouble("y", 0.5)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As RGBQuad, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image,
    ' and we will use it as our source reference.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters pdeo_Clamp, useBilinear, curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
          
    'Kaleidoscoping requires some specialized variables
    
    'Convert the input angles to radians
    primaryAngle = primaryAngle * (PI / 180#)
    secondaryAngle = secondaryAngle * (PI / 180#)
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Additional kaleidoscope variables
    Dim theta As Double, sRadius As Double, tRadius As Double, sDistance As Double
    Const ONE_DIV_PI As Double = 1# / PI
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
            
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) * 0.5
    
    sRadius = sRadius * (effectRadius / 100#)
    
    fSupport.AliasTargetDIB srcDIB
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapRGBQuadArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Calculate distance
        sDistance = Sqr((nX * nX) + (nY * nY))
        
        'Calculate theta
        theta = PDMath.Atan2_Faster(nY, nX) - primaryAngle - secondaryAngle
        theta = ConvertTriangle((theta * ONE_DIV_PI) * numMirrors * 0.5)
        
        'Calculate remapped x and y values
        If (sRadius > 0#) Then
            tRadius = sRadius / Cos(theta)
            sDistance = tRadius * ConvertTriangle(sDistance / tRadius)
        Else
            tRadius = sDistance
        End If
        
        theta = theta + primaryAngle
        
        srcX = midX + sDistance * Cos(theta)
        srcY = midY + sDistance * Sin(theta)
        
        'Retrieve the source pixel at this position and render it into the destination buffer
        dstImageData(x) = fSupport.GetColorsFromSource(srcX, srcY, x, y)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    fSupport.UnaliasTargetDIB
    workingDIB.UnwrapRGBQuadArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

'Change the active options panel
Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    picContainer(buttonIndex).Visible = True
    picContainer(Abs(1 - buttonIndex)).Visible = False
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Kaleidoscope", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Populate the options selector
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    
    'Populate the quality selector
    btsQuality.AddItem "quality", 0
    btsQuality.AddItem "speed", 1
    btsQuality.ListIndex = 0
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltAngle2_Change()
    UpdatePreview
End Sub

Private Sub sltMirrors_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then KaleidoscopeImage GetLocalParamString(), True, pdFxPreview
End Sub

'Return a repeating triangle shape in the range [0, 1] with wavelength 1.
'
'For correct results you need to multiply the output by 2 (to achieve [0, 1] amplitude),
' but this effect actually looks a little better when halved!  (The kaleidoscope pattern
' looks really busy otherwise.)
Private Function ConvertTriangle(ByVal trInput As Double) As Double
    ConvertTriangle = (trInput - Int(trInput))
    If (ConvertTriangle >= 0.5) Then ConvertTriangle = 1# - ConvertTriangle
    'ConvertTriangle = ConvertTriangle * 2#
End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    
    cmdBar.SetPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub sltXCenter_Change()
    UpdatePreview
End Sub

Private Sub sltYCenter_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "mirrors", sltMirrors.Value
        .AddParam "angle", sltAngle.Value
        .AddParam "secondaryangle", sltAngle2.Value
        .AddParam "radius", sltRadius.Value
        .AddParam "quality", (btsQuality.ListIndex = 0)
        .AddParam "x", sltXCenter.Value
        .AddParam "y", sltYCenter.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
