VERSION 5.00
Begin VB.Form FormWater 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Underwater"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   776
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltScale 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "scale"
      Max             =   250
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
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
   End
   Begin PhotoDemon.pdSlider sltTurbulence 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "turbulence"
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdSlider sldColor 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "color shift"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   3540
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1296
      Caption         =   "random seed:"
   End
End
Attribute VB_Name = "FormWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Underwater" Effect
'Copyright 2013-2026 by Tanner Helland
'Created: 01/January/2001?
'Last updated: 18/October/17
'Last update: created dedicated UI, expose more options to the user
'
'PhotoDemon has always provided some sort of silly "underwater" effect.
'
'In 7.0, the filter was expanded to provide a full UI and user-controllable parameters.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplyWaterFX(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Submerging image in artificial water..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxScale As Double, fxTurbulence As Double, fxSeed As String, fxColor As Double
    
    With cParams
        fxScale = .GetDouble("scale", sltScale.Value)
        fxTurbulence = .GetDouble("turbulence", sltTurbulence.Value)
        fxSeed = .GetString("seed")
        fxColor = .GetDouble("color", 0.5) * 0.01
    End With
    
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_String fxSeed
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    If toPreview Then fxScale = fxScale * curDIBValues.previewModifier
    
    Dim minDimension As Single
    If (curDIBValues.Width < curDIBValues.Height) Then minDimension = curDIBValues.Width Else minDimension = curDIBValues.Height
    fxTurbulence = fxTurbulence * (minDimension * 0.025)
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If the source image is tiny, abandon ship to prevent OOB errors
    If (finalX < 1) Or (finalY < 1) Then
        EffectPrep.FinalizeImageData toPreview, dstPic
        Exit Sub
    End If
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'Allow direct access to pixels
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
            
    'Because interpolation may be used, it's necessary to keep pixel values within special ranges.
    ' (This spares us needing to check for OOB on the inner pixel loop.)
    Dim xLimit As Long, yLimit As Long
    xLimit = finalX - 1
    yLimit = finalY - 1
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
          
    'This wave transformation requires specialized variables
    Dim xWavelength As Double
    xWavelength = fxScale
    If (xWavelength = 0#) Then xWavelength = 100000# Else xWavelength = 1# / xWavelength
    
    Dim xAmplitude As Double
    xAmplitude = fxScale * 0.5
        
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Finally, a bunch of variables used in color calculation
    Dim srcR As Long, srcG As Long, srcB As Long, srcA As Long
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = (x \ 3)
    Next x
                 
    'Loop through each pixel in the image, converting values as we go
    Dim yResult As Double
    
    For y = initY To finalY
        yResult = Sin(y * xWavelength) * xAmplitude
    For x = initX To finalX
    
        xStride = x * 4
        
        'Calculate new source pixel locations.  Note that we deliberately do not modify y by any
        ' trigonemetric functions - is stable, save for the "turbulence" parameter.
        srcX = x + yResult
        srcY = y
        
        'If turbulence is active, apply it now
        If (fxTurbulence > 0#) Then
            srcX = srcX + (cRandom.GetRandomFloat_WH() - 0.5) * fxTurbulence
            srcY = srcY + (cRandom.GetRandomFloat_WH() - 0.5) * fxTurbulence
        End If
        
        'Make sure the source coordinates are in-bounds
        If (srcX < 0#) Then
            srcX = Abs(srcX)
            If (srcX > xLimit) Then srcX = srcX Mod finalX
        End If
        
        If (srcY < 0#) Then
            srcY = Abs(srcY)
            If (srcY > yLimit) Then srcY = srcY Mod finalY
        End If
        
        If (srcX > xLimit) Then
            srcX = Abs(xLimit - (srcX - xLimit))
            If (srcX > xLimit) Then srcX = srcX Mod finalX
        End If
        
        If (srcY > yLimit) Then
            srcY = Abs(yLimit - (srcY - yLimit))
            If (srcY > yLimit) Then srcY = srcY Mod finalY
        End If
        
        'Interpolate the source pixel for better results
        srcB = GetInterpolatedVal(srcX, srcY, srcImageData, 0, 4)
        srcG = GetInterpolatedVal(srcX, srcY, srcImageData, 1, 4)
        srcR = GetInterpolatedVal(srcX, srcY, srcImageData, 2, 4)
        srcA = GetInterpolatedVal(srcX, srcY, srcImageData, 3, 4)
            
        'Now, modify the colors to give a bluish-green tint to the image
        grayVal = gLookup(srcR + srcG + srcB)
        r = gray - srcG - srcB
        g = gray - r - srcB
        b = gray - r - g
        
        'Keep all values in range
        If (r < 0) Then r = 0 Else If (r > 255) Then r = 255
        If (g < 0) Then g = 0 Else If (g > 255) Then g = 255
        If (b < 0) Then b = 0 Else If (b > 255) Then b = 255
        
        'Fade the colors according to the user's fade setting
        r = Colors.BlendColors(srcR, r, fxColor)
        g = Colors.BlendColors(srcG, g, fxColor)
        b = Colors.BlendColors(srcB, b, fxColor)
        
        'Write the colors (and alpha, if necessary) out to the destination image's data
        dstImageData(xStride, y) = b
        dstImageData(xStride + 1, y) = g
        dstImageData(xStride + 2, y) = r
        dstImageData(xStride + 3, y) = srcA
            
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    srcDIB.UnwrapArrayFromDIB srcImageData
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Water", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub rndSeed_Change()
    UpdatePreview
End Sub

Private Sub sldColor_Change()
    UpdatePreview
End Sub

Private Sub sltScale_Change()
    UpdatePreview
End Sub

Private Sub sltTurbulence_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyWaterFX GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "scale", sltScale.Value
        .AddParam "turbulence", sltTurbulence.Value
        .AddParam "color", sldColor.Value
        .AddParam "seed", rndSeed.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
