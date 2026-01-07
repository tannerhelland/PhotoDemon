VERSION 5.00
Begin VB.Form FormIgnite 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Ignite"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
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
   ScaleWidth      =   784
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "color intensity"
      Min             =   1
      SigDigits       =   1
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "flame height"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdSlider sltOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
End
Attribute VB_Name = "FormIgnite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Burn" Fire FX Form
'Copyright 2001-2026 by Tanner Helland
'Created: some time 2001
'Last updated: 03/August/17
'Last update: migrate to XML params, performance improvements
'
'This fun little tool is a product of my own creation.  It works as follows:
'
' 1) Analyze image edges and create a contour map
' 2) For each pixel in the image, blur it upward at a distance relative to its luminance.  Apply linear decay as
'     each pixel is faded upward.
' 3) Recolor the image from (2) according to the user's intensity value; higher intensity warms the colors more.
' 4) Composite the final result against the original image, using the SCREEN blend mode (so that black pixels
'     are ignored, and bright pixels are maintained).
'
'While a "fire" effect has existed in PD for many years, it didn't receive its own dialog until July '14.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance during previews, we reuse a single temporary copy of the preview image
Private m_edgeDIB As pdDIB

'Apply the "burn" fire effect filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBurn(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Lighting image on fire..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxIntensity As Double
    Dim fxRadius As Double, fxOpacity As Long
    
    With cParams
        fxIntensity = .GetDouble("intensity", sltIntensity.Value)
        fxRadius = .GetDouble("radius", sltRadius.Value)
        fxOpacity = .GetLong("opacity", sltOpacity.Value)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    'Radius is simply a proportion of the current image's height
    fxRadius = (fxRadius * 0.01) * curDIBValues.Height
    If (fxRadius < 1#) Then fxRadius = 1#
    
    Dim progMax As Long
    If (Not toPreview) Then
        progMax = curDIBValues.Width + curDIBValues.Height * 2
        SetProgBarMax progMax
    End If
    
    'First things first: start by analyzing image edges and generating a white-on-black contour map
    If (m_edgeDIB Is Nothing) Then Set m_edgeDIB = New pdDIB
    m_edgeDIB.CreateFromExistingDIB workingDIB
    Filters_Layers.CreateContourDIB True, workingDIB, m_edgeDIB, toPreview, progMax, 0
    
    'Next, we're going to do two things: blurring the flame upward, while also applying some decay
    ' to the flame.
    m_edgeDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x * 0.333333333333333
    Next x
    
    'From those grayscale values, calculate a corresponding flame height.  (Brighter pixels cause the flame to travel further.)
    Dim distLookUp(0 To 765) As Single
    For x = 0 To 765
        distLookUp(x) = CDbl(x / 765#) * fxRadius
    Next x
    
    Dim fDistance As Double, fTargetMin As Long, innerY As Long
    Dim inR As Byte, inG As Byte, inB As Byte
    Dim fadeVal As Double
    
    'Loop through each pixel in the image, applying flame decay as we go
    For y = initY To finalY
    For x = initX To finalX
        
        'Get the source pixel color values
        xStride = x * 4
        b = imageData(xStride, y)
        g = imageData(xStride + 1, y)
        r = imageData(xStride + 2, y)
        
        'Calculate a distance value using our precalculated look-up values.  Basically, this is the max distance
        ' a flame can travel, and it's directly tied to the pixel's luminance (brighter pixels travel further).
        fDistance = distLookUp(r + g + b)
        
        'Calculate an upper bound
        fTargetMin = y - fDistance
        If (fTargetMin < 0) Then
            fTargetMin = 0
            fDistance = y
        End If
        
        'Trace a path upward, blending this value with neighboring pixels as we go
        If (fDistance > 0#) Then
            
            fDistance = 1# / fDistance
            
            For innerY = fTargetMin To y
                
                inB = imageData(xStride, innerY)
                inG = imageData(xStride + 1, innerY)
                inR = imageData(xStride + 2, innerY)
                
                'Blend this pixel's value with the value at this pixel, using the distance traveled as our blend metric
                fadeVal = (innerY - fTargetMin) * fDistance
                
                imageData(xStride, innerY) = (b * fadeVal) + (1# - fadeVal) * inB
                imageData(xStride + 1, innerY) = (g * fadeVal) + (1# - fadeVal) * inG
                imageData(xStride + 2, innerY) = (r * fadeVal) + (1# - fadeVal) * inR
                
            Next innerY
        
        End If
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal finalX + y
            End If
        End If
    Next y
    
    'Loop through the contour map one final time, recolor pixels to flame-like warm colors
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        m_edgeDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Perform the fire conversion
        r = grayVal * fxIntensity
        If (r > 255) Then r = 255
        g = grayVal
        b = grayVal \ fxIntensity
        
        'Assign the new "fire" value to each color channel
        imageData(x) = b
        imageData(x + 1) = g
        imageData(x + 2) = r
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal finalX + finalY + y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    m_edgeDIB.UnwrapArrayFromDIB imageData
    
    'Apply premultiplication prior to compositing
    m_edgeDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    'A pdCompositor class will help us selectively blend the flame results back onto the main image
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, m_edgeDIB, BM_Screen, fxOpacity
    
    'If this is *not* a preview, free any intermediary data
    If (Not toPreview) Then Set m_edgeDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True

End Sub

Private Sub cmdBar_OKClick()
    Process "Ignite", , GetLocalParamString(), UNDO_Layer
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

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxBurn GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "intensity", sltIntensity.Value
        .AddParam "radius", sltRadius.Value
        .AddParam "opacity", sltOpacity.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

