VERSION 5.00
Begin VB.Form FormSplitTone 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Split toning"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
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
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   5730
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5505
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9710
   End
   Begin PhotoDemon.pdSlider sltBalance 
      Height          =   705
      Left            =   6000
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "balance"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorMiddle=   16777215
   End
   Begin PhotoDemon.pdColorSelector cpHighlight 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "highlight color"
      curColor        =   16744192
   End
   Begin PhotoDemon.pdColorSelector cpShadow 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "shadow color"
      curColor        =   32767
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "toning strength"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormSplitTone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Split Toning Dialog
'Copyright 2015-2026 by Tanner Helland
'Created: 07/May/14
'Last updated: 20/July/17
'Last update: migrate to XML params
'
'Split toning (a digital relative of traditional Duotone printing) allows the user to apply two unique tones to an
' image: one for the highlights, and one for the shadows.  A balance slider controls the midpoint between highlights
' and shadows, while an optional strength parameter allows the user to blend the split tone results with the
' original image.  (This differs from traditional Duotone printing, where the image is reproduced using *only* the
' two inks specified.)
'
'For a comprehensive explanation of split-toning (and its historical relevance), see this article:
' http://www.alternativephotography.com/wp/toning/split-toning-history
'
' ... and for a good example of the effects that can be achieved with split toning, see this article:
' http://www.digitalcameraworld.com/2013/02/09/split-toning-in-photoshop-how-to-get-creative-with-your-black-and-white-conversions/
'
'PhotoDemon's version of this tool has gone through a lot of iterations.  The current incarnation tries to adhere to
' the traditional split-toning model, where the image is faded through gray at its specified midtone.  This limits
' the bulk of the coloring to the ends of the luminance spectrum, which reduces muddiness and draws the eye to the
' areas of greatest contrast in the image.  I think it's quite an excellent tool, and its results should be comparable
' to what you'd get from (much more expensive) professional software like Adobe Lightroom.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub SplitTone(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Split-toning image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim highlightColor As Long, shadowColor As Long, balance As Double, strength As Double
    
    With cParams
        highlightColor = .GetLong("highlightcolor", vbWhite)
        shadowColor = .GetLong("shadowcolor", vbBlack)
        balance = .GetDouble("balance", 0#)
        strength = .GetDouble("strength", 100#)
    End With
    
    'From the incoming colors, determine corresponding hue and saturation values
    Dim highlightHue As Double, highlightSaturation As Double, shadowHue As Double, shadowSaturation As Double
    Dim ignoreLuminance As Double
    PreciseRGBtoHSL Colors.ExtractRed(highlightColor) / 255#, Colors.ExtractGreen(highlightColor) / 255#, Colors.ExtractBlue(highlightColor) / 255#, highlightHue, highlightSaturation, ignoreLuminance
    PreciseRGBtoHSL Colors.ExtractRed(shadowColor) / 255#, Colors.ExtractGreen(shadowColor) / 255#, Colors.ExtractBlue(shadowColor) / 255#, shadowHue, shadowSaturation, ignoreLuminance
    
    'Convert balance mix value from an incoming range of [-100, 100] to a new range of [0,1].  We use this value
    ' to map colors between the shadow tone, neutral gray, and the highlight tone.
    Dim balGradient As Double, invBalGradient As Double
    invBalGradient = (balance + 100#) / 200#
    balGradient = 1# - invBalGradient
    
    'Prevent divide-by-zero errors, below
    If (invBalGradient <= 0.0000001) Then invBalGradient = 0.0000001
    If (balGradient <= 0.0000001) Then balGradient = 0.0000001
    
    'To avoid the need for many divisions on the inner loop, calculate inverse values now
    Dim multBalGradient As Double, multInvBalGradient As Double
    multInvBalGradient = 1# / invBalGradient
    multBalGradient = 1# / balGradient
    
    'Strength controls the ratio at which the split-toned pixels are merged with the original pixels.
    ' Convert it from a [0, 100] to [0, 1] scale.
    strength = strength * 0.01
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim v As Long, vFloat As Double
    
    Dim rHighlight As Double, gHighlight As Double, bHighlight As Double
    Dim rShadow As Double, gShadow As Double, bShadow As Double
    Dim thisGradient As Double
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
    
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Calculate HSL-compatible luminance
        v = Colors.GetHQLuminance(r, g, b)
        vFloat = v * ONE_DIV_255
        
        'Retrieve RGB conversions for the supplied highlight and shadow values, but retaining the pixel's current luminance (v)
        Colors.PreciseHSLtoRGB highlightHue, highlightSaturation, vFloat, rHighlight, gHighlight, bHighlight
        Colors.PreciseHSLtoRGB shadowHue, shadowSaturation, vFloat, rShadow, gShadow, bShadow
        
        'Highlight and shadow values are returned in the range [0, 1]; convert them to [0, 255] before continuing
        rHighlight = rHighlight * 255#
        rShadow = rShadow * 255#
        gHighlight = gHighlight * 255#
        gShadow = gShadow * 255#
        bHighlight = bHighlight * 255#
        bShadow = bShadow * 255#
        
        'We now have shadow and highlight colors for this pixel, already modified according to this pixel's luminance.
        
        'New strategy!  We don't want to color midtones, and midtones are defined according to the Balance parameter.
        ' So in a nutshell: if a pixel's luminance falls above the Balance param, fade it between gray and the highlight.
        ' If a pixel's luminance is below the Balance param, fade it between the shadow and gray.
        If (vFloat > balGradient) Then
        
            thisGradient = ((vFloat - balGradient) * multInvBalGradient)
            newR = rHighlight * thisGradient + v * (1# - thisGradient)
            newG = gHighlight * thisGradient + v * (1# - thisGradient)
            newB = bHighlight * thisGradient + v * (1# - thisGradient)
            
        Else
        
            thisGradient = (Abs(balGradient - vFloat) * multBalGradient)
            newR = rShadow * thisGradient + v * (1# - thisGradient)
            newG = gShadow * thisGradient + v * (1# - thisGradient)
            newB = bShadow * thisGradient + v * (1# - thisGradient)
            
        End If
                
        'Finally, apply the new RGB values to the image by blending them with their original color at the user's requested strength.
        imageData(x) = newB * strength + b * (1# - strength)
        imageData(x + 1) = newG * strength + g * (1# - strength)
        imageData(x + 2) = newR * strength + r * (1# - strength)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Split toning", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateBalanceSlider
    UpdatePreview
End Sub

'To help orient the user, slightly different reset values are used for this tool.
Private Sub cmdBar_ResetClick()
    cpHighlight.Color = RGB(255, 200, 150)
    cpShadow.Color = RGB(150, 200, 255)
End Sub

Private Sub cpHighlight_ColorChanged()
    UpdateBalanceSlider
    UpdatePreview
End Sub

Private Sub cpShadow_ColorChanged()
    UpdateBalanceSlider
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

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.SplitTone GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub sltBalance_Change()
    UpdatePreview
End Sub

'Redraw the balance slider gradient to match the currently selected split-toning values
Private Sub UpdateBalanceSlider()
    sltBalance.SetGradientColorsAndValueAtOnce cpShadow.Color, cpHighlight.Color, sltBalance.Value
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "highlightcolor", cpHighlight.Color
        .AddParam "shadowcolor", cpShadow.Color
        .AddParam "balance", sltBalance.Value
        .AddParam "strength", sltStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
