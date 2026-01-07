VERSION 5.00
Begin VB.Form FormCrossScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Cross-screen"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
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
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   778
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5340
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9419
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "angle"
      Max             =   359.9
      SigDigits       =   1
      Value           =   45
      DefaultValue    =   45
   End
   Begin PhotoDemon.pdSlider sltDistance 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2940
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "distance"
      Min             =   1
      Max             =   200
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   200
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   1140
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "threshold"
      Min             =   1
      Max             =   200
      Value           =   20
      DefaultValue    =   20
   End
   Begin PhotoDemon.pdSlider sltSpokes 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "spokes"
      Min             =   1
      Max             =   8
      Value           =   4
      DefaultValue    =   4
   End
   Begin PhotoDemon.pdSlider sltSoftness 
      Height          =   705
      Left            =   6000
      TabIndex        =   7
      Top             =   4740
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "softness"
   End
End
Attribute VB_Name = "FormCrossScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Cross-Screen (Star) Tool
'Copyright 2014-2026 by Tanner Helland
'Created: 20/January/15
'Last updated: 30/July/17
'Last update: performance improvements, migrate to XML params
'
'Cross-screen filters are physical filters placed over the lens of a camera:
' https://en.wikipedia.org/wiki/Photographic_filter#Cross_screen
'
'Different diffraction patterns in the lens create stars of varying spoke counts in regions where lighting is strong.
'
'Finding a digital replacement for a filter like this is tough; in fact, the only one I've seen is a $50 plugin for
' PhotoShop.  (http://www.scarablabs.com/star-filter-photoshop)  I haven't actually tested that solution, so I can't
' vouch for its performance or quality, but given the rarity of digital versions of this filter, I have to think
' others have run into problems creating their own.  (Which of course, makes our version that much sweeter!  ;)
'
'Performance is pretty good, all things considered, but be careful in the IDE.  As usual, I STRONGLY recommend compiling
' before using this tool.
'
'Note also that this filter wraps a motion-blur-ish effect.  PD has multiple rotation engines available, and after some
' profiling, I've decided to go with GDI+ for this filter.  It's fast and of sufficient quality, but in case we need
' to revisit this decision in the future, I've left the engine code for FreeImage and our own internal PD method as well.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To reduce churn, we reuse a few different temporary DIBs whenever we can
Private m_rotateDIB As pdDIB, m_mbDIB As pdDIB, m_mbDIBTemp As pdDIB, m_thresholdDIB As pdDIB

'Apply a cross-screen blur to an image
'Inputs: 1) luminance threshold for pixels to be considered for filtering
'        2) angle of the generated star patterns
'        3) Distance of the star spokes
'        4) Strength (opacity) of the generated spokes, which is actually just gamma correction applied to the star mask
Public Sub CrossScreenFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying cross-screen filter..."
        
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim csSpokes As Long, csSoftening As Long
    Dim csThreshold As Double, csAngle As Double, csDistance As Double, csStrength As Double
    
    With cParams
        csSpokes = .GetLong("spokes", sltSpokes.Value)
        csThreshold = .GetDouble("threshold", sltThreshold.Value)
        csAngle = .GetDouble("angle", sltAngle.Value)
        csDistance = .GetDouble("distance", sltDistance.Value)
        csStrength = .GetDouble("strength", sltStrength.Value)
        csSoftening = .GetLong("softness", sltSoftness.Value)
    End With
    
    'Progress reports are manually calculated on this function, as it involves a rather complicated series of steps,
    ' whose count is variable based on the number of spokes being processed.
    '
    'Six steps are hard-coded, and the rest are contingent on spoke count.
    Dim calculatedProgBarMax As Long
    calculatedProgBarMax = 6 + csSpokes * 2
    
    'Call prepImageData, which will initialize a workingDIB object for us (with all selection tool masks applied)
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, calculatedProgBarMax
    
    'Distance is calculated as (csDistance / 100) * (smallestImageDimension).  This yields identical results in both the preview
    ' and final image, and it also makes distance scale nicely by image size.
    Dim minDimension As Long
    If workingDIB.GetDIBWidth < workingDIB.GetDIBHeight Then
        minDimension = workingDIB.GetDIBWidth
    Else
        minDimension = workingDIB.GetDIBHeight
    End If
    
    csDistance = (csDistance / 100#) * (minDimension * 0.5)
    If csDistance < 1# Then csDistance = 1#
    
    'We can save a lot of time by avoiding alpha handling.  Query the base image to see if we need to deal with alpha.
    Dim alphaIsRelevant As Boolean
    alphaIsRelevant = Not DIBs.IsDIBAlphaBinary(workingDIB, False)
    
    'If alpha is relevant, we need to make a copy of the current image's alpha channel, so we can restore it when we're done
    Dim alphaBackupDIB As pdDIB
    If alphaIsRelevant Then
        Set alphaBackupDIB = New pdDIB
        alphaBackupDIB.CreateFromExistingDIB workingDIB
    End If
    
    'A pdCompositor class will help us blend various images together
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    'Temporary DIBs are required to assemble all the composite spokes
    If m_mbDIB Is Nothing Then Set m_mbDIB = New pdDIB
    If m_mbDIBTemp Is Nothing Then Set m_mbDIBTemp = New pdDIB
    If m_thresholdDIB Is Nothing Then Set m_thresholdDIB = New pdDIB
    
    'We start by creating a threshold DIB from the base image.  This threshold DIB will contain only pure black and pure
    ' white pixels, and we use it to determine the regions of the image that need cross-screen filtering.
    m_thresholdDIB.CreateFromExistingDIB workingDIB
    
    'Use the ever-excellent pdFilterLUT class to apply the threshold
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    
    Dim tmpLUT() As Byte
    cLUT.FillLUT_RemappedRange tmpLUT, 255 - csThreshold, 255, 0, 255
    cLUT.ApplyLUTsToDIB_Gray m_thresholdDIB, tmpLUT
    
    'Progress is reported artificially, because it's too complex to handle using normal means
    If (Not toPreview) Then
        If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal 1
    End If
    
    Dim i As Long, numSpokeIterations As Long
    Dim spokeIntervalDegrees As Double
    
    'We now need to produce a unique motion-blurred version of the threshold DIB for each "spoke" requested by the user.
    ' There are two code paths here, because even-numbered spokes require half as many calculations (as symmetry allows us
    ' to calculate two spokes at once.
    
    'Both paths share an identical base step, however, when we create the initial spoke and place it inside m_mbDIB.
    ' m_mbDIB serves as the "central" spoke DIB, and we will also be merging subsequent spokes onto it as we go.
    m_mbDIB.CreateFromExistingDIB m_thresholdDIB
    GetMotionBlurredDIB m_thresholdDIB, m_mbDIB, csAngle, csDistance, True, ((csSpokes Mod 2) = 0)
    
    If (Not toPreview) Then
        If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal 1
    End If
    
    'Let's do even spokes first, because they are the simplest.
    If (csSpokes Mod 2) = 0 Then
        
        'For each subsequent pair of spokes, we will render it to its own layer, then merge it down onto the m_mbDIB layer.
        If (csSpokes > 2) Then
        
            numSpokeIterations = (csSpokes \ 2)
            spokeIntervalDegrees = 180# / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "central" m_mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                m_mbDIBTemp.CreateFromExistingDIB m_thresholdDIB
                GetMotionBlurredDIB m_thresholdDIB, m_mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, True
                
                If (Not toPreview) Then
                    If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + i * 2
                End If
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.QuickMergeTwoDibsOfEqualSize m_mbDIB, m_mbDIBTemp, BM_LinearDodge, 100
                
                If (Not toPreview) Then
                    If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 3 + (i * 2)
                End If
                
            Next i
            
        End If
        
    'Odd spokes are more involved...
    Else
    
        'For each subsequent spoke, we will render it to its own layer, then merge it down onto the m_mbDIB layer.
        ' (Note that we do not have the luxury of knocking out two spokes at once, as each spoke requires a unique angle.)
        If (csSpokes > 1) Then
        
            numSpokeIterations = csSpokes
            spokeIntervalDegrees = 360# / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "central" m_mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                m_mbDIBTemp.CreateFromExistingDIB m_thresholdDIB
                GetMotionBlurredDIB m_thresholdDIB, m_mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, False
                
                If (Not toPreview) Then
                    If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2) - 1
                End If
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.QuickMergeTwoDibsOfEqualSize m_mbDIB, m_mbDIBTemp, BM_LinearDodge, 100
                
                If (Not toPreview) Then
                    If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2)
                End If
                
            Next i
            
        End If
    
    End If
    
    'Release any backup DIBs used during the motion blur stage
    If (Not (m_rotateDIB Is Nothing)) And (Not toPreview) Then m_rotateDIB.EraseDIB
    
    'Remove premultipled alpha from the final, fully composited DIB, and release any temporary DIBs that
    ' are no longer needed.
    If alphaIsRelevant Then m_mbDIB.SetAlphaPremultiplication False
    m_thresholdDIB.EraseDIB
    If (Not toPreview) Then Set m_mbDIBTemp = Nothing
    
    'We now need to brighten up m_mbDIB.
    Dim lMax As Long, lMin As Long
    GetDIBMaxMinLuminance m_mbDIB, lMin, lMax
    cLUT.FillLUT_RemappedRange tmpLUT, lMin, lMax, 0, 255
    
    'On top of the remapped range (which is most important), we also gamma-correct the DIB according to the input strength parameter
    Dim gammaLUT() As Byte, finalLUT() As Byte
    cLUT.FillLUT_Gamma gammaLUT, 0.5 + (csStrength * 0.01)
    cLUT.MergeLUTs tmpLUT, gammaLUT, finalLUT
    cLUT.ApplyLUTsToDIB_Gray m_mbDIB, finalLUT
    
    'We also want to apply a slight blur to the final result, to improve the feathering of the light boundaries (as they may be
    ' quite sharp due to the remapping).
    If ((Not toPreview) And (csSoftening > 0)) Or (csSoftening * curDIBValues.previewModifier > 0) Then
        
        If toPreview Then
            QuickBlurDIB m_mbDIB, csSoftening * curDIBValues.previewModifier
        Else
            QuickBlurDIB m_mbDIB, csSoftening
        End If
        
    End If
    
    If (Not toPreview) Then
        If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 3
    End If
    
    'At this point, workingDIB is still intact (phew!).  We are going to mask workingDIB against our newly generate m_mbDIB image.
    ' This gives a nice, lightly colored version of the star effect, using luminance from the stars, but colors from the
    ' underlying image.
    m_thresholdDIB.CreateFromExistingDIB workingDIB
    If alphaIsRelevant Then
        m_thresholdDIB.SetAlphaPremultiplication True
        m_mbDIB.SetAlphaPremultiplication True
    End If
    cComposite.QuickMergeTwoDibsOfEqualSize m_thresholdDIB, m_mbDIB, BM_HardLight, 100
    
    'm_thresholdDIB now contains the final, fully processed light effect.
    If (Not toPreview) Then
        If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 2
    End If
    
    'The final step is to merge the light effect onto the original image, using the Strength input parameter
    ' to control opacity of the merge.
    If alphaIsRelevant Then workingDIB.SetAlphaPremultiplication True
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, m_thresholdDIB, BM_LinearDodge, 100
    
    If alphaIsRelevant Then
        workingDIB.SetAlphaPremultiplication False
        workingDIB.CopyAlphaFromExistingDIB alphaBackupDIB
        workingDIB.SetAlphaPremultiplication True
    End If
    
    If (Not toPreview) Then
        If Interface.UserPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 1
    End If
    
    'If we're not in preview mode, clear all temporary DIBs prior to exiting
    If (Not toPreview) Then Set m_mbDIB = Nothing
    If (Not toPreview) Then Set m_thresholdDIB = Nothing
    
PrematureCrossScreenExit:
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

'Used to motion-blur the intermediate images required by the cross-screen filter.  During testing, I've looked at using
' both a box blur and an IIR blur to generate this; the IIR produces much more natural results, and it can blur in-place,
' which is a double win for us.
Private Sub GetMotionBlurredDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal mbAngle As Double, ByVal mbDistance As Double, Optional ByVal toPreview As Boolean = False, Optional ByVal spokesAreSymmetrical As Boolean = True)

    Dim finalX As Long, finalY As Long
    finalX = srcDIB.GetDIBWidth
    finalY = srcDIB.GetDIBHeight
    
    'Create a second DIB, which will receive the results of this one
    If m_rotateDIB Is Nothing Then Set m_rotateDIB = New pdDIB
    
    'As of October 2015, I've finally cracked the math to have GDI+ generate a rotated+padded+clamped DIB for us.
    ' This greatly simplifies this function, while also providing higher-quality results!
    GDI_Plus.GDIPlus_GetRotatedClampedDIB srcDIB, m_rotateDIB, mbAngle
    
    If Filters_Area.HorizontalBlur_IIR(m_rotateDIB, mbDistance, 1, spokesAreSymmetrical, toPreview, m_rotateDIB.GetDIBWidth * 3, m_rotateDIB.GetDIBWidth) Then
        
        'Finally, we need to rotate the image back to its original orientation, using the opposite parameters of the
        ' first conversion.
        
        'Use GDI+ to apply the inverse rotation.  Note that it will automatically center the rotated image within
        ' the destination boundaries, sparing us the trouble of manually trimming the clamped edges
        dstDIB.CreateFromExistingDIB srcDIB
        GDI_Plus.GDIPlus_RotateDIBPlgStyle m_rotateDIB, dstDIB, -mbAngle, True
        
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Cross-screen", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltSpokes = 4
    sltThreshold = 20
    sltAngle = 45
    sltDistance = 10
    sltStrength = 50
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    'Release all temporary DIBs
    If (Not m_rotateDIB Is Nothing) Then Set m_rotateDIB = Nothing
    If (Not m_mbDIB Is Nothing) Then Set m_mbDIB = Nothing
    If (Not m_mbDIBTemp Is Nothing) Then Set m_mbDIBTemp = Nothing
    If (Not m_thresholdDIB Is Nothing) Then Set m_thresholdDIB = Nothing
    
    'Release any subclasses visual theming and translation code
    ReleaseFormTheming Me
    
End Sub

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.CrossScreenFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltDistance_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltSoftness_Change()
    UpdatePreview
End Sub

Private Sub sltSpokes_Change()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub sltThreshold_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "spokes", sltSpokes.Value
        .AddParam "threshold", sltThreshold.Value
        .AddParam "angle", sltAngle.Value
        .AddParam "distance", sltDistance.Value
        .AddParam "strength", sltStrength.Value
        .AddParam "softness", sltSoftness.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
