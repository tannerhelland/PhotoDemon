VERSION 5.00
Begin VB.Form FormCrossScreen 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cross-screen (stars)"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Max             =   359.9
      SigDigits       =   1
      Value           =   45
   End
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   2940
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "distance"
      Min             =   1
      Max             =   200
      SigDigits       =   1
      Value           =   10
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   720
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "strength"
      Max             =   200
      SigDigits       =   1
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   720
      Left            =   6000
      TabIndex        =   5
      Top             =   1140
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "threshold"
      Min             =   1
      Max             =   200
      Value           =   20
   End
   Begin PhotoDemon.sliderTextCombo sltSpokes 
      Height          =   720
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "spokes"
      Min             =   1
      Max             =   8
      Value           =   4
   End
   Begin PhotoDemon.sliderTextCombo sltSoftness 
      Height          =   720
      Left            =   6000
      TabIndex        =   7
      Top             =   4740
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
'Copyright 2014-2015 by Tanner Helland
'Created: 20/January/15
'Last updated: 26/January/15
'Last update: minor performance and quality tweaks
'
'Cross-screen filters are physical filters placed over the lens of a camera:
' http://en.wikipedia.org/wiki/Photographic_filter#Cross_screen
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
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
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
Public Sub CrossScreenFilter(ByVal csSpokes As Long, ByVal csThreshold As Double, ByVal csAngle As Double, ByVal csDistance As Double, ByVal csStrength As Double, ByVal csSoftening As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying cross-screen filter..."
        
    'Progress reports are manually calculated on this function, as it involves a rather complicated series of steps,
    ' whose count is variable based on the number of spokes being processed.
    '
    'Six steps are hard-coded, and the rest are contingent on spoke count.
    Dim calculatedProgBarMax As Long
    calculatedProgBarMax = 6 + csSpokes * 2
    
    'Call prepImageData, which will initialize a workingDIB object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, calculatedProgBarMax
    
    'Distance is calculated as (csDistance / 100) * (smallestImageDimension).  This yields identical results in both the preview
    ' and final image, and it also makes distance scale nicely by image size.
    Dim minDimension As Long
    If workingDIB.getDIBWidth < workingDIB.getDIBHeight Then
        minDimension = workingDIB.getDIBWidth
    Else
        minDimension = workingDIB.getDIBHeight
    End If
    
    csDistance = (csDistance / 100) * (minDimension * 0.5)
    If csDistance = 0 Then csDistance = 1
    
    'We can save a lot of time by avoiding alpha handling.  Query the base image to see if we need to deal with alpha.
    Dim alphaIsRelevant As Boolean
    alphaIsRelevant = Not DIB_Handler.isDIBAlphaBinary(workingDIB, False)
    
    'If alpha is relevant, we need to make a copy of the current image's alpha channel, so we can restore it when we're done
    Dim alphaBackupDIB As pdDIB
    If alphaIsRelevant Then
        Set alphaBackupDIB = New pdDIB
        alphaBackupDIB.createFromExistingDIB workingDIB
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
    m_thresholdDIB.createFromExistingDIB workingDIB
    
    'Use the ever-excellent pdFilterLUT class to apply the threshold
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    
    Dim tmpLUT() As Byte
    cLUT.fillLUT_RemappedRange tmpLUT, 255 - csThreshold, 255, 0, 255
    cLUT.applyLUTsToDIB_Gray m_thresholdDIB, tmpLUT, True
    
    'Progress is reported artificially, because it's too complex to handle using normal means
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal 1
    End If
    
    Dim i As Long, numSpokeIterations As Long
    Dim spokeIntervalDegrees As Double
    
    'We now need to produce a unique motion-blurred version of the threshold DIB for each "spoke" requested by the user.
    ' There are two code paths here, because even-numbered spokes require half as many calculations (as symmetry allows us
    ' to calculate two spokes at once.
    
    'Both paths share an identical base step, however, when we create the initial spoke and place it inside m_mbDIB.
    ' m_mbDIB serves as the "master" spoke DIB, and we will also be merging subsequent spokes onto it as we go.
    m_mbDIB.createFromExistingDIB m_thresholdDIB
    getMotionBlurredDIB m_thresholdDIB, m_mbDIB, csAngle, csDistance, True, ((csSpokes Mod 2) = 0)
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal 1
    End If
    
    'Let's do even spokes first, because they are the simplest.
    If (csSpokes Mod 2) = 0 Then
        
        'For each subsequent pair of spokes, we will render it to its own layer, then merge it down onto the m_mbDIB layer.
        If csSpokes > 2 Then
        
            numSpokeIterations = (csSpokes \ 2)
            spokeIntervalDegrees = 180 / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "master" m_mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                m_mbDIBTemp.createFromExistingDIB m_thresholdDIB
                getMotionBlurredDIB m_thresholdDIB, m_mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, True
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + i * 2
                End If
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.quickMergeTwoDibsOfEqualSize m_mbDIB, m_mbDIBTemp, BL_LINEARDODGE, 100
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 3 + (i * 2)
                End If
                
            Next i
            
        End If
        
    'Odd spokes are more involved...
    Else
    
        'For each subsequent spoke, we will render it to its own layer, then merge it down onto the m_mbDIB layer.
        ' (Note that we do not have the luxury of knocking out two spokes at once, as each spoke requires a unique angle.)
        If csSpokes > 1 Then
        
            numSpokeIterations = csSpokes
            spokeIntervalDegrees = 360 / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "master" m_mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                m_mbDIBTemp.createFromExistingDIB m_thresholdDIB
                getMotionBlurredDIB m_thresholdDIB, m_mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, False
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2) - 1
                End If
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.quickMergeTwoDibsOfEqualSize m_mbDIB, m_mbDIBTemp, BL_LINEARDODGE, 100
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2)
                End If
                
            Next i
            
        End If
    
    End If
    
    'Release any backup DIBs used during the motion blur stage
    If (Not (m_rotateDIB Is Nothing)) And (Not toPreview) Then m_rotateDIB.eraseDIB
    
    'Remove premultipled alpha from the final, fully composited DIB, and release any temporary DIBs that
    ' are no longer needed.
    If alphaIsRelevant Then m_mbDIB.setAlphaPremultiplication False
    m_thresholdDIB.eraseDIB
    If Not toPreview Then Set m_mbDIBTemp = Nothing
    
    'We now need to brighten up m_mbDIB.
    Dim lMax As Long, lMin As Long
    getDIBMaxMinLuminance m_mbDIB, lMin, lMax
    cLUT.fillLUT_RemappedRange tmpLUT, lMin, lMax, 0, 255
    
    'On top of the remapped range (which is most important), we also gamma-correct the DIB according to the input strength parameter
    Dim gammaLUT() As Byte, finalLUT() As Byte
    cLUT.fillLUT_Gamma gammaLUT, 0.5 + (csStrength / 100)
    cLUT.MergeLUTs tmpLUT, gammaLUT, finalLUT
    cLUT.applyLUTsToDIB_Gray m_mbDIB, finalLUT, True
    
    'We also want to apply a slight blur to the final result, to improve the feathering of the light boundaries (as they may be
    ' quite sharp due to the remapping).
    If ((Not toPreview) And (csSoftening > 0)) Or (csSoftening * curDIBValues.previewModifier > 0) Then
        
        If toPreview Then
            quickBlurDIB m_mbDIB, csSoftening * curDIBValues.previewModifier
        Else
            quickBlurDIB m_mbDIB, csSoftening
        End If
        
    End If
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 3
    End If
    
    'At this point, workingDIB is still intact (phew!).  We are going to mask workingDIB against our newly generate m_mbDIB image.
    ' This gives a nice, lightly colored version of the star effect, using luminance from the stars, but colors from the
    ' underlying image.
    m_thresholdDIB.createFromExistingDIB workingDIB
    If alphaIsRelevant Then
        m_thresholdDIB.setAlphaPremultiplication True
        m_mbDIB.setAlphaPremultiplication True
    End If
    cComposite.quickMergeTwoDibsOfEqualSize m_thresholdDIB, m_mbDIB, BL_HARDLIGHT, 100
    
    'm_thresholdDIB now contains the final, fully processed light effect.
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 2
    End If
    
    'The final step is to merge the light effect onto the original image, using the Strength input parameter
    ' to control opacity of the merge.
    If alphaIsRelevant Then workingDIB.setAlphaPremultiplication True
    cComposite.quickMergeTwoDibsOfEqualSize workingDIB, m_thresholdDIB, BL_LINEARDODGE, 100
    
    If alphaIsRelevant Then
        workingDIB.setAlphaPremultiplication False
        workingDIB.copyAlphaFromExistingDIB alphaBackupDIB
        workingDIB.setAlphaPremultiplication True
    End If
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 1
    End If
    
    'If we're not in preview mode, clear all temporary DIBs prior to exiting
    If Not toPreview Then Set m_mbDIB = Nothing
    If Not toPreview Then Set m_thresholdDIB = Nothing
    
PrematureCrossScreenExit:
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True
    
End Sub

'Used to motion-blur the intermediate images required by the cross-screen filter.  During testing, I've looked at using
' both a box blur and an IIR blur to generate this; the IIR produces much more natural results, and it can blur in-place,
' which is a double win for us.
Private Sub getMotionBlurredDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal mbAngle As Double, ByVal mbDistance As Double, Optional ByVal toPreview As Boolean = False, Optional ByVal spokesAreSymmetrical As Boolean = True)

    Dim finalX As Long, finalY As Long
    finalX = srcDIB.getDIBWidth
    finalY = srcDIB.getDIBHeight
    
    'Create a second DIB, which will receive the results of this one
    If m_rotateDIB Is Nothing Then Set m_rotateDIB = New pdDIB
    
    'As of October 2015, I've finally cracked the math to have GDI+ generate a rotated+padded+clamped DIB for us.
    ' This greatly simplifies this function, while also providing higher-quality results!
    GDI_Plus.GDIPlus_GetRotatedClampedDIB srcDIB, m_rotateDIB, mbAngle
    
    If Filters_Area.HorizontalBlur_IIR(m_rotateDIB, mbDistance, 1, spokesAreSymmetrical, toPreview, m_rotateDIB.getDIBWidth * 3, m_rotateDIB.getDIBWidth) Then
        
        'Finally, we need to rotate the image back to its original orientation, using the opposite parameters of the
        ' first conversion.
        
        'Use GDI+ to apply the inverse rotation.  Note that it will automatically center the rotated image within
        ' the destination boundaries, sparing us the trouble of manually trimming the clamped edges
        dstDIB.createFromExistingDIB srcDIB
        GDI_Plus.GDIPlus_RotateDIBPlgStyle m_rotateDIB, dstDIB, -mbAngle, True
        
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Cross-screen", , buildParams(sltSpokes, sltThreshold, sltAngle, sltDistance, sltStrength, sltSoftness), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltSpokes = 4
    sltThreshold = 20
    sltAngle = 45
    sltDistance = 10
    sltStrength = 50
End Sub

Private Sub Form_Activate()

    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    'Release all temporary DIBs
    If Not m_rotateDIB Is Nothing Then Set m_rotateDIB = Nothing
    If Not m_mbDIB Is Nothing Then Set m_mbDIB = Nothing
    If Not m_mbDIBTemp Is Nothing Then Set m_mbDIBTemp = Nothing
    If Not m_thresholdDIB Is Nothing Then Set m_thresholdDIB = Nothing
    
    'Release any subclasses visual theming and translation code
    ReleaseFormTheming Me
    
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then CrossScreenFilter sltSpokes, sltThreshold, sltAngle, sltDistance, sltStrength, sltSoftness, True, fxPreview
End Sub

Private Sub sliderTextCombo1_Change()
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltSoftness_Change()
    updatePreview
End Sub

Private Sub sltSpokes_Change()
    updatePreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub

Private Sub sltThreshold_Change()
    updatePreview
End Sub
