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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   359.9
      SigDigits       =   1
      Value           =   45
   End
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   10
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   200
      SigDigits       =   1
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   200
      Value           =   20
   End
   Begin PhotoDemon.sliderTextCombo sltSpokes 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   8
      Value           =   4
   End
   Begin PhotoDemon.sliderTextCombo sltSoftness 
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "softness:"
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
      TabIndex        =   13
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "spokes"
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
      TabIndex        =   11
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "threshold:"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "strength:"
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
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   3960
      Width           =   960
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance:"
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
      TabIndex        =   4
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label lblTitle 
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
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   2040
      Width           =   660
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

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_Tooltip As clsToolTip

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
    alphaIsRelevant = Not workingDIB.isAlphaBinary(False)
    
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
    Dim mbDIB As pdDIB, mbDIBTemp As pdDIB
    Set mbDIB = New pdDIB
    Set mbDIBTemp = New pdDIB
    
    'We start by creating a threshold DIB from the base image.  This threshold DIB will contain only pure black and pure
    ' white pixels, and we use it to determine the regions of the image that need cross-screen filtering.
    Dim thresholdDIB As pdDIB
    Set thresholdDIB = New pdDIB
    thresholdDIB.createFromExistingDIB workingDIB
    
    'Use the ever-excellent pdFilterLUT class to apply the threshold
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    
    Dim tmpLUT() As Byte
    'cLUT.fillLUT_Threshold tmpLUT, 255 - csThreshold
    cLUT.fillLUT_RemappedRange tmpLUT, 255 - csThreshold, 255, 0, 255
    cLUT.applyLUTsToDIB_Gray thresholdDIB, tmpLUT, True
    
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
    
    'Both paths share an identical base step, however, when we create the initial spoke and place it inside mbDIB.
    ' mbDIB serves as the "master" spoke DIB, and we will also be merging subsequent spokes onto it as we go.
    mbDIB.createFromExistingDIB thresholdDIB
    getMotionBlurredDIB thresholdDIB, mbDIB, csAngle, csDistance, True, ((csSpokes Mod 2) = 0)
    If alphaIsRelevant Then mbDIB.fixPremultipliedAlpha True
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal 1
    End If
    
    'Let's do even spokes first, because they are the simplest.
    If (csSpokes Mod 2) = 0 Then
        
        'For each subsequent pair of spokes, we will render it to its own layer, then merge it down onto the mbDIB layer.
        If csSpokes > 2 Then
        
            numSpokeIterations = (csSpokes \ 2)
            spokeIntervalDegrees = 180 / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "master" mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                mbDIBTemp.createFromExistingDIB thresholdDIB
                getMotionBlurredDIB thresholdDIB, mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, Not alphaIsRelevant
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + i * 2
                End If
                
                'Premultiply alpha (as required by the compositor)
                If alphaIsRelevant Then mbDIBTemp.fixPremultipliedAlpha True
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.quickMergeTwoDibsOfEqualSize mbDIB, mbDIBTemp, BL_LINEARDODGE, 100
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 3 + (i * 2)
                End If
                
            Next i
            
        End If
        
    'Odd spokes are more involved...
    Else
    
        'For each subsequent spoke, we will render it to its own layer, then merge it down onto the mbDIB layer.
        ' (Note that we do not have the luxury of knocking out two spokes at once, as each spoke requires a unique angle.)
        If csSpokes > 1 Then
        
            numSpokeIterations = csSpokes
            spokeIntervalDegrees = 360 / numSpokeIterations
            
            'Now, repeat a simple pattern: for each subsequent spoke, render it to its own layer, then merge it down onto
            ' the "master" mbDIB layer.
            For i = 1 To numSpokeIterations - 1
                
                'Create the new spoke layer
                mbDIBTemp.createFromExistingDIB thresholdDIB
                getMotionBlurredDIB thresholdDIB, mbDIBTemp, csAngle + (i * spokeIntervalDegrees), csDistance, True, False, Not alphaIsRelevant
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2) - 1
                End If
                
                'Premultiply alpha (as required by the compositor)
                If alphaIsRelevant Then mbDIBTemp.fixPremultipliedAlpha True
                
                'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
                ' over-emphasizes bright areas, which gives a nice "bloom" effect.
                cComposite.quickMergeTwoDibsOfEqualSize mbDIB, mbDIBTemp, BL_LINEARDODGE, 100
                
                If Not toPreview Then
                    If userPressedESC() Then GoTo PrematureCrossScreenExit
                    SetProgBarVal 2 + (i * 2)
                End If
                
            Next i
            
        End If
    
    End If
    
    'Remove premultipled alpha from the final, fully composited DIB, and release any temporary DIBs that
    ' are no longer needed.
    If alphaIsRelevant Then mbDIB.fixPremultipliedAlpha False
    thresholdDIB.eraseDIB
    Set mbDIBTemp = Nothing
    
    'We now need to brighten up mbDIB.
    Dim lMax As Long, lMin As Long
    getDIBMaxMinLuminance mbDIB, lMin, lMax
    cLUT.fillLUT_RemappedRange tmpLUT, lMin, lMax, 0, 255
    
    'On top of the remapped range (which is most important), we also gamma-correct the DIB according to the input strength parameter
    Dim gammaLUT() As Byte, finalLUT() As Byte
    cLUT.fillLUT_Gamma gammaLUT, 0.5 + (csStrength / 100)
    cLUT.MergeLUTs tmpLUT, gammaLUT, finalLUT
    cLUT.applyLUTsToDIB_Gray mbDIB, finalLUT, True
    
    'We also want to apply a slight blur to the final result, to improve the feathering of the light boundaries (as they may be
    ' quite sharp due to the remapping).
    If ((Not toPreview) And (csSoftening > 0)) Or (csSoftening * curDIBValues.previewModifier > 0) Then
        
        If toPreview Then
            quickBlurDIB mbDIB, csSoftening * curDIBValues.previewModifier
        Else
            quickBlurDIB mbDIB, csSoftening
        End If
        
    End If
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 3
    End If
    
    'At this point, workingDIB is still intact (phew!).  We are going to mask workingDIB against our newly generate mbDIB image.
    ' This gives a nice, lightly colored version of the star effect, using luminance from the stars, but colors from the
    ' underlying image.
    thresholdDIB.createFromExistingDIB workingDIB
    If alphaIsRelevant Then
        thresholdDIB.fixPremultipliedAlpha True
        mbDIB.fixPremultipliedAlpha True
    End If
    cComposite.quickMergeTwoDibsOfEqualSize thresholdDIB, mbDIB, BL_HARDLIGHT, 100
    
    'thresholdDIB now contains the final, fully processed light effect.
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 2
    End If
    
    'The final step is to merge the light effect onto the original image, using the Strength input parameter
    ' to control opacity of the merge.
    If alphaIsRelevant Then workingDIB.fixPremultipliedAlpha True
    cComposite.quickMergeTwoDibsOfEqualSize workingDIB, thresholdDIB, BL_LINEARDODGE, 100
    
    If alphaIsRelevant Then
        workingDIB.fixPremultipliedAlpha False
        workingDIB.copyAlphaFromExistingDIB alphaBackupDIB
        workingDIB.fixPremultipliedAlpha True
    End If
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit
        SetProgBarVal calculatedProgBarMax - 1
    End If
    
    'Clear all temporary DIBs
    Set mbDIB = Nothing
    Set thresholdDIB = Nothing
    
PrematureCrossScreenExit:
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True
    
End Sub

'Used to motion-blur the intermediate images required by the cross-screen filter
Private Sub getMotionBlurredDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal mbAngle As Double, ByVal mbDistance As Double, Optional ByVal toPreview As Boolean = False, Optional ByVal spokesAreSymmetrical As Boolean = True, Optional ByVal useGDIPlus As Boolean = False)

    Dim finalX As Long, finalY As Long
    finalX = srcDIB.getDIBWidth
    finalY = srcDIB.getDIBHeight
    
    'Before doing any rotating or blurring, we need to increase the size of the image we're working with.  If we
    ' don't do this, intermediate rotation actions will chop off the image's corners, and the resulting effect
    ' will look terrible.
    Dim hScaleAmount As Long, vScaleAmount As Long
    Dim nWidth As Double, nHeight As Double
    Math_Functions.findBoundarySizeOfRotatedRect finalX, finalY, mbAngle, nWidth, nHeight
        
    'Use the rotated size to calculate optimal padding amounts
    hScaleAmount = (nWidth - srcDIB.getDIBWidth) \ 2
    vScaleAmount = (nHeight - srcDIB.getDIBHeight) \ 2
    
    If hScaleAmount < 0 Then hScaleAmount = 0
    If vScaleAmount < 0 Then vScaleAmount = 0
    
    'I built a separate function to enlarge the image and fill the blank borders with clamped pixels from the source image:
    Dim tmpClampDIB As pdDIB
    Set tmpClampDIB = New pdDIB
    padDIBClampedPixels hScaleAmount, vScaleAmount, srcDIB, tmpClampDIB
    
    'Create a second DIB, which will receive the results of this one
    Dim rotateDIB As pdDIB
    Set rotateDIB = New pdDIB
    
    'PD has a number of different rotation engines available.  After profiling each one, I have found GDI+ to be the fastest.
    ' Code for the other engines is still here, in case those methods prove faster after future updates.  (For example, FreeImage may
    ' be faster once they finally implement arbitrary view support, so we don't have to make so many intermediate DIB copies.)
    
    If useGDIPlus Then
    
        'GDI+ code:
        rotateDIB.createBlank tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, tmpClampDIB.getDIBColorDepth, 0, 255
        GDIPlusRotateDIB rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, tmpClampDIB, 0, 0, tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, -mbAngle, InterpolationModeHighQualityBicubic
    
    Else
    
        'FreeImage code:
        Plugin_FreeImage_Expanded_Interface.FreeImageRotateDIBFast tmpClampDIB, rotateDIB, -mbAngle, False, False
        
    End If
    
    'Internal pure-VB code:
    'rotateDIB.createBlank tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, tmpClampDIB.getDIBColorDepth, 0, 255
    'CreateRotatedDIB mbAngle, EDGE_CLAMP, True, tmpClampDIB, rotateDIB, 0.5, 0.5, toPreview, tmpClampDIB.getDIBWidth * 3
    
    'Next, apply a horizontal blur, using the blur radius supplied by the user
    Dim rightRadius As Long
    If spokesAreSymmetrical Then rightRadius = mbDistance Else rightRadius = 0
        
    If CreateHorizontalBlurDIB(mbDistance, rightRadius, rotateDIB, tmpClampDIB, toPreview, tmpClampDIB.getDIBWidth * 3, tmpClampDIB.getDIBWidth) Then
        
        'Finally, rotate the image back to its original orientation, using the opposite parameters of the first conversion.
        ' As before, multiple rotation engines could be used, but GDI+ is presently fastest:
        
        If useGDIPlus Then
        
            'GDI+ code:
            'GDI_Plus.GDIPlusFillDIBRect rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, 0, 255
            'GDIPlusRotateDIB rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, tmpClampDIB, 0, 0, tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, mbAngle, InterpolationModeHighQualityBicubic
        
        Else
        
            'FreeImage code:
            Plugin_FreeImage_Expanded_Interface.FreeImageRotateDIBFast tmpClampDIB, rotateDIB, mbAngle, False, False
            
        End If
        
        'Internal pure-VB code:
        'CreateRotatedDIB -mbAngle, EDGE_CLAMP, True, tmpClampDIB, rotateDIB, 0.5, 0.5, toPreview, tmpClampDIB.getDIBWidth * 3, tmpClampDIB.getDIBWidth * 2
        
        'Erase the temporary clamp DIB
        tmpClampDIB.eraseDIB
        Set tmpClampDIB = Nothing
        
        'rotateDIB now contains the image we want, but it also has all the (now-useless) padding from
        ' the rotate operation.  Chop out the valid section and copy it into workingDIB.
        dstDIB.createFromExistingDIB srcDIB
        BitBlt dstDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, rotateDIB.getDIBDC, hScaleAmount, vScaleAmount, vbSrcCopy
        
    End If
    
    'Erase the temporary rotation DIB
    rotateDIB.eraseDIB
    Set rotateDIB = Nothing
    
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

    'Assign the system hand cursor to all relevant objects
    Set m_Tooltip = New clsToolTip
    makeFormPretty Me, m_Tooltip
        
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
