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
      TabIndex        =   4
      Top             =   2640
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
      TabIndex        =   6
      Top             =   3600
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
      TabIndex        =   7
      Top             =   4560
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
      TabIndex        =   9
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   200
      Value           =   20
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
      TabIndex        =   10
      Top             =   1320
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
      TabIndex        =   8
      Top             =   4200
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
      TabIndex        =   5
      Top             =   3240
      Width           =   945
   End
   Begin VB.Label lblIDEWarning 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   6000
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
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
      Top             =   2280
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
'Last updated: 21/January/15
'Last update: wrap up initial build
'
'Cross-screen filters are physical filters placed over the lens of a camera:
'http://en.wikipedia.org/wiki/Photographic_filter#Cross_screen
'
'Different diffraction patterns in the lens create stars of varying spoke counts in regions where lighting is strong.
'
'Finding a digital replacement for a filter like this is tough; in factm the only one I've seen is a $50 plugin for
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
Dim m_ToolTip As clsToolTip

'Apply a cross-screen blur to an image
'Inputs: 1) luminance threshold for pixels to be considered for filtering
'        2) angle of the generated star patterns
'        3) Distance of the star spokes
'        4) Strength (opacity) of the generated spokes, which is actually just gamma correction applied to the star mask
Public Sub CrossScreenFilter(ByVal csThreshold As Double, ByVal csAngle As Double, ByVal csDistance As Double, ByVal csStrength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying cross-screen filter..."
    
    'Progress reports are manually calculated on this function, as it involves a rather complicated series of steps
    Dim calculatedProgBarMax As Long
    calculatedProgBarMax = 700
    
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
    
    'We start by creating a threshold DIB from the base image.  This threshold DIB will contain only pure black and pure
    ' white pixels, and we use it to determine the regions of the image that need cross-screen filtering.
    Dim thresholdDIB As pdDIB
    Set thresholdDIB = New pdDIB
    thresholdDIB.createFromExistingDIB workingDIB
    
    'When debugging these complicated filters, it is sometimes helpful to dump intermediate images to file.
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "0 - Original image.png", thresholdDIB
    
    'Use the ever-excellent pdFilterLUT class to apply the threshold
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    
    Dim tmpLUT() As Byte
    cLUT.fillLUT_Threshold tmpLUT, 255 - csThreshold
    cLUT.applyLUTsToDIB_Gray thresholdDIB, tmpLUT, True
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "1 - Threshold.png", thresholdDIB
    
    'Progress is reported artificially, because it's too complex to handle using normal means
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 100
    End If
    
    'We now need to produce two motion-blurred versions of the threshold DIB.  These will contain the actual star map.
    Dim mbDIB As pdDIB
    Set mbDIB = New pdDIB
    mbDIB.createFromExistingDIB thresholdDIB
    getMotionBlurredDIB thresholdDIB, mbDIB, csAngle, csDistance, True
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "2 - Motion Blur 1.png", mbDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 300
    End If
    
    'Repeat on a second DIB, but modify the rotation by 90 degrees to create a star shape
    getMotionBlurredDIB thresholdDIB, thresholdDIB, csAngle + 90, csDistance, True
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "3 - Motion Blur 2.png", thresholdDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 500
    End If
    
    'Apply premultiplication to the layers prior to compositing
    mbDIB.fixPremultipliedAlpha True
    thresholdDIB.fixPremultipliedAlpha True
    
    'A pdCompositor class will help us blend these images together
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    'Composite our two motion-blurred images together.  This blend mode is somewhat like alpha-blending, but it
    ' over-emphasizes bright areas, which gives a nice "bloom" effect.
    cComposite.quickMergeTwoDibsOfEqualSize thresholdDIB, mbDIB, BL_LINEARDODGE, 100
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "4 - Composited motion blurs.png", thresholdDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 550
    End If
    
    'Copy the finished DIB from the bottom layer back into mbDIB, and remove premultiplied alpha as necessary
    mbDIB.createFromExistingDIB thresholdDIB
    mbDIB.fixPremultipliedAlpha False
    thresholdDIB.eraseDIB
    
    'We now need to brighten up mbDIB.
    Dim lMax As Long, lMin As Long
    getDIBMaxMinLuminance mbDIB, lMin, lMax
    cLUT.fillLUT_RemappedRange tmpLUT, lMin, lMax, 0, 255
    
    'On top of the remapped range (which is most important), we also gamma-correct the DIB according to the input strength parameter
    Dim gammaLUT() As Byte, finalLUT() As Byte
    cLUT.fillLUT_Gamma gammaLUT, 0.5 + (csStrength / 100)
    cLUT.MergeLUTs tmpLUT, gammaLUT, finalLUT
    cLUT.applyLUTsToDIB_Gray mbDIB, finalLUT, True
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "5 - Gamma corrected motion blurs.png", mbDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 600
    End If
    
    'At this point, workingDIB is still intact (phew!).  We are going to mask workingDIB against our newly generate mbDIB image.
    ' This gives a nice, lightly colored version of the star effect, using luminance from the stars, but colors from the
    ' underlying image.
    thresholdDIB.createFromExistingDIB workingDIB
    cComposite.quickMergeTwoDibsOfEqualSize thresholdDIB, mbDIB, BL_HARDLIGHT, 100
    
    'thresholdDIB now contains the final, fully processed light effect.
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "6 - lighting cues applied to motion blur.png", thresholdDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 650
    End If
    
    'The final step is to merge the light effect onto the original image, using the Strength input parameter
    ' to control opacity of the merge.
    cComposite.quickMergeTwoDibsOfEqualSize workingDIB, thresholdDIB, BL_LINEARDODGE, 100
    
    'QuickSaveDIBAsPNG g_UserPreferences.getDebugPath & "7 - final result.png", workingDIB
    
    If Not toPreview Then
        If userPressedESC() Then GoTo PrematureCrossScreenExit:
        SetProgBarVal 700
    End If
    
    'Clear all temporary DIBs
    Set mbDIB = Nothing
    Set thresholdDIB = Nothing
    
PrematureCrossScreenExit:
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True
    
End Sub

'Used to motion-blur the intermediate images required by the cross-screen filter
Private Sub getMotionBlurredDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal mbAngle As Double, ByVal mbDistance As Double, Optional ByVal toPreview As Boolean = False)

    Dim finalX As Long, finalY As Long
    finalX = srcDIB.getDIBWidth
    finalY = srcDIB.getDIBHeight
    
    'Before doing any rotating or blurring, we need to increase the size of the image we're working with.  If we
    ' don't do this, the rotation will chop off the image's corners, and the resulting motion blur will look terrible.
        
    'If FreeImage is enabled, use it to calculate an optimal extension size.  If it is not enabled, do a
    ' quick-and-dirty estimation using basic geometry.
    Dim hScaleAmount As Long, vScaleAmount As Long
    If g_ImageFormats.FreeImageEnabled Then
                
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
        'Use that handle to request an image rotation, then store the new image's width and height
        Dim nWidth As Long, nHeight As Long
        If fi_DIB <> 0 Then
        
            Dim returnDIB As Long
            returnDIB = FreeImage_Rotate(fi_DIB, -mbAngle, 0)
                    
            nWidth = FreeImage_GetWidth(returnDIB)
            nHeight = FreeImage_GetHeight(returnDIB)
            
            If returnDIB <> 0 Then FreeImage_Unload returnDIB
            FreeImage_Unload fi_DIB
    
        Else
            nWidth = workingDIB.getDIBWidth * 2
            nHeight = workingDIB.getDIBHeight * 2
        End If
        
        'Use the returned size to calculate optimal offsets
        hScaleAmount = (nWidth - srcDIB.getDIBWidth) \ 2
        vScaleAmount = (nHeight - srcDIB.getDIBHeight) \ 2
        
        If hScaleAmount < 0 Then hScaleAmount = 0
        If vScaleAmount < 0 Then vScaleAmount = 0
        
    Else
        
        'This is basically a worst-case estimation of the final image size, and because of that, the function will
        ' be quite slow.  This is a very fringe case, however, as most users should have FreeImage available.
        hScaleAmount = Sqr(srcDIB.getDIBWidth * srcDIB.getDIBWidth + srcDIB.getDIBHeight * srcDIB.getDIBHeight)
        If toPreview Then hScaleAmount = hScaleAmount \ 4 Else hScaleAmount = hScaleAmount \ 2
        vScaleAmount = hScaleAmount
        
    End If
    
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
    
    'GDI+ code:
    rotateDIB.createBlank tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, tmpClampDIB.getDIBColorDepth, 0, 255
    GDIPlusRotateDIB rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, tmpClampDIB, 0, 0, tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, -mbAngle, InterpolationModeHighQualityBicubic
    
    'FreeImage code:
    'Plugin_FreeImage_Expanded_Interface.FreeImageRotateDIBFast tmpClampDIB, rotateDIB, -mbAngle, False, False
    
    'Internal pure-VB code:
    'CreateRotatedDIB mbAngle, EDGE_CLAMP, True, tmpClampDIB, rotateDIB, 0.5, 0.5, toPreview, tmpClampDIB.getDIBWidth * 3
    
    'Next, apply a horizontal blur, using the blur radius supplied by the user
    Dim rightRadius As Long
    rightRadius = mbDistance
        
    If CreateHorizontalBlurDIB(mbDistance, rightRadius, rotateDIB, tmpClampDIB, toPreview, tmpClampDIB.getDIBWidth * 3, tmpClampDIB.getDIBWidth) Then
        
        'Finally, rotate the image back to its original orientation, using the opposite parameters of the first conversion.
        ' As before, multiple rotation engines could be used, but GDI+ is presently fastest:
        
        'GDI+ code:
        GDI_Plus.GDIPlusFillDIBRect rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, 0, 255
        GDIPlusRotateDIB rotateDIB, 0, 0, rotateDIB.getDIBWidth, rotateDIB.getDIBHeight, tmpClampDIB, 0, 0, tmpClampDIB.getDIBWidth, tmpClampDIB.getDIBHeight, mbAngle, InterpolationModeHighQualityBicubic
        
        'FreeImage code:
        'Plugin_FreeImage_Expanded_Interface.FreeImageRotateDIBFast tmpClampDIB, rotateDIB, mbAngle, False, False
        
        'Internal pure-VB code:
        'CreateRotatedDIB -mbAngle, EDGE_CLAMP, True, tmpClampDIB, rotateDIB, 0.5, 0.5, toPreview, tmpClampDIB.getDIBWidth * 3, tmpClampDIB.getDIBWidth * 2
        
        'Erase the temporary clamp DIB
        tmpClampDIB.eraseDIB
        Set tmpClampDIB = Nothing
        
        'rotateDIB now contains the image we want, but it also has all the (now-useless) padding from
        ' the rotate operation.  Chop out the valid section and copy it into workingDIB.
        dstDIB.createFromExistingDIB srcDIB
        BitBlt dstDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, rotateDIB.getDIBDC, hScaleAmount, vScaleAmount, vbSrcCopy
        
        'Erase the temporary rotation DIB
        rotateDIB.eraseDIB
        Set rotateDIB = Nothing
        
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Cross-screen", , buildParams(sltThreshold, sltAngle, sltDistance, sltStrength), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltThreshold = 20
    sltAngle = 45
    sltDistance = 10
    sltStrength = 50
End Sub

Private Sub Form_Activate()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING! This tool is very slow when used inside the IDE. Please compile for best results.")
        lblIDEWarning.Visible = True
    End If
    
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
    If cmdBar.previewsAllowed Then CrossScreenFilter sltThreshold, sltAngle, sltDistance, sltStrength, True, fxPreview
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

Private Sub sltStrength_Change()
    updatePreview
End Sub

Private Sub sltThreshold_Change()
    updatePreview
End Sub
