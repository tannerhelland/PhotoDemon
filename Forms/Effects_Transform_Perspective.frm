VERSION 5.00
Begin VB.Form FormPerspective 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Perspective"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14175
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
   ScaleHeight     =   635
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBoxInteractive picDraw 
      Height          =   8475
      Left            =   6000
      Top             =   120
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   14949
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   4305
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   7594
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   8775
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonStrip btsSettings 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      Caption         =   "options"
   End
   Begin PhotoDemon.pdContainer pnlSettings 
      Height          =   2895
      Index           =   1
      Left            =   120
      Top             =   5760
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   5106
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   372
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "custom foreshortening (x, y)"
         FontSize        =   12
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1270
         Caption         =   "quality"
         Min             =   1
         Max             =   5
         Value           =   2
         NotchPosition   =   2
         NotchValueCustom=   2
      End
      Begin PhotoDemon.pdDropDown cboEdges 
         Height          =   855
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1508
         Caption         =   "if pixels lie outside the image..."
      End
      Begin PhotoDemon.pdSlider sldForeshortening 
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   14
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         Min             =   -3
         Max             =   3
         SigDigits       =   2
         GradientColorRight=   1703935
      End
      Begin PhotoDemon.pdSlider sldForeshortening 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         Min             =   -3
         Max             =   3
         SigDigits       =   2
         GradientColorRight=   1703935
      End
   End
   Begin PhotoDemon.pdContainer pnlSettings 
      Height          =   2895
      Index           =   0
      Left            =   120
      Top             =   5760
      Width           =   5790
      _ExtentX        =   11456
      _ExtentY        =   7223
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   960
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "coordinates (x, y)"
         FontSize        =   12
      End
      Begin PhotoDemon.pdDropDown cboMapping 
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1296
         Caption         =   "transformation type"
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   7
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   4
         Left            =   3240
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner spnCoords 
         Height          =   375
         Index           =   5
         Left            =   4440
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Min             =   -32000
         Max             =   32000
      End
   End
End
Attribute VB_Name = "FormPerspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Perspective Distortion
'Copyright 2013-2026 by Tanner Helland
'Created: 08/April/13
'Last updated: 02/November/22
'Last update: new custom foreshortening support (see https://github.com/tannerhelland/PhotoDemon/issues/454)
'
'This tool allows the user to remap their image (or layer) to any arbitrary quadrilateral.  The code is
' fairly standard linear algebra, as a series of equations must be solved to generate the homography matrix
' required by the transform.  For a more detailed explanation of the math and theory behind this transformation,
' called a "projective transform", see Wikipedia:
'
' https://en.wikipedia.org/wiki/Homography
'
'As with all distort tools in PhotoDemon, reverse-mapping plus supersampling is supported for high-quality
' antialiasing.  A "bonus" simpler remapping function is also provided for generating the on-screen preview
' of the effect; this may be a more useful reference for beginners, although it only operates at a fixed
' quality with much more limited processing options.
'
'I learned from a number of helpful references while building this tool.  Thank you to the following resources:
'
' http://www.cs.cmu.edu/~ph/texfund/texfund.pdf
' http://www.imagemagick.org/Usage/distorts/#perspective
' http://stackoverflow.com/questions/169902/projective-transformation
' http://freespace.virgin.net/hugo.elias/graphics/x_persp.htm
' http://stackoverflow.com/questions/530396/how-to-draw-a-perspective-correct-grid-in-2d?lq=1
' https://stackoverflow.com/questions/471962/how-do-i-efficiently-determine-if-a-polygon-is-convex-non-convex-or-complex
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify all measurements by the ratio between the (generally smaller)
' preview image and the original full-size image.  These values are required for mapping between
' the interactive UI area and the image coordinate space during the final transform.
Private m_OrigImageWidth As Double, m_OrigImageHeight As Double

'Width and height of the preview image, cached locally
Private m_PreviewWidth As Long, m_PreviewHeight As Long

'To improve performance, we cache a copy of the source image; the perspective transform obviously
' requires discrete source and destination images, and we reduce memory thrashing by maintaining a
' persistent source copy during previews.
Private m_srcDIB As pdDIB

'We track two sets of control point coordinates - the original points, and the new points.  The difference between
' these is passed to the perspective function.
Private m_oPoints(0 To 3) As PointFloat
Private m_nPoints(0 To 3) As PointFloat

'The perspective function can report-back the coordinates of the perspective transform, translated into
' "final" destination space.  We use these to raise a useful tooltip for the user, so they can see precise
' measurements of each point.
Private m_dstLayerSpacePoints(0 To 3) As PointFloat

'Mouse status is tracked between MouseDown and MouseMove events; this allows for drag events.
Private m_isMouseDown As Boolean

'Currently selected and hovered nodes in the workspace area (if any); -1 = no point selected
Private m_ActivePoint As Long, m_HoverPoint As Long

'Buffer to which the current interactive "perspective" control is rendered.
Private m_Buffer As pdDIB

'Overlay for the interactive buffer where we pre-render a fast, lower-quality version of the
' "perspective" copy of the image.  Must be zeroed before rendering.
Private m_Overlay As pdDIB

'The current mouse coordinates are rendered to a dedicated image, which is then overlaid atop the interactive box
Private m_mouseCoordFont As pdFont, m_mouseCoordDIB As pdDIB

'At load-time we'll generate a fixed-size copy of the source layer at a proportional size to the
' interactive tool area.  This improves performance by limiting the amount of memory that we have to
' "dip into" while rendering the on-screen preview.
Private m_ProportionalSource As pdDIB

'Prevent recursive redraws when synchronizing text UI elements
Private m_SuspendSync As Boolean

'The user can also move all four points simultaneously.  To enable this, we have to cache initial point positions
' at _MouseDown.
Private m_PointsAtMoveStart(0 To 3) As PointFloat, m_InitPoint As PointFloat, m_MoveActive As Boolean

'Apply horizontal and/or vertical perspective to an image by shrinking it in one or more directions
' Input: the coordinates of the four corners of the transformed image, stored inside a "|"-delimited string.  To see how
'        these points are generated by the preview picture box, visit the getPerspectiveParamString() function at the
'        bottom of this page.
Public Sub PerspectiveImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Applying new perspective..."
    
    'We use an XML parser to retrieve individual parameters from the incoming parameter string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent translated pixels from spreading across the image as we go.)
    If (m_srcDIB Is Nothing) Then Set m_srcDIB = New pdDIB
    m_srcDIB.CreateFromExistingDIB workingDIB
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    'Local loop variables can be more efficiently cached by VB's compiler, so transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'See if the user wants a rect -> quad ("Normal" in GIMP) or quad -> rect ("Corrective" in GIMP) mapping
    Dim correctiveProjection As Boolean
    correctiveProjection = (cParams.GetLong("mapping", 1) <> 0)
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters cParams.GetLong("edges", pdeo_Erase), (cParams.GetLong("quality", 1) <> 1), curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    '***************************************
    ' /* BEGIN SUPERSAMPLING PREPARATION */
    
    'Start by retrieving the supersampling parameter from the param string
    Dim superSamplingAmount As Long
    superSamplingAmount = cParams.GetLong("quality", 1)
    
    'Due to the way this filter works, supersampling yields much better results.  Because supersampling is extremely
    ' energy-intensive, this tool uses a sliding value for quality, as opposed to a binary TRUE/FALSE for antialiasing.
    ' (For all but the lowest quality setting, antialiasing will be used, and higher quality values will simply increase
    '  the amount of supersamples taken.)
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim tmpSum As Long, tmpSumFirst As Long
    
    'Use the passed super-sampling constant (displayed to the user as "quality") to come up with a number of actual
    ' pixels to sample.  (The total amount of sampled pixels will range from 1 to 13).  Note that supersampling
    ' coordinates are precalculated and cached using a modified rotated grid function, which is consistent throughout PD.
    Dim numSamples As Long
    Dim ssX() As Single, ssY() As Single
    Filters_Area.GetSupersamplingTable superSamplingAmount, numSamples, ssX, ssY
    
    'Because supersampling will be used in the inner loop as (samplecount - 1), permanently decrease the sample
    ' count in advance.
    numSamples = numSamples - 1
    
    'Additional variables are needed for supersampling handling
    Dim sampleIndex As Long, numSamplesUsed As Long
    Dim superSampleVerify As Long, ssVerificationLimit As Long
    
    'Adaptive supersampling allows us to bypass supersampling if a pixel doesn't appear to benefit from it.  The superSampleVerify
    ' variable controls how many pixels are sampled before we perform an adaptation check.  At present, the rule is:
    ' Quality 3: check a minimum of 2 samples, Quality 4: check minimum 3 samples, Quality 5: check minimum 4 samples
    superSampleVerify = superSamplingAmount - 2
    
    'Alongside a variable number of test samples, adaptive supersampling requires some threshold that indicates samples
    ' are close enough that further supersampling is unlikely to improve output.  We calculate this as a minimum variance
    ' as 1.5 per channel (for a total of 6 variance per pixel), multiplied by the total number of samples taken.
    ssVerificationLimit = superSampleVerify * 6
    
    'To improve performance for quality 1 and 2 (which perform no supersampling), we can forcibly disable supersample checks
    ' by setting the verification checker to some impossible value.
    If (superSampleVerify <= 0) Then superSampleVerify = LONG_MAX
    
    ' /* END SUPERSAMPLING PREPARATION */
    '*************************************
    
    'Store region width and height as floating-point
    Dim imgWidth As Double, imgHeight As Double
    imgWidth = (finalX - initX) - 1
    imgHeight = (finalY - initY) - 1
    
    'If this is a preview, we need to adjust the width and height values to match the size of the preview box
    Dim wModifier As Double, hModifier As Double
    wModifier = curDIBValues.previewModifier
    hModifier = curDIBValues.previewModifier
    
    'Scale quad coordinates to the size of the image.  (This is useful for batch processing, because we perform all
    ' calculations in terms of the unit square.  Thus the transform values of a 1000x1000 image will still be valid
    ' for say a 100x100 image.)
    Dim invWidth As Double, invHeight As Double
    If (imgWidth > 0#) Then invWidth = 1# / imgWidth Else invWidth = 999999#
    If (imgHeight > 0#) Then invHeight = 1# / imgHeight Else invHeight = 999999#
    
    'Foreshortening is handled manually, *after* we've mapped pixels to the target domain.
    '
    '(For "normal" perspective rendering, you could skip foreshortening entirely.
    ' It adds a non-trivial amount of processing time to the function.)
    Dim xForeshortening As Double, yForeshortening As Double, customForeshortening As Boolean, useOrigCoords As Boolean
    xForeshortening = TranslateForeshorteningUIValue(cParams.GetDouble("x-foreshorten", 0#, True))
    yForeshortening = TranslateForeshorteningUIValue(cParams.GetDouble("y-foreshorten", 0#, True))
    customForeshortening = (xForeshortening <> 1#) Or (yForeshortening <> 1#)
    
    'Copy the points given by the user (which are currently strings) into individual floating-point variables
    Dim x0 As Double, x1 As Double, x2 As Double, x3 As Double
    Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
    
    x0 = cParams.GetDouble("topleftx")
    y0 = cParams.GetDouble("toplefty")
    x1 = cParams.GetDouble("toprightx")
    y1 = cParams.GetDouble("toprighty")
    x2 = cParams.GetDouble("bottomrightx")
    y2 = cParams.GetDouble("bottomrighty")
    x3 = cParams.GetDouble("bottomleftx")
    y3 = cParams.GetDouble("bottomlefty")
    
    If toPreview Then
        x0 = x0 * wModifier
        y0 = y0 * hModifier
        x1 = x1 * wModifier
        y1 = y1 * hModifier
        x2 = x2 * wModifier
        y2 = y2 * hModifier
        x3 = x3 * wModifier
        y3 = y3 * hModifier
    End If
    
    'Map those coordinates to the unit square by multiplying them by "1 / image_width"
    ' or "1 / image_height" as appropriate.
    x0 = x0 * invWidth
    y0 = y0 * invHeight
    x1 = x1 * invWidth
    y1 = y1 * invHeight
    x2 = x2 * invWidth
    y2 = y2 * invHeight
    x3 = x3 * invWidth
    y3 = y3 * invHeight
    
    'First things first: we need to map the original image (now in terms of the unit square)
    ' to the arbitrary quadrilateral defined by the user's parameters
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dx3 As Double, dy3 As Double
    dx1 = x1 - x2
    dy1 = y1 - y2
    dx2 = x3 - x2
    dy2 = y3 - y2
    dx3 = x0 - x1 + x2 - x3
    dy3 = y0 - y1 + y2 - y3
    
    'Technically, these are points in a matrix - and they could be defined as an array.  But VB accesses
    ' individual data types more quickly than an array, so we declare them separately.
    Dim h11 As Double, h21 As Double, h31 As Double
    Dim h12 As Double, h22 As Double, h32 As Double
    Dim h13 As Double, h23 As Double, h33 As Double
    
    'Certain values can lead to divide-by-zero problems - check those in advance and convert 0 to something like 0.000001
    Dim chkDenom As Double
    chkDenom = (dx1 * dy2 - dy1 * dx2)
    If (chkDenom < 1E-20) And (chkDenom > -1 * 1E-20) Then chkDenom = 1E-20
    
    h13 = (dx3 * dy2 - dx2 * dy3) / chkDenom
    h23 = (dx1 * dy3 - dy1 * dx3) / chkDenom
    h11 = x1 - x0 + h13 * x1
    h21 = x3 - x0 + h23 * x3
    h31 = x0
    h12 = y1 - y0 + h13 * y1
    h22 = y3 - y0 + h23 * y3
    h32 = y0
    h33 = 1
    
    'Next, we need to calculate the key set of transformation parameters, using the reverse-map data we just generated.
    ' Again, these are technically just matrix entries, but we get better performance by declaring them individually.
    Dim hA As Double, hB As Double, hC As Double
    Dim hD As Double, hE As Double, hF As Double
    Dim hG As Double, hH As Double, hI As Double
    
    hA = h22 * h33 - h32 * h23
    hB = h31 * h23 - h21 * h33
    hC = h21 * h32 - h31 * h22
    hD = h32 * h13 - h12 * h33
    hE = h11 * h33 - h31 * h13
    hF = h31 * h12 - h11 * h32
    hG = h12 * h23 - h22 * h13
    hH = h21 * h13 - h11 * h23
    hI = h11 * h22 - h21 * h12
        
    'We now have two options.  We can either...
    ' 1) Proceed with the projection mapping, which assumes a RECTANGULAR source area and a QUADRILATERAL destination
    '    area. (GIMP calls this Normal/Forward which is confusing from a coding standpoint, because it's still
    '    reverse-mapping, where the user-drawn quadrilateral defines the boundaries of the destination area.)
    ' 2) Invert the mapping matrix we've created, which changes the assumption to a QUADRILATERAL source area and a
    '    RECTANGULAR destination area.  (GIMP calls this Corrective/Backward which is again confusing, as "Normal"
    '    mapping is still a perfectly valid way to correct perspective distortion.  Anyway, this additional operation
    '    simply changes the user-drawn quadrilateral to define the boundaries of the source area.)
    If correctiveProjection Then
    
        'Invert the transformation using the adjoint of the forward mapping.  Said another way, we're basically
        ' reversing the plane-to-plane mapping that defines this projection.  (This means we want the quadrilateral
        ' to define a section of the SOURCE image instead of a section of the DESTINATION image.)
        '
        'For a detailed explanation of this process, please read pages 24-25 of Paul Heckbert's thesis on projective
        ' transformations, which is IMO a great source for understanding projective mappings in general:
        ' http://www.cs.cmu.edu/~ph/texfund/texfund.pdf
        Dim newA2 As Double, newB2 As Double, newC As Double
        Dim newD As Double, newE As Double, newF As Double
        Dim newG2 As Double, newH As Double, newI As Double
        
        newA2 = hE * hI - hF * hH
        newB2 = hC * hH - hB * hI
        newC = hB * hF - hC * hE
        
        newD = hF * hG - hD * hI
        newE = hA * hI - hC * hG
        newF = hC * hD - hA * hF
        
        newG2 = hD * hH - hE * hG
        newH = hB * hG - hA * hH
        newI = hA * hE - hB * hD
    
        hA = newA2
        hB = newB2
        hC = newC
        hD = newD
        hE = newE
        hF = newF
        hG = newG2
        hH = newH
        hI = newI
        
    End If
        
    'Scale those values to match the size of the transformed image
    hA = hA * invWidth
    hD = hD * invWidth
    hG = hG * invWidth
    hB = hB * invHeight
    hE = hE * invHeight
    hH = hH * invHeight
    
    'With all this data calculated in advanced, we can now proceed with the actual transform  - and it's quite simple!
    Dim srcX As Double, srcY As Double
    Dim newX As Double, newY As Double
    
    Dim tmpQuad As RGBQuad
    fSupport.AliasTargetDIB m_srcDIB
    
    'Loop through each pixel in the image, converting values as we go.  Note that PD now guarantees 32-bpp inputs,
    ' which allows us to skip the "check for alpha" part of this process.
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
        
        'Reset all supersampling values
        newR = 0
        newG = 0
        newB = 0
        newA = 0
        numSamplesUsed = 0
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            'Pull coordinates from the lookup table
            newX = x + ssX(sampleIndex)
            newY = y + ssY(sampleIndex)
            
            'Reverse-map the coordinates back onto the original image (to allow for resampling)
            chkDenom = (hG * newX + hH * newY + hI)
            If (chkDenom <> 0#) Then chkDenom = 1# / chkDenom
            
            srcX = (hA * newX + hB * newY + hC) * chkDenom
            srcY = (hD * newX + hE * newY + hF) * chkDenom
            
            'Apply custom foreshortening, if any.
            If customForeshortening Then
                
                'Use the filter support class to correctly handle wrap/reflect/etc
                If fSupport.HandleEdgesOnly_Normalized(srcX, srcY, useOrigCoords) Then
                    
                    'Pixel is being erased; ignore it
                    tmpQuad.Blue = 0
                    tmpQuad.Green = 0
                    tmpQuad.Red = 0
                    tmpQuad.Alpha = 0
                
                'Pixel is in-bounds
                Else
                    
                    'The "use original coordinates" edge-wrap mode must be handled specially
                    If useOrigCoords Then
                        srcX = x
                        srcY = y
                    Else
                        
                        'Apply custom forshortening
                        If (xForeshortening <> 1#) Then srcX = srcX ^ xForeshortening
                        If (yForeshortening <> 1#) Then srcY = srcY ^ yForeshortening
                        
                        'Scale back to normal image space
                        srcX = srcX * imgWidth
                        srcY = srcY * imgHeight
                        
                    End If
                    
                    'Use the filter support class to interpolate colors (if quality settings allow)
                    tmpQuad = fSupport.HandleInterpolationOnly(srcX, srcY)
                    
                End If
                
            'Normal mode is much easier to handle
            Else
                
                'Scale back to the normal image space domain
                srcX = srcX * imgWidth
                srcY = srcY * imgHeight
                
                'Use the filter support class to interpolate and edge-wrap pixels as necessary
                tmpQuad = fSupport.GetColorsFromSource(srcX, srcY, x, y)
                
            End If
            
            'If supersampling is active, we need to individually track r/g/b/a components
            b = tmpQuad.Blue
            g = tmpQuad.Green
            r = tmpQuad.Red
            a = tmpQuad.Alpha
            
            'If adaptive supersampling is active, apply the "adaptive" aspect.  Basically, calculate a variance for the currently
            ' collected samples.  If variance is low, assume this pixel does not require further supersampling.
            ' (Note that this is an ugly shorthand way to calculate variance, but it's fast, and the chance of false outliers is
            '  small enough to make it preferable over a true variance calculation.)
            If (sampleIndex = superSampleVerify) Then
                
                'Calculate variance for the first two pixels (Q3), three pixels (Q4), or four pixels (Q5)
                tmpSum = (r + g + b + a) * superSampleVerify
                tmpSumFirst = newR + newG + newB + newA
                
                'If variance is below 1.5 per channel per pixel, abort further supersampling
                If (Abs(tmpSum - tmpSumFirst) < ssVerificationLimit) Then Exit For
            
            End If
            
            'Increase the sample count
            numSamplesUsed = numSamplesUsed + 1
            
            'Add the retrieved values to our running averages
            newR = newR + r
            newG = newG + g
            newB = newB + b
            newA = newA + a
        
        Next sampleIndex
        
        'Find the average values of all samples, apply to the pixel, and move on!
        If (numSamplesUsed > 1) Then
            newR = newR \ numSamplesUsed
            newG = newG \ numSamplesUsed
            newB = newB \ numSamplesUsed
            newA = newA \ numSamplesUsed
        End If
        
        xStride = x * 4
        dstImageData(xStride) = newB
        dstImageData(xStride + 1) = newG
        dstImageData(xStride + 2) = newR
        dstImageData(xStride + 3) = newA
        
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
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'Greatly simplified perspective renderer for rendering a high-speed preview image onto the interactive area.
' Note that some special handling is required for things like the transparency grid (which is properly
' constrained to image boundaries in the interactive area!)
Private Sub RenderImageToArbitraryQuad(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef listOfPoints() As PointFloat)
    
    'Zero out the destination DIB before doing any actual rendering
    dstDIB.ResetDIB 0
    Dim dstImageData() As RGBQuad, dstSA1D As SafeArray1D
    Dim srcImageData() As RGBQuad, srcSA As SafeArray2D
    
    srcDIB.WrapRGBQuadArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    
    'Calculate min/max values for the destination points.  We only need to render within this area.
    Dim minX As Single, maxX As Single, minY As Single, maxY As Single
    minX = listOfPoints(0).x
    maxX = listOfPoints(0).x
    minY = listOfPoints(0).y
    minY = listOfPoints(0).y
    
    For x = 1 To 3
        If (listOfPoints(x).x < minX) Then minX = listOfPoints(x).x
        If (listOfPoints(x).x > maxX) Then maxX = listOfPoints(x).x
        If (listOfPoints(x).y < minY) Then minY = listOfPoints(x).y
        If (listOfPoints(x).y > maxY) Then maxY = listOfPoints(x).y
    Next x
    
    'Clamp to dimensions of the destination image
    initX = minX
    initY = minY
    finalX = maxX
    finalY = maxY
    
    If (initX < 0) Then initX = 0
    If (initY < 0) Then initY = 0
    If (finalX > dstDIB.GetDIBWidth - 1) Then finalX = dstDIB.GetDIBWidth - 1
    If (finalY > dstDIB.GetDIBHeight - 1) Then finalY = dstDIB.GetDIBHeight - 1
    
    'Paint a checkerboard over the target area.  (We will blank it out as relevant.)
    GDI_Plus.GDIPlusFillDIBRect_Pattern dstDIB, initX, initY, finalX - initX, finalY - initY, g_CheckerboardPattern, fixBoundaryPainting:=True, noAntialiasing:=True
    
    'Store destination and source width and height; we use these for all calculations to ensure correct
    ' scaling between surfaces.
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = (maxX - minX) - 1
    If (dstWidth < 1#) Then dstWidth = 1#
    dstHeight = (maxY - minY) - 1
    If (dstHeight < 1#) Then dstHeight = 1#
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = srcDIB.GetDIBWidth - 1
    srcHeight = srcDIB.GetDIBHeight - 1
    
    'Calculate translations between source/destination sizes and the unit square
    Dim invDstWidth As Double, invDstHeight As Double, invSrcWidth As Double, invSrcHeight As Double
    If (dstWidth > 0#) Then invDstWidth = 1# / dstWidth Else invDstWidth = 999999999#
    If (dstHeight > 0#) Then invDstHeight = 1# / dstHeight Else invDstHeight = 999999999#
    If (srcWidth > 0#) Then invSrcWidth = 1# / srcWidth Else invSrcWidth = 999999999#
    If (srcHeight > 0#) Then invSrcHeight = 1# / srcHeight Else invSrcHeight = 999999999#
    
    'Copy the points given by the user (which are currently strings) into individual floating-point variables
    Dim x0 As Double, x1 As Double, x2 As Double, x3 As Double
    Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
    
    'Scale the incoming points relative to minimum coordinate values (so that [0, 0] becomes the new minimum)
    x0 = (listOfPoints(0).x - minX) * invDstWidth
    y0 = (listOfPoints(0).y - minY) * invDstHeight
    x1 = (listOfPoints(1).x - minX) * invDstWidth
    y1 = (listOfPoints(1).y - minY) * invDstHeight
    x2 = (listOfPoints(2).x - minX) * invDstWidth
    y2 = (listOfPoints(2).y - minY) * invDstHeight
    x3 = (listOfPoints(3).x - minX) * invDstWidth
    y3 = (listOfPoints(3).y - minY) * invDstHeight
    
    'Start calculating the projection homography between the source image (full size) and the arbitrary
    ' quadrilateral we were handed.
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dx3 As Double, dy3 As Double
    dx1 = x1 - x2
    dy1 = y1 - y2
    dx2 = x3 - x2
    dy2 = y3 - y2
    dx3 = x0 - x1 + x2 - x3
    dy3 = y0 - y1 + y2 - y3
    
    'Technically, these are points in a matrix - and they could be defined as an array.  But VB accesses
    ' individual data types more quickly than an array, so we declare them separately.
    Dim h11 As Double, h21 As Double, h31 As Double
    Dim h12 As Double, h22 As Double, h32 As Double
    Dim h13 As Double, h23 As Double, h33 As Double
    
    'Certain values can lead to divide-by-zero problems - check those in advance and convert 0 to
    ' an arbitrary, extremely small value
    Dim chkDenom As Double
    chkDenom = (dx1 * dy2 - dy1 * dx2)
    If (chkDenom < 1E-20) And (chkDenom > -1 * 1E-20) Then chkDenom = 1E-20
    
    h13 = (dx3 * dy2 - dx2 * dy3) / chkDenom
    h23 = (dx1 * dy3 - dy1 * dx3) / chkDenom
    h11 = x1 - x0 + h13 * x1
    h21 = x3 - x0 + h23 * x3
    h31 = x0
    h12 = y1 - y0 + h13 * y1
    h22 = y3 - y0 + h23 * y3
    h32 = y0
    h33 = 1
    
    'Next, we need to calculate the key set of transformation parameters, using the reverse-map data
    ' we just generated.  (We don't want a forward map - we want a *reverse* map from the destination
    ' to the source image, so we can interpolate as relevant.)
    Dim hA As Double, hB As Double, hC As Double
    Dim hD As Double, hE As Double, hF As Double
    Dim hG As Double, hH As Double, hI As Double
    
    hA = h22 * h33 - h32 * h23
    hB = h31 * h23 - h21 * h33
    hC = h21 * h32 - h31 * h22
    hD = h32 * h13 - h12 * h33
    hE = h11 * h33 - h31 * h13
    hF = h31 * h12 - h11 * h32
    hG = h12 * h23 - h22 * h13
    hH = h21 * h13 - h11 * h23
    hI = h11 * h22 - h21 * h12
    
    'With all this data calculated in advanced, we can now proceed with the actual transform - and it's quite simple!
    Dim srcX As Double, srcY As Double, srcXInt As Long, srcYInt As Long
    Dim newX As Double, newY As Double
    
    'Foreshortening is handled manually, *after* we've mapped pixels to the target domain.
    '
    '(For "normal" perspective rendering, you could skip foreshortening entirely.
    ' It adds a non-trivial amount of processing time to the function.)
    Dim xForeshortening As Double, yForeshortening As Double
    xForeshortening = TranslateForeshorteningUIValue(sldForeshortening(0).Value)
    yForeshortening = TranslateForeshorteningUIValue(sldForeshortening(1).Value)
    
    'We're going to manually blank out pixels "outside" the quadrilateral
    Dim zeroQuad As RGBQuad
    zeroQuad.Red = 0
    zeroQuad.Green = 0
    zeroQuad.Blue = 0
    zeroQuad.Alpha = 0
    
    Dim newAlpha As Single
    
    'Loop through each pixel in the image, converting values as we go.
    For y = initY To finalY
        dstDIB.WrapRGBQuadArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
    
        'Scale coordinates to the unit square
        newX = (x - minX) * invDstWidth
        newY = (y - minY) * invDstHeight
        
        'Reverse-map the coordinates back onto the original image (to allow for resampling)
        chkDenom = (hG * newX + hH * newY + hI)
        If (chkDenom <> 0#) Then chkDenom = 1# / chkDenom
        
        srcX = (hA * newX + hB * newY + hC) * chkDenom
        srcY = (hD * newX + hE * newY + hF) * chkDenom
        
        'Check boundaries and assign pixels accordingly
        If (srcX >= 0#) And (srcY >= 0#) Then
            If (srcX < 1#) And (srcY < 1#) Then
                
                'Apply custom foreshortening, if any.
                If (xForeshortening <> 1#) Then srcX = srcX ^ xForeshortening
                If (yForeshortening <> 1#) Then srcY = srcY ^ yForeshortening
                
                'Scale back to the normal image space domain
                srcXInt = Int(srcX * srcWidth)
                srcYInt = Int(srcY * srcHeight)
                
                If (srcImageData(srcXInt, srcYInt).Alpha = 255) Then
                    dstImageData(x) = srcImageData(srcXInt, srcYInt)
                Else
                
                    'Alpha-blend the source pixel against the checkerboard we rendered onto the surface
                    With srcImageData(srcXInt, srcYInt)
                        
                        newAlpha = .Alpha / 255!
                        dstImageData(x).Blue = Int(.Blue) + Int(dstImageData(x).Blue) * (1! - newAlpha)
                        dstImageData(x).Green = Int(.Green) + Int(dstImageData(x).Green) * (1! - newAlpha)
                        dstImageData(x).Red = Int(.Red) + Int(dstImageData(x).Red) * (1! - newAlpha)
                        
                        'Because the destination is guaranteed to only have opaque pixels,
                        ' we don't need to deal with alpha here.
                        
                    End With
                
                End If
                
            Else
                dstImageData(x) = zeroQuad
            End If
        Else
            dstImageData(x) = zeroQuad
        End If
        
    Next x
    Next y
    
    'Safely deallocate all image arrays
    srcDIB.UnwrapRGBQuadArrayFromDIB srcImageData
    dstDIB.UnwrapRGBQuadArrayFromDIB dstImageData
    
End Sub

Private Sub btsSettings_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

Private Sub cboMapping_Click()
    UpdatePreview
    RedrawEditor
End Sub

Private Sub cmdBar_AddCustomPresetData()
    
    'Place all node data into a single string, then write that string out to file
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    Dim i As Long
    For i = 0 To 3
        cParams.AddParam "x" & Trim$(Str$(i)), m_nPoints(i).x
        cParams.AddParam "y" & Trim$(Str$(i)), m_nPoints(i).y
    Next i
    
    cmdBar.AddPresetData "NodeLocations", cParams.GetParamString()
    
End Sub

Private Sub cmdBar_OKClick()
    
    'Free all preview-related DIBs before continuing, since they consume meaningful resource amounts that
    ' are no longer required
    If (Not m_Buffer Is Nothing) Then m_Buffer.EraseDIB True
    If (Not m_Overlay Is Nothing) Then m_Overlay.EraseDIB True
    If (Not m_mouseCoordDIB Is Nothing) Then m_mouseCoordDIB.EraseDIB True
    If (Not m_ProportionalSource Is Nothing) Then m_ProportionalSource.EraseDIB True
    
    Process "Perspective", , GetPerspectiveParamString, UNDO_Layer
    
End Sub

Private Sub cmdBar_RandomizeClick()

    Randomize Timer
    
    'Set the points in the current area to random values - not much to see here!
    Dim i As Long
    For i = 0 To 3
        m_nPoints(i).x = Rnd * picDraw.GetWidth
        m_nPoints(i).y = Rnd * picDraw.GetHeight
    Next i
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    
    'Retrieve the string that contains the node coordinates, and place it into an XML parser
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString cmdBar.RetrievePresetData("NodeLocations")
    
    Dim i As Long
    For i = 0 To 3
        
        'Retrieve this node's x and y values (but only if it exists; otherwise, leave the points where they are)
        If cParams.DoesParamExist("x" & Trim$(Str$(i))) Then
            m_nPoints(i).x = cParams.GetDouble("x" & Trim$(Str$(i)))
            m_nPoints(i).y = cParams.GetDouble("y" & Trim$(Str$(i)))
        End If
        
    Next i
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
    RedrawEditor
End Sub

Private Sub cmdBar_ResetClick()
        
    'Set edge handling to match the default specified in Form_Load
    cboEdges.ListIndex = pdeo_Erase
    
    'Default quality is interpolation, but no supersampling
    sltQuality.Value = 2
    
    'Copy the original values into the "current values" point array and redraw everything
    Dim i As Long
    For i = 0 To 3
        m_nPoints(i).x = m_oPoints(i).x
        m_nPoints(i).y = m_oPoints(i).y
    Next i
        
    UpdatePreview
    RedrawEditor
    
End Sub

Private Sub Form_Load()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Set up surfaces necessary for rendering the live preview
    Set m_Buffer = New pdDIB
    m_Buffer.CreateBlank picDraw.GetWidth, picDraw.GetHeight, 32, 0, 255
    m_Buffer.SetInitialAlphaPremultiplicationState True
    Set m_Overlay = New pdDIB
    m_Overlay.CreateBlank picDraw.GetWidth, picDraw.GetHeight, 32, 0, 0
    m_Overlay.SetInitialAlphaPremultiplicationState True
    CacheSourceImageForPreview
    
    'Initialize the dynamic mouse coordinate font and DIB display
    Set m_mouseCoordDIB = New pdDIB
    Set m_mouseCoordFont = New pdFont
    
    With m_mouseCoordFont
        .SetFontColor RGB(25, 25, 25)
        .SetFontBold True
        .SetFontSize 10
        .CreateFontObject
        .SetTextAlignment vbLeftJustify
    End With
    
    'Disable all previews while we initialize the dialog
    cmdBar.SetPreviewStatus False
    
    'This tool supports two settings panels (basic and advanced, currently)
    btsSettings.AddItem "basic", 0
    btsSettings.AddItem "advanced", 1
    UpdatePanelVisibility
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Erase
    
    'Populate the mapping type combo box
    cboMapping.Clear
    cboMapping.AddItem "forward (outline defines destination area)", 0
    cboMapping.AddItem "reverse (outline defines source area)", 1
    
    'Determine a good scale for the interactive area.  We want the dimensions to fill roughly half the
    ' available area, with aspect ratio preserved, so that the user has plenty of space to move around
    ' the corner nodes.
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = picDraw.GetWidth \ 2
    targetHeight = picDraw.GetHeight \ 2
    
    m_OrigImageWidth = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
    m_OrigImageHeight = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
    
    PDMath.ConvertAspectRatio m_OrigImageWidth, m_OrigImageHeight, targetWidth, targetHeight, m_PreviewWidth, m_PreviewHeight
    
    'Determine initial points for the draw area
    m_oPoints(0).x = (picDraw.GetWidth - m_PreviewWidth) / 2
    m_oPoints(0).y = (picDraw.GetHeight - m_PreviewHeight) / 2
    
    m_oPoints(1).x = m_oPoints(0).x + m_PreviewWidth
    m_oPoints(1).y = m_oPoints(0).y
    
    m_oPoints(2).x = m_oPoints(0).x + m_PreviewWidth
    m_oPoints(2).y = m_oPoints(0).y + m_PreviewHeight
    
    m_oPoints(3).x = m_oPoints(0).x
    m_oPoints(3).y = m_oPoints(0).y + m_PreviewHeight
    
    'Copy those values into the "current values" point array, and back them up to the "original values" array.
    ' (The scale between these two values is passed to the renderer, which allows us to generalize transforms
    ' across image sizes when recorded as part of a macro.)
    Dim i As Long
    For i = 0 To 3
        m_nPoints(i).x = m_oPoints(i).x
        m_nPoints(i).y = m_oPoints(i).y
    Next i
        
    'Mark the mouse as not being down
    m_isMouseDown = False
    m_ActivePoint = -1
    m_HoverPoint = -1
        
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
        
    'Create the preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    RedrawEditor
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Swap between "basic" and "advanced" settings panels
Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = pnlSettings.lBound To pnlSettings.UBound
        pnlSettings(i).Visible = (btsSettings.ListIndex = i)
    Next i
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then PerspectiveImage GetPerspectiveParamString(), True, pdFxPreview
End Sub

Private Sub RedrawEditor()
    
    'Start by clearing the back buffer (using the current theme's backcolor)
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_Buffer
    
    Set cBrush = New pd2DBrush
    If (Not g_Themer Is Nothing) Then cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
    PD2D.FillRectangleI cSurface, cBrush, 0, 0, m_Buffer.GetDIBWidth, m_Buffer.GetDIBHeight
    
    'Next, we want to draw a grid across the buffer.  This helps orient the user as to where the
    ' image's endpoints normally lie.
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenWidth 1!
    cPen.SetPenOpacity 33!
    If (Not g_Themer Is Nothing) Then cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayDefault)
    
    PD2D.DrawLineI cSurface, cPen, 0!, m_Buffer.GetDIBHeight \ 2, m_Buffer.GetDIBWidth, m_Buffer.GetDIBHeight \ 2
    PD2D.DrawLineI cSurface, cPen, m_Buffer.GetDIBWidth \ 2, 0, m_Buffer.GetDIBWidth \ 2, m_Buffer.GetDIBHeight
    
    'Reset opacity before continuing
    cPen.SetPenOpacity 100!
    
    Dim i As Long
    
    'Next, we will do one of two things:
    ' 1) For forward mapping, draw a silhouette and high-speed preview of the "perspective" image,
    '    overlaid by the original image outline.
    ' 2) For reverse mapping, just draw the image itself (since the user needs to target some quad within that image).
    '
    'For case (1), extreme quad orientations may *not* actually render the image (basically, if rendering would
    ' result in the image being "outside" the quad's boundaries), and we track this case so that we can render
    ' additional orientation markers in subsequent steps.
    Dim livePreviewRendered As Boolean: livePreviewRendered = False
    If (cboMapping.ListIndex = 0) Then
        
        'Render a high-speed preview of the transform to a dedicated overlay DIB, then alpha-blend that
        ' against the full interactive area buffer.
        If PDMath.IsPolygonConvex(m_nPoints, 4) Then
            livePreviewRendered = True
            RenderImageToArbitraryQuad m_ProportionalSource, m_Overlay, m_nPoints
            m_Overlay.AlphaBlendToDC m_Buffer.GetDIBDC
        End If
        
        For i = 0 To 3
            
            'For all points but the first, connect point (n) to (n+1)...
            If (i < 3) Then
                PD2D.DrawLineI cSurface, cPen, m_oPoints(i).x, m_oPoints(i).y, m_oPoints(i + 1).x, m_oPoints(i + 1).y
            
            '...and on the last point, reconnect to (0)
            Else
                PD2D.DrawLineI cSurface, cPen, m_oPoints(i).x, m_oPoints(i).y, m_oPoints(0).x, m_oPoints(0).y
            End If
            
        Next i
        
    Else
    
        If cmdBar.PreviewsAllowed Then
            
            Dim tmpSA As SafeArray2D
            EffectPrep.PrepImageData tmpSA, True, pdFxPreview
            
            Dim cSrcSurface As pd2DSurface
            Set cSrcSurface = New pd2DSurface
            cSrcSurface.WrapSurfaceAroundPDDIB workingDIB
            
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_Buffer, m_oPoints(0).x, m_oPoints(0).y, m_oPoints(1).x - m_oPoints(0).x, m_oPoints(2).y - m_oPoints(0).y, g_CheckerboardPattern, fixBoundaryPainting:=True, noAntialiasing:=True
            PD2D.DrawSurfaceResizedCroppedF cSurface, m_oPoints(0).x, m_oPoints(0).y, m_oPoints(1).x - m_oPoints(0).x, m_oPoints(2).y - m_oPoints(0).y, cSrcSurface, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
            Set cSrcSurface = Nothing
            
        End If
        
    End If
    
    'Next, draw connecting lines to form an image outline.
    cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    
    If (Not g_Themer Is Nothing) Then cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
    cPen.SetPenWidth 1.6!
    
    For i = 0 To 3
        If (i < 3) Then
            PD2D.DrawLineF cSurface, cPen, m_nPoints(i).x, m_nPoints(i).y, m_nPoints(i + 1).x, m_nPoints(i + 1).y
        Else
            PD2D.DrawLineF cSurface, cPen, m_nPoints(i).x, m_nPoints(i).y, m_nPoints(0).x, m_nPoints(0).y
        End If
    Next i
    
    'Next, draw circles at the corners of the perspective area, and hover the currently active node (if any)
    Dim targetColor As Long
    Dim clrActive As Long, clrDefault As Long
    If (Not g_Themer Is Nothing) Then
        clrDefault = g_Themer.GetGenericUIColor(UI_Accent, True)
        clrActive = g_Themer.GetGenericUIColor(UI_AccentDark, True)
    End If
    
    Dim targetThickness As Single, defaultThickness As Single, hoverThickness As Single
    defaultThickness = 2!
    hoverThickness = 3.5!
    
    For i = 0 To 3
        
        If ((i = m_ActivePoint) Or (i = m_HoverPoint)) Then targetColor = clrActive Else targetColor = clrDefault
        cPen.SetPenColor targetColor
        
        'When hovered/active, draw a larger point
        If ((i = m_ActivePoint) Or (i = m_HoverPoint)) Then targetThickness = hoverThickness Else targetThickness = defaultThickness
        cPen.SetPenWidth targetThickness
        
        'Draw the circle
        PD2D.DrawCircleF cSurface, cPen, m_nPoints(i).x, m_nPoints(i).y, Interface.FixDPIFloat(7)
        
    Next i
    
    'Finally, draw the center cross to help the user orient to the center point of the perspective effect.
    ' (Note that we only do this if a full preview was *not* rendered.)
    If (Not livePreviewRendered) Then
        If (Not g_Themer Is Nothing) Then cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
        cPen.SetPenWidth 1!
        cPen.SetPenOpacity 50!
        PD2D.DrawLineF cSurface, cPen, m_nPoints(0).x, m_nPoints(0).y, m_nPoints(2).x, m_nPoints(2).y
        PD2D.DrawLineF cSurface, cPen, m_nPoints(1).x, m_nPoints(1).y, m_nPoints(3).x, m_nPoints(3).y
    End If
    
    'Before exiting, draw a border around the "finished" interactive buffer.
    If (Not g_Themer Is Nothing) Then cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayDark)
    cPen.SetPenWidth 1!
    cPen.SetPenLineJoin P2_LJ_Miter
    PD2D.DrawRectangleI cSurface, cPen, 0, 0, m_Buffer.GetDIBWidth - 1, m_Buffer.GetDIBHeight - 1
    
    'Finally, if a node is hovered, display a live coordinate overlay for that coordinate *in layer space*.
    If (m_HoverPoint >= 0) Then
    
        'Generate two strings: the "name" of this point (e.g. top-right) and the coordinate of this point.
        ' We generate these separately, because we want to calculate width independently for each string,
        ' and use the larger of the two as our bounding rect for the coordinate overlay.
        Dim strFinal As String, strName As String, strCoord As String
        
        'Note that coordActualX/Y are calculated by the param string generator - we rely on its work
        ' for these values!
        strCoord = "(" & Int(m_dstLayerSpacePoints(m_HoverPoint).x + 0.5) & ", " & Int(m_dstLayerSpacePoints(m_HoverPoint).y + 0.5) & ")"
        
        'The "name" of the selected point is a fixed string
        Select Case m_HoverPoint
            Case 0
                strName = g_Language.TranslateMessage("top-left")
            Case 1
                strName = g_Language.TranslateMessage("top-right")
            Case 2
                strName = g_Language.TranslateMessage("bottom-right")
            Case 3
                strName = g_Language.TranslateMessage("bottom-left")
        End Select
        
        'Find the larger of the two strings
        Dim maxStringWidth As Long
        maxStringWidth = m_mouseCoordFont.GetWidthOfString(strName)
        If (m_mouseCoordFont.GetWidthOfString(strCoord) > maxStringWidth) Then maxStringWidth = m_mouseCoordFont.GetWidthOfString(strCoord)
        
        'Concatenate the two strings
        strFinal = strName & vbCrLf & strCoord
        
        'Calculate the size of the concatenated input/output string (in pixels, both width and height, with the width limited
        ' to the larger of the original two strings)
        Dim strFinalWidth As Long, strFinalHeight As Long
        strFinalWidth = maxStringWidth
        strFinalHeight = m_mouseCoordFont.GetHeightOfWordwrapString(strFinal, strFinalWidth + 1)
        
        'Create a new DIB at the size of the string (with a slight bit of padding on all sides)
        Dim coordBoxWidth As Long, coordBoxHeight As Long
        coordBoxWidth = strFinalWidth + Interface.FixDPI(8)
        coordBoxHeight = strFinalHeight + Interface.FixDPI(5)
        
        'Normally we would never want to (knowingly) create a 24-bpp DIB, but GDI font rendering is broken
        ' on 32-bpp targets so we *must* use 24-bpp here
        If (m_mouseCoordDIB Is Nothing) Then Set m_mouseCoordDIB = New pdDIB
        m_mouseCoordDIB.CreateBlank coordBoxWidth, coordBoxHeight, 24, vbWhite
        m_mouseCoordDIB.SetInitialAlphaPremultiplicationState True
        
        'Render the coordinate string onto the temporary DIB
        m_mouseCoordFont.AttachToDC m_mouseCoordDIB.GetDIBDC
        m_mouseCoordFont.FastRenderMultilineText Interface.FixDPI(4), Interface.FixDPI(2), strFinal
        m_mouseCoordFont.ReleaseFromDC
        
        'Render a 1px border around the coordinate overlay
        cSurface.WrapSurfaceAroundPDDIB m_mouseCoordDIB
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        cPen.SetPenColor vbBlack
        cPen.SetPenOpacity 100!
        cPen.SetPenWidth 1!
        cPen.SetPenLineJoin P2_LJ_Miter
        PD2D.DrawRectangleI cSurface, cPen, 0, 0, m_mouseCoordDIB.GetDIBWidth - 1, m_mouseCoordDIB.GetDIBHeight - 1
        
        'Calculate render coordinates for the coordinate box.  These vary according to point, but generally
        ' we want to be on the "outside corner" of the given point.
        Dim boxPadding As Long
        boxPadding = Interface.FixDPI(12)
        
        Dim startPoint As PointFloat
        startPoint = m_nPoints(m_HoverPoint)
        
        Dim coordX As Long, coordY As Long
        Select Case m_HoverPoint
        
            'top-left
            Case 0
                coordX = startPoint.x - (m_mouseCoordDIB.GetDIBWidth + boxPadding)
                coordY = startPoint.y - (m_mouseCoordDIB.GetDIBHeight + boxPadding)
                
            'top-right
            Case 1
                coordX = startPoint.x + boxPadding
                coordY = startPoint.y - (m_mouseCoordDIB.GetDIBHeight + boxPadding)
                
            'bottom-right
            Case 2
                coordX = startPoint.x + boxPadding
                coordY = startPoint.y + boxPadding
                
            'bottom-left
            Case 3
                coordX = startPoint.x - (m_mouseCoordDIB.GetDIBWidth + boxPadding)
                coordY = startPoint.y + boxPadding
                
        End Select
        
        'Fit the final coordinates in-bounds
        If (coordX < 0) Then coordX = 0
        If (coordY < 0) Then coordY = 0
        If (coordX + m_mouseCoordDIB.GetDIBWidth > picDraw.GetWidth) Then coordX = picDraw.GetWidth - m_mouseCoordDIB.GetDIBWidth
        If (coordY + m_mouseCoordDIB.GetDIBHeight > picDraw.GetHeight) Then coordY = picDraw.GetHeight - m_mouseCoordDIB.GetDIBHeight
        
        'Render the completed coordinate overlay DIB onto the main interactive box
        m_mouseCoordDIB.AlphaBlendToDC m_Buffer.GetDIBDC, 192, coordX, coordY
        
    End If
    
    'Flip the completed buffer to the screen
    Set cSurface = Nothing
    picDraw.RequestRedraw True
    
    'Finally, sync the text boxes to the current corner dimensions
    m_SuspendSync = True
    
    For i = 0 To 3
        spnCoords(i * 2).Value = m_dstLayerSpacePoints(i).x
        VBHacks.DoEvents_SingleHwnd spnCoords(i * 2).hWnd
        spnCoords(i * 2 + 1).Value = m_dstLayerSpacePoints(i).y
        VBHacks.DoEvents_SingleHwnd spnCoords(i * 2 + 1).hWnd
    Next i
    
    m_SuspendSync = False
    
End Sub

'Simple distance routine to see if a location on the picture box is near an existing point
Private Function CheckClick(ByVal x As Long, ByVal y As Long) As Long
    
    'Returning -1 says we're not close to an existing point
    CheckClick = -1
    
    Dim i As Long
    Dim dist As Double, bestDist As Double, bestIndex As Long
    bestDist = DOUBLE_MAX
    bestIndex = -1
    
    For i = 0 To 3
    
        dist = PDMath.DistanceTwoPoints(x, y, m_nPoints(i).x, m_nPoints(i).y)
        If (dist < bestDist) Then
            bestDist = dist
            bestIndex = i
        End If
        
    Next i
    
    'If we're close to an existing point, return the index of that point
    If (bestDist < Interface.GetStandardInteractionDistance()) Then CheckClick = bestIndex
    
End Function

'Is a point (from the interactive area) inside the perspective-corrected image's quad?
Private Function IsPointInQuad(ByVal x As Long, ByVal y As Long) As Boolean
    
    'Build a path from the current list of points
    Dim cShape As pd2DPath
    Set cShape = New pd2DPath
    cShape.AddPolygon 4, VarPtr(m_nPoints(0)), True, False
    
    'Let the path do the rest!
    IsPointInQuad = cShape.IsPointInsidePathL(x, y)
    
End Function

'Take the current tool settings and merge them into a parameter string
Private Function GetPerspectiveParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'First, add the four corner points of the transform.  Note that these have to be mapped from the current UI
    ' coordinate space to an absolute coordinate space (suitable for storage in screen-independent places like macros).
    Dim xModifier As Double, yModifier As Double
    If (m_PreviewWidth <> 0#) Then xModifier = (m_OrigImageWidth / m_PreviewWidth) Else xModifier = 1#
    If (m_PreviewHeight <> 0#) Then yModifier = (m_OrigImageHeight / m_PreviewHeight) Else yModifier = 1#
    
    Dim tmpX As Double, tmpY As Double
    
    'Top-left
    tmpX = (m_nPoints(0).x - m_oPoints(0).x) * xModifier
    tmpY = (m_nPoints(0).y - m_oPoints(0).y) * yModifier
    cParams.AddParam "topleftx", tmpX
    cParams.AddParam "toplefty", tmpY
    m_dstLayerSpacePoints(0).x = tmpX
    m_dstLayerSpacePoints(0).y = tmpY
    
    'Top-right
    tmpX = m_OrigImageWidth + ((m_nPoints(1).x - m_oPoints(1).x) * xModifier)
    tmpY = (m_nPoints(1).y - m_oPoints(1).y) * yModifier
    cParams.AddParam "toprightx", tmpX
    cParams.AddParam "toprighty", tmpY
    m_dstLayerSpacePoints(1).x = tmpX
    m_dstLayerSpacePoints(1).y = tmpY
    
    'Bottom-right
    tmpX = m_OrigImageWidth + ((m_nPoints(2).x - m_oPoints(2).x) * xModifier)
    tmpY = m_OrigImageHeight + (m_nPoints(2).y - m_oPoints(2).y) * yModifier
    cParams.AddParam "bottomrightx", tmpX
    cParams.AddParam "bottomrighty", tmpY
    m_dstLayerSpacePoints(2).x = tmpX
    m_dstLayerSpacePoints(2).y = tmpY
    
    'Bottom-left
    tmpX = (m_nPoints(3).x - m_oPoints(3).x) * xModifier
    tmpY = m_OrigImageHeight + (m_nPoints(3).y - m_oPoints(3).y) * yModifier
    cParams.AddParam "bottomleftx", tmpX
    cParams.AddParam "bottomlefty", tmpY
    m_dstLayerSpacePoints(3).x = tmpX
    m_dstLayerSpacePoints(3).y = tmpY
    
    'Next, note the type of mapping (quadrilateral to square, or square to quadrilateral)
    cParams.AddParam "mapping", cboMapping.ListIndex
    
    'Custom foreshortening was added in 9.2
    cParams.AddParam "x-foreshorten", sldForeshortening(0).Value, True
    cParams.AddParam "y-foreshorten", sldForeshortening(1).Value, True
    
    'Finally, quality and supersampling settings
    cParams.AddParam "edges", cboEdges.ListIndex
    cParams.AddParam "quality", sltQuality.Value
    
    GetPerspectiveParamString = cParams.GetParamString()
    
End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub picDraw_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.BitBltWrapper targetDC, 0, 0, m_Buffer.GetDIBWidth, m_Buffer.GetDIBHeight, m_Buffer.GetDIBDC, 0, 0, vbSrcCopy
End Sub

Private Sub picDraw_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    m_isMouseDown = True
    
    'If the mouse is over a point, mark it as the active point
    m_ActivePoint = CheckClick(x, y)
    
    'If the user *didn't* click a point, look for move operations (in the interior of the quad)
    If (m_ActivePoint < 0) Then
    
        m_MoveActive = IsPointInQuad(x, y)
        
        If m_MoveActive Then
            
            m_InitPoint.x = x
            m_InitPoint.y = y
            
            Dim i As Long
            For i = 0 To 3
                m_PointsAtMoveStart(i) = m_nPoints(i)
            Next i
            
        End If
            
    End If
    
End Sub

Private Sub picDraw_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If the mouse is not down, indicate to the user that points (and the shape itself) can be moved
    If (Not m_isMouseDown) Then
        
        Dim origHoverPoint As Long
        origHoverPoint = m_HoverPoint
        
        'If the user is close to a knot, change the mousepointer to 'move'
        m_HoverPoint = CheckClick(x, y)
        If (m_HoverPoint >= 0) Then
            picDraw.RequestCursor IDC_HAND
        Else
            
            'If the cursor is inside the quadrilateral, allow the user to move *all* points simultaneously
            If IsPointInQuad(x, y) Then
                picDraw.RequestCursor IDC_SIZEALL
            Else
                picDraw.RequestCursor IDC_DEFAULT
            End If
            
            picDraw.AssignTooltip vbNullString, raiseTipsImmediately:=False
            
        End If
        
        If (origHoverPoint <> m_HoverPoint) Then RedrawEditor
    
    'If the mouse is down, move the current point and redraw the preview
    Else
        
        'Mirror the hover point to match the active point
        m_HoverPoint = m_ActivePoint
        
        'If the user is dragging a corner node, update that node's position and redraw accordingly
        If (m_ActivePoint >= 0) Then
        
            m_nPoints(m_ActivePoint).x = x
            m_nPoints(m_ActivePoint).y = y
            UpdatePreview
            RedrawEditor
        
        'Similarly, if the user is moving the entire quad, update *all* node positions
        ElseIf m_MoveActive Then
            
            Dim i As Long
            For i = 0 To 3
                m_nPoints(i).x = m_PointsAtMoveStart(i).x + (x - m_InitPoint.x)
                m_nPoints(i).y = m_PointsAtMoveStart(i).y + (y - m_InitPoint.y)
            Next i
            
            UpdatePreview
            RedrawEditor
        
        End If
        
    End If

End Sub

Private Sub picDraw_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_isMouseDown = False
    m_MoveActive = False
    m_ActivePoint = -1
End Sub

Private Sub picDraw_Resize(ByVal newWidth As Long, ByVal newHeight As Long)
    If (Not m_Buffer Is Nothing) And PDMain.IsProgramRunning() Then
        m_Buffer.CreateBlank newWidth, newHeight, 32, 0, 255
        m_Buffer.SetInitialAlphaPremultiplicationState True
        m_Overlay.CreateBlank newWidth, newHeight, 32, 0, 0
        m_Overlay.SetInitialAlphaPremultiplicationState True
        CacheSourceImageForPreview
        RedrawEditor
    End If
End Sub

Private Sub sldForeshortening_Change(Index As Integer)
    RedrawEditor
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

'Scale the source layer (which is potentially enormous) to a maximum size of the interactive area;
' this lets us perform faster mapping for the on-screen effect preview
Private Sub CacheSourceImageForPreview()
    
    If (m_ProportionalSource Is Nothing) Then Set m_ProportionalSource = New pdDIB
    
    Dim newWidth As Long, newHeight As Long
    With PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB
        PDMath.ConvertAspectRatio .GetDIBWidth, .GetDIBHeight, picDraw.GetWidth, picDraw.GetHeight, newWidth, newHeight
    End With
    
    m_ProportionalSource.CreateBlank newWidth, newHeight, 32, 0, 0
    GDI_Plus.GDIPlus_StretchBlt m_ProportionalSource, 0, 0, newWidth, newHeight, PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB, 0, 0, PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetDIBWidth, PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetDIBHeight, dstCopyIsOkay:=True
    m_ProportionalSource.SetInitialAlphaPremultiplicationState PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetAlphaPremultiplication()
    
End Sub

Private Sub spnCoords_BeforeResetClick(Index As Integer)
    m_SuspendSync = True
End Sub

Private Sub spnCoords_Change(Index As Integer)
    
    'Prevent recursive setting propagation
    If (Not m_SuspendSync) Then ReflectNewTextChanges Index
    
End Sub

Private Sub spnCoords_ResetClick(Index As Integer)
    
    'Assign x-values first, and *do not* refresh the screen until the second value is set
    m_SuspendSync = True
    
    Select Case Index
        Case 1
            spnCoords(0).Value = 0
            spnCoords(1).Value = 0
        Case 3
            spnCoords(2).Value = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
            spnCoords(3).Value = 0
        Case 5
            spnCoords(4).Value = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
            spnCoords(5).Value = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
        Case 7
            spnCoords(6).Value = 0
            spnCoords(7).Value = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
    End Select
    
    m_SuspendSync = False
    ReflectNewTextChanges Index
    
End Sub

Private Sub ReflectNewTextChanges(ByVal Index As Long)
    
    'Any changes to text boxes require us to mirror said changes to the m_nPoints() array.
    Dim origCoords As PointFloat
    Select Case (Index \ 2)
        Case 0
            origCoords.x = 0
            origCoords.y = 0
        Case 1
            origCoords.x = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
            origCoords.y = 0
        Case 2
            origCoords.x = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
            origCoords.y = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
        Case 3
            origCoords.x = 0
            origCoords.y = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
    End Select
    
    Dim xModifier As Double, yModifier As Double
    If (m_PreviewWidth <> 0#) Then xModifier = (m_OrigImageWidth / m_PreviewWidth) Else xModifier = 1#
    If (m_PreviewHeight <> 0#) Then yModifier = (m_OrigImageHeight / m_PreviewHeight) Else yModifier = 1#
    
    Dim spnIndexX As Long, spnIndexY As Long
    spnIndexX = (Index \ 2) * 2
    spnIndexY = spnIndexX + 1
    m_nPoints(Index \ 2).x = m_oPoints(Index \ 2).x + (spnCoords(spnIndexX).Value - origCoords.x) / xModifier
    m_nPoints(Index \ 2).y = m_oPoints(Index \ 2).y + (spnCoords(spnIndexY).Value - origCoords.y) / yModifier
    
    UpdatePreview
    RedrawEditor
    
End Sub

Private Function TranslateForeshorteningUIValue(ByVal srcValue As Double) As Double
    If (srcValue >= 0#) Then
        TranslateForeshorteningUIValue = srcValue + 1#
    Else
        TranslateForeshorteningUIValue = (srcValue + (sldForeshortening(0).Max + 1#)) / (sldForeshortening(0).Max + 1#)
    End If
End Function
