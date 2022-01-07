VERSION 5.00
Begin VB.Form FormPerspective 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Perspective"
   ClientHeight    =   9060
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
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   8040
      Left            =   6000
      ScaleHeight     =   534
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   534
      TabIndex        =   4
      Top             =   120
      Width           =   8040
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   120
      TabIndex        =   3
      Top             =   6660
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
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
   Begin PhotoDemon.pdDropDown cboMapping 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5850
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "transformation type"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   8310
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormPerspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Perspective Distortion
'Copyright 2013-2022 by Tanner Helland
'Created: 08/April/13
'Last updated: 19/February/20
'Last update: large performance improvements, overhaul UI to properly support theming
'
'This tool allows the user to apply arbitrary perspective to an image.  The code is fairly involved linear
' algebra, as a series of equations must be solved to generate the homography matrix used for the transform.
' For a more detailed explanation of the math and theory behind projective transforms, please visit:
'
' https://en.wikipedia.org/wiki/Homography
'
'As with all distorts, reverse-mapping plus supersampling is supported for high-quality antialiasing.
'
'I used a number of projects as references while build this tool.  Thank you to the following:
'
' http://www.cs.cmu.edu/~ph/texfund/texfund.pdf
' http://www.imagemagick.org/Usage/distorts/#perspective
' http://stackoverflow.com/questions/169902/projective-transformation
' http://freespace.virgin.net/hugo.elias/graphics/x_persp.htm
' http://stackoverflow.com/questions/530396/how-to-draw-a-perspective-correct-grid-in-2d?lq=1
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

'Mouse status is tracked between MouseDown and MouseMove events; this allows for drag events.
Private m_isMouseDown As Boolean

'Currently selected and hovered nodes in the workspace area (if any); -1 = no point selected
Private m_ActivePoint As Long, m_HoverPoint As Long

'Buffer to which the current interactive "perspective" control is rendered.
Private m_Buffer As pdDIB

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

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
    imgWidth = finalX - initX
    imgHeight = finalY - initY
    
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
    
        'Invert the transformation using the adjoint of the forward mapping.  If you don't know what an adjoint is
        ' (don't worry, most don't! :), we're basically reversing the plane-to-plane mapping by which we've defined
        ' this particular projection.  (This means that we want the quadrilateral to define a section of the SOURCE
        ' image instead of a section of the DESTINATION image.)  For a detailed explanation of this process, please
        ' read pages 24-25 of Paul Heckbert's thesis on projective transformations, which is IMO a great source for
        ' understanding projective mappings in general: http://www.cs.cmu.edu/~ph/texfund/texfund.pdf
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
            
            srcX = imgWidth * (hA * newX + hB * newY + hC) * chkDenom
            srcY = imgHeight * (hD * newX + hE * newY + hF) * chkDenom
            
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            tmpQuad = fSupport.GetColorsFromSource(srcX, srcY, x, y)
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

Private Sub cboMapping_Click()
    RedrawEditor
    UpdatePreview
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
    Process "Perspective", , GetPerspectiveParamString, UNDO_Layer
End Sub

Private Sub cmdBar_RandomizeClick()

    Randomize Timer
    
    'Set the points in the current area to random values - not much to see here!
    Dim i As Long
    For i = 0 To 3
        m_nPoints(i).x = Rnd * picDraw.ScaleWidth
        m_nPoints(i).y = Rnd * picDraw.ScaleHeight
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
    RedrawEditor
    UpdatePreview
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
        
    RedrawEditor
    UpdatePreview
    
End Sub

Private Sub Form_Load()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    Set m_Buffer = New pdDIB
    m_Buffer.CreateBlank picDraw.ScaleWidth, picDraw.ScaleHeight, 32, 0, 255
    
    'Disable all previews while we initialize the dialog
    cmdBar.SetPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Erase
    
    'Populate the mapping type combo box
    cboMapping.Clear
    cboMapping.AddItem "forward (outline defines destination area)", 0
    cboMapping.AddItem "reverse (outline defines source area)", 1
    
    'Note the current image's width and height, which is needed to map between the on-screen interactive UI area,
    ' and the final transform.
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, True, pdFxPreview, , , True
    m_PreviewWidth = curDIBValues.Width
    m_PreviewHeight = curDIBValues.Height
    m_OrigImageWidth = curDIBValues.Width / curDIBValues.previewModifier
    m_OrigImageHeight = curDIBValues.Height / curDIBValues.previewModifier
    
    'Determine initial points for the draw area
    m_oPoints(0).x = (picDraw.ScaleWidth - m_PreviewWidth) / 2
    m_oPoints(0).y = (picDraw.ScaleHeight - m_PreviewHeight) / 2
    
    m_oPoints(1).x = m_oPoints(0).x + m_PreviewWidth
    m_oPoints(1).y = m_oPoints(0).y
    
    m_oPoints(2).x = m_oPoints(0).x + m_PreviewWidth
    m_oPoints(2).y = m_oPoints(0).y + m_PreviewHeight
    
    m_oPoints(3).x = m_oPoints(0).x
    m_oPoints(3).y = m_oPoints(0).y + m_PreviewHeight
    
    'Copy those values into the "current values" point array
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
    RedrawEditor
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then PerspectiveImage GetPerspectiveParamString, True, pdFxPreview
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
    PD2D.DrawLineI cSurface, cPen, m_Buffer.GetDIBHeight \ 2, 0, m_Buffer.GetDIBWidth \ 2, m_Buffer.GetDIBHeight
    
    'Next, we will do one of two things:
    ' 1) For forward mapping, draw a silhouette around the original image outline.
    ' 2) For reverse mapping, just draw the image itself.
    If (cboMapping.ListIndex = 0) Then
        
        cPen.SetPenOpacity 100!
        
        Dim i As Long
        For i = 0 To 3
            
            'For all points but the first, connect point (n) to (n+1); on the last point,
            ' reconnect to (0)
            If (i < 3) Then
                PD2D.DrawLineI cSurface, cPen, m_oPoints(i).x, m_oPoints(i).y, m_oPoints(i + 1).x, m_oPoints(i + 1).y
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
            
            PD2D.DrawSurfaceResizedCroppedF cSurface, m_oPoints(0).x, m_oPoints(0).y, m_oPoints(1).x - m_oPoints(0).x, m_oPoints(2).y - m_oPoints(0).y, cSrcSurface, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
            
        End If
        
    End If
    
    'Next, draw connecting lines to form an image outline.  Use GDI+ for superior results (e.g. antialiasing).
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
    
    'Finally, draw the center cross to help the user orient to the center point of the perspective effect
    If (Not g_Themer Is Nothing) Then cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
    cPen.SetPenWidth 1!
    cPen.SetPenOpacity 50!
    PD2D.DrawLineF cSurface, cPen, m_nPoints(0).x, m_nPoints(0).y, m_nPoints(2).x, m_nPoints(2).y
    PD2D.DrawLineF cSurface, cPen, m_nPoints(1).x, m_nPoints(1).y, m_nPoints(3).x, m_nPoints(3).y
    
    'Flip the completed buffer to the screen
    Set cSurface = Nothing
    GDI.BitBltWrapper picDraw.hDC, 0, 0, m_Buffer.GetDIBWidth, m_Buffer.GetDIBHeight, m_Buffer.GetDIBDC, 0, 0, vbSrcCopy
    Set picDraw.Picture = picDraw.Image
    picDraw.Refresh

End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    m_isMouseDown = True
    
    'If the mouse is over a point, mark it as the active point
    m_ActivePoint = CheckClick(x, y)
    
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the mouse is not down, indicate to the user that points can be moved
    If (Not m_isMouseDown) Then
        
        Dim origHoverPoint As Long
        origHoverPoint = m_HoverPoint
        
        'If the user is close to a knot, change the mousepointer to 'move'
        m_HoverPoint = CheckClick(x, y)
        If (m_HoverPoint >= 0) Then
        
            If (picDraw.MousePointer <> vbSizePointer) Then picDraw.MousePointer = vbSizePointer
            
            'Display a tool-tip (helpful when the user wants a severely distorted transform matrix)
            Select Case CheckClick(x, y)
                Case 0
                    picDraw.ToolTipText = g_Language.TranslateMessage("top-left")
                Case 1
                    picDraw.ToolTipText = g_Language.TranslateMessage("top-right")
                Case 2
                    picDraw.ToolTipText = g_Language.TranslateMessage("bottom-right")
                Case 3
                    picDraw.ToolTipText = g_Language.TranslateMessage("bottom-left")
                    
            End Select
            
        Else
            If (picDraw.MousePointer <> vbDefault) Then picDraw.MousePointer = vbDefault
            picDraw.ToolTipText = vbNullString
        End If
        
        If (origHoverPoint <> m_HoverPoint) Then RedrawEditor
    
    'If the mouse is down, move the current point and redraw the preview
    Else
        
        m_HoverPoint = m_ActivePoint
        
        If (m_ActivePoint >= 0) Then
            m_nPoints(m_ActivePoint).x = x
            m_nPoints(m_ActivePoint).y = y
            RedrawEditor
            UpdatePreview
        End If
    
    End If

End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_isMouseDown = False
    m_ActivePoint = -1
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
    
    'Top-right
    tmpX = m_OrigImageWidth + ((m_nPoints(1).x - m_oPoints(1).x) * xModifier)
    tmpY = (m_nPoints(1).y - m_oPoints(1).y) * yModifier
    cParams.AddParam "toprightx", tmpX
    cParams.AddParam "toprighty", tmpY
    
    'Bottom-right
    tmpX = m_OrigImageWidth + ((m_nPoints(2).x - m_oPoints(2).x) * xModifier)
    tmpY = m_OrigImageHeight + (m_nPoints(2).y - m_oPoints(2).y) * yModifier
    cParams.AddParam "bottomrightx", tmpX
    cParams.AddParam "bottomrighty", tmpY
    
    'Bottom-left
    tmpX = (m_nPoints(3).x - m_oPoints(3).x) * xModifier
    tmpY = m_OrigImageHeight + (m_nPoints(3).y - m_oPoints(3).y) * yModifier
    cParams.AddParam "bottomleftx", tmpX
    cParams.AddParam "bottomlefty", tmpY
    
    'Next, note the type of mapping (quadrilateral to square, or square to quadrilateral)
    cParams.AddParam "mapping", cboMapping.ListIndex
    
    'Finally, quality and supersampling settings
    cParams.AddParam "edges", cboEdges.ListIndex
    cParams.AddParam "quality", sltQuality.Value
    
    GetPerspectiveParamString = cParams.GetParamString()
    
End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub picDraw_Resize()
    If (Not m_Buffer Is Nothing) Then
        m_Buffer.CreateBlank picDraw.ScaleWidth, picDraw.ScaleHeight, 32, 0, 255
        RedrawEditor
    End If
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub
