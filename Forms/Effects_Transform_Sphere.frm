VERSION 5.00
Begin VB.Form FormSpherize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Spherize"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
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
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   12105
      _ExtentX        =   21352
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
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -180
      Max             =   180
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltOffsetY 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "vertical offset"
      Min             =   -100
      Max             =   100
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltOffsetX 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "horizontal offset"
      Min             =   -100
      Max             =   100
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   7
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "quality"
      Min             =   1
      Max             =   5
      Value           =   2
      NotchPosition   =   2
      NotchValueCustom=   2
   End
   Begin PhotoDemon.pdButtonStrip btsExterior 
      Height          =   1080
      Left            =   6000
      TabIndex        =   2
      Top             =   3840
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1905
      Caption         =   "area outside sphere"
   End
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   5160
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
End
Attribute VB_Name = "FormSpherize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Spherize" Distortion
'Copyright 2013-2026 by Tanner Helland
'Created: 05/June/13
'Last updated: 20/February/20
'Last update: large performance improvements
'
'This tool allows the user to map an image around a sphere, with optional rotation and bidirectional offsets.
' Supersampling and reverse-mapped interpolation are available to improve mapping quality.
'
'In keeping with the true mathematical nature of spheres, this tool forces x and y mapping to the minimum
' dimension (width or height).  Technically this isn't necessary, as the code works just fine with rectangular
' shapes, but as my lens distort tool already handles rectangular shapes, I'm forcing this one to spheres only.
'
'For a bit of extra fun, the empty space around the sphere can be mapped one of two ways - as light rays "beaming"
' from behind the sphere, or simply erased to white.  Your choice.
'
'The transformation used by this tool is a heavily modified version of an equation originally shared by Paul Bourke.
' You can see Paul's original (and very helpful article) at the following link, good as of 05 June '13:
' http://paulbourke.net/miscellaneous/imagewarp/
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsExterior_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Apply a "swirl" effect to an image
Public Sub SpherizeImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim sphereAngle As Double, xOffset As Double, yOffset As Double
    Dim useRays As Boolean, edgeHandling As Long, superSamplingAmount As Long
    
    With cParams
        sphereAngle = .GetDouble("angle", sltAngle.Value)
        xOffset = .GetDouble("xoffset", 0#)
        yOffset = .GetDouble("yoffset", 0#)
        useRays = .GetBool("rays", False)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        superSamplingAmount = .GetLong("quality", sltQuality.Value)
    End With
    
    'Reverse the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    ' Also, convert it to radians.
    sphereAngle = sphereAngle * (PI / 180#)

    If (Not toPreview) Then Message "Wrapping image around sphere..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image,
    ' and we will use it as our source reference.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters edgeHandling, (superSamplingAmount <> 1), curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    '***************************************
    ' /* BEGIN SUPERSAMPLING PREPARATION */
    
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
    Dim j As Double, k As Double
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
    
    'Sphere transformations require a number of specialized variables
    
    'Polar conversion values
    Dim theta As Double, radius As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    
    Dim minDimension As Long
    Dim minDimVertical As Boolean
    
    If (tWidth < tHeight) Then
        minDimension = tWidth
        minDimVertical = False
    Else
        minDimension = tHeight
        minDimVertical = True
    End If
        
    Dim halfDimDiff As Double
    halfDimDiff = Abs(tWidth - tHeight) / 2#
    
    'Convert offsets to usable amounts
    xOffset = (xOffset / 100#) * tWidth
    yOffset = (yOffset / 100#) * tHeight
    
    'A slow part of this algorithm is translating all (x, y) coordinates to the range [-1, 1].  We can perform this
    ' step in advance by constructing a look-up table for all possible x values and all possible y values.  This
    ' greatly improves performance.
    Dim xLookup() As Double, yLookup() As Double
    ReDim xLookup(initX To finalX) As Double
    ReDim yLookup(initY To finalY) As Double
    
    For x = initX To finalX
        If minDimVertical Then
            xLookup(x) = (2# * (x - halfDimDiff)) / minDimension - 1#
        Else
            xLookup(x) = (2# * x) / minDimension - 1#
        End If
    Next x
    
    For y = initY To finalY
        If minDimVertical Then
            yLookup(y) = (2# * y) / minDimension - 1#
        Else
            yLookup(y) = (2# * (y - halfDimDiff)) / minDimension - 1#
        End If
    Next y
    
    'Do the same thing for our supersampling coordinates
    For sampleIndex = 0 To numSamples
        ssX(sampleIndex) = ssX(sampleIndex) / minDimension 'tWidth
        ssY(sampleIndex) = ssY(sampleIndex) / minDimension 'tHeight
    Next sampleIndex
    
    'We can also calculate a few constants in advance
    Dim twoDivByPI As Double
    twoDivByPI = 2# / PI
    
    Dim halfMinDimension As Double
    halfMinDimension = minDimension / 2#
    
    Dim tmpQuad As RGBQuad
    fSupport.AliasTargetDIB srcDIB
          
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
        
        'Reset all supersampling values
        newR = 0
        newG = 0
        newB = 0
        newA = 0
        numSamplesUsed = 0
        
        'Remap the coordinates around a center point of (0, 0), and normalize them to (-1, 1)
        j = xLookup(x)
        k = yLookup(y)
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
        
            'Offset the pixel amount by the supersampling lookup table
            nX = j + ssX(sampleIndex)
            nY = k + ssY(sampleIndex)
            
            'Next, map them to polar coordinates and apply the spherification
            radius = Sqr(nX * nX + nY * nY)
            theta = PDMath.Atan2(nY, nX)
            
            radius = Asin(radius) * twoDivByPI
            
            'Apply optional rotation
            theta = theta - sphereAngle
                    
            'Convert them back to the Cartesian plane
            nX = radius * Cos(theta) + 1#
            srcX = halfMinDimension * nX + xOffset
            nY = radius * Sin(theta) + 1#
            srcY = halfMinDimension * nY + yOffset
            
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            If useRays Then
                tmpQuad = fSupport.GetColorsFromSource(srcX, srcY, x, y)
                b = tmpQuad.Blue
                g = tmpQuad.Green
                r = tmpQuad.Red
                a = tmpQuad.Alpha
            Else
                If (radius <= 1) Then
                    tmpQuad = fSupport.GetColorsFromSource(srcX, srcY, x, y)
                    b = tmpQuad.Blue
                    g = tmpQuad.Green
                    r = tmpQuad.Red
                    a = tmpQuad.Alpha
                Else
                    r = 0
                    g = 0
                    b = 0
                    a = 0
                End If
            End If
            
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

'OK button
Private Sub cmdBar_OKClick()
    Process "Spherize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltQuality = 2
    cboEdges.ListIndex = pdeo_Wrap
End Sub

Private Sub Form_Load()
    
    'Don't attempt to preview the image until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Wrap
    
    'Populate the "area outside sphere" button
    btsExterior.AddItem "empty", 0
    btsExterior.AddItem "rays of light", 1
    btsExterior.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
        
    'Create the preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then SpherizeImage GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltOffsetX_Change()
    UpdatePreview
End Sub

Private Sub sltOffsetY_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "angle", sltAngle.Value
        .AddParam "xoffset", sltOffsetX.Value
        .AddParam "yoffset", sltOffsetY.Value
        .AddParam "rays", (btsExterior.ListIndex = 1)
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "quality", sltQuality.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
