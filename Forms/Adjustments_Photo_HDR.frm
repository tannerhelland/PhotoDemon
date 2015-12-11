VERSION 5.00
Begin VB.Form FormHDR 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " HDR"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "quality"
      Min             =   1
      Max             =   100
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   50
   End
End
Attribute VB_Name = "FormHDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Imitation HDR Tool
'Copyright 2014-2015 by Tanner Helland
'Created: 09/July/14
'Last updated: 19/January/15
'Last update: introduce new experimental approach; it's way faster, but it cannot reproduce fine details as sharply.
'              If testers find this new option preferable, I may look at introducing it as a side-by-side option with
'              the old CLAHE method, or possibly retiring CLAHE entirely.
'
'This is a heavily optimized imiation HDR function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform pretty damn well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'HDR normally works by having a photographer take multiple shots of a scene (3-5, typically), each at a unique exposure.
' Software then merges those photos together, selecting pixels from each exposure and blending them to produce an
' evenly exposed photo across a wide luminance range.  This not only produces a neat visual effect, but also allows the
' capturing of detail that would otherwise be impossible from a single exposure.
'
'While a merge-to-HDR function that operates in the traditional manner would be nice to eventually include in PD, the
' trouble of asking a photographer to capture multiple back-to-back photos, each at a different exposure, without
' shaking the camera, is no small feat.  The inclusion of HDR as a built-in mode on many cameras and smartphones has
' also reduced the utility such a technique in a separate piece of software.
'
'So instead, what I've done here is put together a tool that mimics the results of HDR, using a contrast-adaptive local
' histogram equalization function (referred to in the literature as CLAHE).  The details are complicated, but basically
' the function calculates a local histogram around each pixel, using a user-supplied radius (presented in PD as
' "quality").  Each histogram is then partially equalized, while discounting outliers at the top and bottom of the
' spectrum (to reduce the potential for noise upsetting the effect).  The partial equalization results are applied to
' each channel, allowing regions of color to stay consistent, without the distortion inherent to global equalization.
'
'Anyway, assuming the original photograph was exposed reasonably well, this function should produce a very good result.
' Poorly exposed original photographs cannot be saved by this technique, however, especially if a smartphone camera
' or other cheap sensor was used, as the inherent noise will screw up the filter's ability to properly solve the
' partial histogram problem.  C'est la vie.  Applying a median or noise-reduction filter in advance might help to
' improve the output.
'
'I've currently limited the radius to 200, because as much as I've optimized the function, it is still very slow on
' huge images when a largae radius is used.  This could be overcome with a constant-time median function that the
' program dynamically switches to once the radius exceeds ~100ish, but a new function like that is a lot of work, so
' I'm postponing work on it until a later date.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a CLAHE (contrast limited adaptive histogram equalization) filter to the image
'Input: radius of the histogram search (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyCLAHE(ByVal fxQuality As Double, ByVal blendStrength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Generating HDR map for image..."
    
    'The passed radius value will be on the order of [0.0, 100.0].  Convert that to the [0, 200] range.
    Dim mRadius As Long
    mRadius = fxQuality * 2
    
    'Convert blend strength to the [0, 1] scale.  (It is presented to the user on the [0, 100] scale.)
    blendStrength = blendStrength / 100
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
        
    'If this is a preview, we need to adjust the radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If mRadius = 0 Then mRadius = 1
    End If
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'CLAHE filtering RGB data takes a lot of variables
    Dim rValues() As Long, gValues() As Long, bValues() As Long
    ReDim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long
    
    Dim rValuesEq() As Long, gValuesEq() As Long, bValuesEq() As Long
    ReDim rValuesEq(0 To 255) As Long, gValuesEq(0 To 255) As Long, bValuesEq(0 To 255) As Long
    
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim histFactor As Double
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    numOfPixels = 0
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + mRadius - 1
        QuickVal = x * qvDepth
    For y = initY To initY + mRadius
    
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        rValues(r) = rValues(r) + 1
        gValues(g) = gValues(g) + 1
        bValues(b) = bValues(b) + 1
        
        'Increase the pixel tally
        numOfPixels = numOfPixels + 1
        
    Next y
    Next x
                
    'Loop through each pixel in the image, tallying median values as we go
    For x = initX To finalX
            
        QuickVal = x * qvDepth
        
        'Determine the bounds of the current median box in the X direction
        lbX = x - mRadius
        If lbX < 0 Then lbX = 0
        ubX = x + mRadius
        
        If ubX > finalX Then
            obuX = True
            ubX = finalX
        Else
            obuX = False
        End If
                
        'As part of my accumulation algorithm, I swap the inner loop's direction with each iteration.
        ' Set y-related loop variables depending on the direction of the next cycle.
        If atBottom Then
            lbY = 0
            ubY = mRadius
        Else
            lbY = finalY - mRadius
            ubY = finalY
        End If
        
        'Remove trailing values from the median box if they lie outside the processing radius
        If lbX > 0 Then
        
            QuickValInner = (lbX - 1) * qvDepth
        
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                numOfPixels = numOfPixels - 1
            Next j
        
        End If
        
        'Add leading values to the median box if they lie inside the processing radius
        If Not obuX Then
        
            QuickValInner = ubX * qvDepth
            
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) + 1
                gValues(g) = gValues(g) + 1
                bValues(b) = bValues(b) + 1
                numOfPixels = numOfPixels + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
                
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, mRadius)
                g = srcImageData(QuickValInner + 1, mRadius)
                b = srcImageData(QuickValInner, mRadius)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                numOfPixels = numOfPixels - 1
            Next i
        
        Else
        
            QuickY = finalY - mRadius
        
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, QuickY)
                g = srcImageData(QuickValInner + 1, QuickY)
                b = srcImageData(QuickValInner, QuickY)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                numOfPixels = numOfPixels - 1
            Next i
        
        End If
        
        'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
        If atBottom Then
            startY = 0
            stopY = finalY
            yStep = 1
        Else
            startY = finalY
            stopY = 0
            yStep = -1
        End If
            
    'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
    For y = startY To stopY Step yStep
            
        'If we are at the bottom and moving up, we will REMOVE rows from the bottom and ADD them at the top.
        'If we are at the top and moving down, we will REMOVE rows from the top and ADD them at the bottom.
        'As such, there are two copies of this function, one per possible direction.
        If atBottom Then
        
            'Calculate bounds
            lbY = y - mRadius
            If lbY < 0 Then lbY = 0
            
            ubY = y + mRadius
            If ubY > finalY Then
                obuY = True
                ubY = finalY
            Else
                obuY = False
            End If
                                
            'Remove trailing values from the box
            If lbY > 0 Then
            
                QuickY = lbY - 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                    numOfPixels = numOfPixels - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, ubY)
                    g = srcImageData(QuickValInner + 1, ubY)
                    b = srcImageData(QuickValInner, ubY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                    numOfPixels = numOfPixels + 1
                Next i
            
            End If
            
        'The exact same code as above, but in the opposite direction
        Else
        
            lbY = y - mRadius
            If lbY < 0 Then
                oblY = True
                lbY = 0
            Else
                oblY = False
            End If
            
            ubY = y + mRadius
            If ubY > finalY Then ubY = finalY
                                
            If ubY < finalY Then
            
                QuickY = ubY + 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                    numOfPixels = numOfPixels - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, lbY)
                    g = srcImageData(QuickValInner + 1, lbY)
                    b = srcImageData(QuickValInner, lbY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                    numOfPixels = numOfPixels + 1
                Next i
            
            End If
        
        End If
                
        'With the histogram box successfully calculated, we can now perform a partial equalization.
        ' FormEqualize describes this process in more detail, but note that we don't have to equalize
        ' the full histogram here - just the histogram up to the current pixel.
        
        'Update our copies of the original RGB values of the current pixel
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        
        'Histogram equalization applies a unique scale factor based on the number of pixels in the histogram
        ' (Because our sliding-box technique generates different pixel counts along edge regions, we can't
        '  pre-calculate this value.)
        histFactor = 255 / numOfPixels
        
        'Partially equalize each individual channel histogram
        rValuesEq(0) = rValues(0) * histFactor
        
        If r > 0 Then
            For i = 1 To r
                rValuesEq(i) = rValuesEq(i - 1) + (histFactor * rValues(i))
            Next i
        End If
        
        gValuesEq(0) = gValues(0) * histFactor
        
        If g > 0 Then
            For i = 1 To g
                gValuesEq(i) = gValuesEq(i - 1) + (histFactor * gValues(i))
            Next i
        End If
        
        bValuesEq(0) = bValues(0) * histFactor
        
        If b > 0 Then
            For i = 1 To b
                bValuesEq(i) = bValuesEq(i - 1) + (histFactor * bValues(i))
            Next i
        End If
        
        'Clamp values as necessary
        If rValuesEq(r) > 255 Then rValuesEq(r) = 255
        If gValuesEq(g) > 255 Then gValuesEq(g) = 255
        If bValuesEq(b) > 255 Then bValuesEq(b) = 255
        
        'Blend these results with the original pixel at the specified value
        newR = BlendColors(r, rValuesEq(r), blendStrength)
        newG = BlendColors(g, gValuesEq(g), blendStrength)
        newB = BlendColors(b, bValuesEq(b), blendStrength)
        
        'Finally, apply the results to the image.
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        
    Next y
        atBottom = Not atBottom
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
        
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

'New test approach to HDR.  Unsharp masking can produce an HDR-like image, and it can do it a hell of a lot faster
' than the CLAHE-based method we've been using.  I'm going to have some testers experiment with the new method, to see
' if they prefer it.
Public Sub ApplyImitationHDR(ByVal fxQuality As Double, ByVal blendStrength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Generating HDR map for image..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'fxQuality represents an HDR radius.  We actually treat this as a percentage of the current image size, using the
    ' largest dimension.  Max quality is 20% of the image.
    Dim largestDimension As Long
    If (finalX - initX) > (finalY - initY) Then largestDimension = (finalX - initX) Else largestDimension = (finalY - initY)
    
    Dim hdrRadius As Long
    hdrRadius = ((fxQuality / 100) * largestDimension) * 0.2
    
    'Strength is used as an analog for multiple parameters.  Here, we use it to calculate a saturation modifier,
    ' which is applied linearly to the final RGB values, as a way to further pop colors.
    Dim satBoost As Double
    satBoost = 1# + (blendStrength / 100) * 0.3
    
    'Strength is presented to the user on a [1, 100] scale, but we actually boost this to a literal value of [1, 200]
    blendStrength = (blendStrength * 2) / 100
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    'If toPreview Then hdrRadius = hdrRadius * curDIBValues.previewModifier
    If hdrRadius = 0 Then hdrRadius = 1
    
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I use the faster method instead.
    Dim gaussBlurSuccess As Long
    gaussBlurSuccess = 0
    
    Dim progBarCalculation As Long
    progBarCalculation = finalY * 3 + finalX * 3
    gaussBlurSuccess = CreateApproximateGaussianBlurDIB(hdrRadius, workingDIB, srcDIB, 3, toPreview, progBarCalculation + finalX)
    
    'Assuming the blur was created successfully, proceed with the masking portion of the filter.
    If (gaussBlurSuccess <> 0) Then
    
        'Now that we have a gaussian DIB created in workingDIB, we can point arrays toward it and the source DIB
        Dim dstImageData() As Byte
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        prepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = curDIBValues.BytesPerPixel
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = findBestProgBarValue()
            
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10.
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = blendStrength + 1
        invScaleFactor = 1 - scaleFactor
    
        Dim blendVal As Double
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long, a As Long
        Dim r2 As Long, g2 As Long, b2 As Long, a2 As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        Dim h As Double, s As Double, l As Double
        Dim tLumDelta As Long
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Retrieve the original image's pixels
            r = dstImageData(QuickVal + 2, y)
            g = dstImageData(QuickVal + 1, y)
            b = dstImageData(QuickVal, y)
            
            'Now, retrieve the gaussian pixels
            r2 = srcImageData(QuickVal + 2, y)
            g2 = srcImageData(QuickVal + 1, y)
            b2 = srcImageData(QuickVal, y)
            
            tLumDelta = Abs(getLuminance(r, g, b) - getLuminance(r2, g2, b2))
            
            newR = (scaleFactor * r) + (invScaleFactor * r2)
            If newR > 255 Then newR = 255
            If newR < 0 Then newR = 0
            
            newG = (scaleFactor * g) + (invScaleFactor * g2)
            If newG > 255 Then newG = 255
            If newG < 0 Then newG = 0
            
            newB = (scaleFactor * b) + (invScaleFactor * b2)
            If newB > 255 Then newB = 255
            If newB < 0 Then newB = 0
            
            blendVal = tLumDelta / 255
            
            newR = BlendColors(newR, r, blendVal)
            newG = BlendColors(newG, g, blendVal)
            newB = BlendColors(newB, b, blendVal)
            
            'Finally, apply a saturation boost proportional to the final calculated strength
            tRGBToHSL newR, newG, newB, h, s, l
            s = s * satBoost
            If s > 1 Then s = 1
            tHSLToRGB h, s, l, newR, newG, newB
            
            dstImageData(QuickVal + 2, y) = newR
            dstImageData(QuickVal + 1, y) = newG
            dstImageData(QuickVal, y) = newB
            
            If qvDepth = 4 Then
                a2 = srcImageData(QuickVal + 3, y)
                a = dstImageData(QuickVal + 3, y)
                newA = (scaleFactor * a) + (invScaleFactor * a2)
                If newA > 255 Then newA = 255
                If newA < 0 Then newA = 0
                dstImageData(QuickVal + 3, y) = BlendColors(newA, a, blendVal)
            End If
                                    
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal progBarCalculation + x
                End If
            End If
        Next x
        
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        srcDIB.eraseDIB
        Set srcDIB = Nothing
        
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        Erase dstImageData
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "HDR", , buildParams(sltRadius.Value, sltStrength.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius = 5
    sltStrength = 20
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyImitationHDR sltRadius.Value, sltStrength.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub
