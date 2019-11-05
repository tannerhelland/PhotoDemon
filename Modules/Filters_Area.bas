Attribute VB_Name = "Filters_Area"
'***************************************************************************
'Filter (Area) Interface
'Copyright 2001-2019 by Tanner Helland
'Created: 12/June/01
'Last updated: 31/July/17
'Last update: migrate the convolution filter functions to XML param strings
'
'Holder module for generalized area filters, including most of the project's convolution filters.
'
'The most interesting function is ConvolveDIB, which applies any arbitrary 5x5 convolution filter to any arbitrary DIB.  This function
' is used internally for nearly all edge-detection functions, and other generic convolution effects.  (Note that some convolution
' filters, like Gaussian Blur, have their own specialized, optimized implementations.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Cache for IIR gaussian blur terms; these are only calculated as necessary, to improve performance
Private m_GaussTerms() As Long, m_GaussTermCount As Long
Private m_lastSigma As Double, m_lastNumSteps As Long
Private m_nu As Double, m_invNu As Double, m_numTerms As Long, m_preScale As Double

'The omnipotent ApplyConvolutionFilter routine, which applies the supplied convolution filter to the current image.
' Note that as of July '17, ApplyConvolutionFilter uses an XML param string for supplying convolution details.
' The relevant ParamString entries are as follows:
'    <name>: String
'    <invert>: Boolean
'    <weight>: Double
'    <bias>: Long
'    <matrix>: a pipe-delimited string containing 25 floating-point values (e.g. 0.0|1.0|0.0|-50.0....).  These values
'              represent the entries in a 5x5 convolution matrix, in left-to-right, top-to-bottom order.
Public Sub ApplyConvolutionFilter_XML(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Prepare a param parser
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
        
    'Note that the only purpose of the FilterType string is to display this message
    If (Not toPreview) Then Message "Applying %1 filter...", cParams.GetString("name")
    
    'Create a local array and point it at the pixel data of the current image.  Note that the current layer is referred to as the
    ' DESTINATION image for the convolution; we will make a separate temp copy of the image to use as the SOURCE.
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'Use the central ConvolveDIB function to apply the convolution
    ConvolveDIB_XML effectParams, srcDIB, workingDIB, toPreview
    
    'Free our temporary DIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

'Apply any convolution filter to a pdDIB object.  This is primarily used by the ApplyConvolutionFilter() function, above,
' but it can also be used to apply multiple convolutions in succession, or to create standalone convolved images that can
' then be used for further image analysis.
Public Function ConvolveDIB_XML(ByVal effectParams As String, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Parameters are passed via XML; this parser will retrieve individual values for us
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    'Create a local array and point it at the destination pixel data
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from corrupting subsequent calculations.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, x2 As Long, y2 As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim checkXMin As Long, checkXMax As Long, checkYMin As Long, checkYMax As Long
    checkXMin = initX
    checkXMax = finalX
    checkYMin = initY
    checkYMax = finalY
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
        
    'We can now parse out the relevant filter values from the param string
    Dim invertResult As Boolean
    invertResult = cParams.GetBool("invert", False)
    
    Dim filterWeightA As Double, filterBiasA As Double
    filterWeightA = cParams.GetDouble("weight", 1#)
    filterBiasA = cParams.GetDouble("bias", 0#)
    
    'The actual filter values are stored inside a single pipe-delimited string
    Dim filterMatrix() As String
    filterMatrix = Split(cParams.GetString("matrix"), "|", , vbBinaryCompare)
    
    Dim iFM(-2 To 2, -2 To 2) As Double
    For x = -2 To 2
    For y = -2 To 2
        iFM(x, y) = TextSupport.CDblCustom(filterMatrix((x + 2) + (y + 2) * 5))
    Next y
    Next x
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Double, g As Double, b As Double
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately when attempting to calculate the value for pixels
    ' outside the image perimeter
    Dim filterWeightTemp As Double
    
    'Temporary calculation variables
    Dim calcX As Long, calcY As Long, convValue As Double
    Dim xOffset As Long
        
    'Apply the filter
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
        
        'Reset our values upon beginning analysis on a new pixel
        r = 0#
        g = 0#
        b = 0#
        filterWeightTemp = filterWeightA
        
        'Run a sub-loop around the current pixel
        For x2 = x - 2 To x + 2
            xOffset = x2 * qvDepth
        For y2 = y - 2 To y + 2
        
            calcX = x2 - x
            calcY = y2 - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming,
            ' but because VB does not provide a "continue next" type mechanism, GoTo's are all we've got.)
            convValue = iFM(calcX, calcY)
            If (convValue <> 0#) Then
            
                'If this pixel lies outside the image perimeter, ignore it and adjust the filter's weight value accordingly
                If (x2 < checkXMin) Or (y2 < checkYMin) Or (x2 > checkXMax) Or (y2 > checkYMax) Then
                    filterWeightTemp = filterWeightTemp - iFM(calcX, calcY)
                
                Else
                
                    'Adjust red, green, and blue according to the values in the filter matrix (FM)
                    b = b + (srcImageData(xOffset, y2) * convValue)
                    g = g + (srcImageData(xOffset + 1, y2) * convValue)
                    r = r + (srcImageData(xOffset + 2, y2) * convValue)
                    
                End If
                
            End If
    
        Next y2
        Next x2
        
        'If a weight has been set, apply it now
        If (filterWeightTemp <> 1#) Then
        
            'Catch potential divide-by-zero errors
            If (filterWeightTemp <> 0#) Then
                filterWeightTemp = 1# / filterWeightTemp
                r = r * filterWeightTemp
                g = g * filterWeightTemp
                b = b * filterWeightTemp
            Else
                r = 0#
                g = 0#
                b = 0#
            End If
            
        End If
        
        'If a bias has been specified, apply it now
        r = r + filterBiasA
        g = g + filterBiasA
        b = b + filterBiasA
        
        'Make sure all values are between 0 and 255
        If (r < 0#) Then
            r = 0#
        ElseIf (r > 255#) Then
            r = 255#
        End If
        
        If (g < 0#) Then
            g = 0#
        ElseIf (g > 255#) Then
            g = 255#
        End If
        
        If (b < 0#) Then
            b = 0#
        ElseIf (b > 255#) Then
            b = 255#
        End If
        
        'If inversion is specified, apply it now
        If invertResult Then
            r = 255# - r
            g = 255# - g
            b = 255# - b
        End If
        
        'Copy the calculated value into the destination array
        dstImageData(quickVal, y) = Int(b)
        dstImageData(quickVal + 1, y) = Int(g)
        dstImageData(quickVal + 2, y) = Int(r)
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate all intermediary array
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    'Return success/failure
    If g_cancelCurrentAction Then ConvolveDIB_XML = 0 Else ConvolveDIB_XML = 1

End Function

'Apply a grid blur to an image; basically, blur every vertical line, then every horizontal line, then average the results
Public Sub FilterGridBlur()

    Message "Generating grids..."

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim iWidth As Long, iHeight As Long
    iWidth = curDIBValues.Width
    iHeight = curDIBValues.Height
            
    Dim numOfPixels As Long
    numOfPixels = iWidth + iHeight
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim rax() As Long, gax() As Long, bax() As Long
    Dim ray() As Long, gay() As Long, bay() As Long
    ReDim rax(0 To iWidth) As Long, gax(0 To iWidth) As Long, bax(0 To iWidth) As Long
    ReDim ray(0 To iHeight) As Long, gay(0 To iHeight), bay(0 To iHeight)
    
    'Generate the averages for vertical lines
    For x = initX To finalX
        r = 0
        g = 0
        b = 0
        quickVal = x * qvDepth
        For y = initY To finalY
            r = r + imageData(quickVal + 2, y)
            g = g + imageData(quickVal + 1, y)
            b = b + imageData(quickVal, y)
        Next y
        rax(x) = r
        gax(x) = g
        bax(x) = b
    Next x
    
    'Generate the averages for horizontal lines
    For y = initY To finalY
        r = 0
        g = 0
        b = 0
        For x = initX To finalX
            quickVal = x * qvDepth
            r = r + imageData(quickVal + 2, y)
            g = g + imageData(quickVal + 1, y)
            b = b + imageData(quickVal, y)
        Next x
        ray(y) = r
        gay(y) = g
        bay(y) = b
    Next y
    
    Message "Applying grid blur..."
        
    'Apply the filter
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
        
        'Average the horizontal and vertical values for each color component
        r = (rax(x) + ray(y)) \ numOfPixels
        g = (gax(x) + gay(y)) \ numOfPixels
        b = (bax(x) + bay(y)) \ numOfPixels
        
        'The colors shouldn't exceed 255, but it doesn't hurt to double-check
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        'Assign the new RGB values back into the array
        imageData(quickVal + 2, y) = r
        imageData(quickVal + 1, y) = g
        imageData(quickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData

End Sub

'Given a quality setting (direct from the user), populate a table of supersampling offsets.  For maximum quality, PD uses a modified
' rotated grid approach (see http://en.wikipedia.org/wiki/Spatial_anti-aliasing), with hard-coded offset tables based on the passed
' quality param.  At present, a maximum of 12 supersamples (plus the original sample) are used at maximum quality.  Beyond this level,
' performance takes a huge hit, but output results are not significantly improved.
Public Sub GetSupersamplingTable(ByVal userQuality As Long, ByRef numAASamples As Long, ByRef ssOffsetsX() As Single, ByRef ssOffsetsY() As Single)
    
    'Old PD versions used a Boolean value for quality.  As such, if the user enabled interpolation, and saved it as part of a preset,
    ' this function may get passed a "-1" for userQuality.  In that case, activate an identical method in the new supersampler.
    If (userQuality < 1) Then userQuality = 2
    
    'Quality is typically presented to the user on a 1-5 scale.  1 = lowest quality/highest speed, 5 = highest quality/lowest speed.
    Select Case userQuality
    
        'Quality settings of 1 and 2 both suspend supersampling.  The only difference is that the calling function, per PD convention,
        ' will disable antialising.
        Case 1, 2
        
            numAASamples = 1
            ReDim ssOffsetsX(0) As Single
            ReDim ssOffsetsY(0) As Single
            ssOffsetsX(0) = 0
            ssOffsetsY(0) = 0
        
        'Cases 3, 4, 5: use rotated grid supersampling, at the recommended rotation of arctan(1/2), with 4 additional sample points
        ' per quality level.
        Case Else
        
            'Four additional samples are provided at each quality level
            numAASamples = (userQuality - 2) * 4 + 1
            ReDim ssOffsetsX(0 To numAASamples - 1) As Single
            ReDim ssOffsetsY(0 To numAASamples - 1) As Single
            
            'The first sample point is always the origin pixel.  This is used as the basis of adaptive supersampling,
            ' and should not be changed.
            ssOffsetsX(0) = 0
            ssOffsetsY(0) = 0
            
            'The other 4 sample points are calculated as follows:
            ' - Rotate (0, 0.5) around (0, 0) by arctan(1/2) radians
            ' - Repeat the above step, but increasing each rotation by 90.
            ssOffsetsX(1) = 0.447077
            ssOffsetsY(1) = 0.22388
            
            ssOffsetsX(2) = -0.447077
            ssOffsetsY(2) = -0.22388
            
            ssOffsetsX(3) = -0.22388
            ssOffsetsY(3) = 0.447077
            
            ssOffsetsX(4) = 0.22388
            ssOffsetsY(4) = -0.447077
            
            'For quality levels 4 and 5, we add a second set of sampling points, closer to the origin, and offset from the originals
            ' by 45 degrees
            If (userQuality > 3) Then
            
                ssOffsetsX(5) = 0.0789123
                ssOffsetsY(5) = 0.237219
                
                ssOffsetsX(6) = -0.237219
                ssOffsetsY(6) = 0.0789123
                
                ssOffsetsX(7) = -0.0789123
                ssOffsetsY(7) = -0.237219
                
                ssOffsetsX(8) = 0.237219
                ssOffsetsY(8) = -0.0789123
            
                'For the final quality level, add a set of 4 more points, calculated by rotating (0, 0.67) around the
                ' origin in 45 degree increments.  The benefits of this are minimal for all but the most extreme
                ' zoom-out situations.
                If (userQuality > 4) Then
                
                    ssOffsetsX(9) = 0.473762
                    ssOffsetsY(9) = 0.473762
                    
                    ssOffsetsX(10) = -0.473762
                    ssOffsetsY(10) = 0.473762
                    
                    ssOffsetsX(11) = -0.473762
                    ssOffsetsY(11) = -0.473762
                    
                    ssOffsetsX(12) = 0.473762
                    ssOffsetsY(12) = -0.473762
                    
                End If
            
            End If
    
    End Select

End Sub

'Half-sample symmetric boundary extension
' - Original C version is copyright (c) 2012-2013, Pascal Getreuer <getreuer@cmla.ens-cachan.fr>
' - Used here under its original simplified BSD license <http://www.opensource.org/licenses/bsd-license.html>
' - Translated into VB6 by Tanner Helland in 2019
Private Function inf_extension(ByVal numSteps As Long, ByVal n As Long) As Long
    
    Do
        If (n < 0) Then
            n = -1 - n  '/* Reflect over n = -1/2.    */
        ElseIf (n >= numSteps) Then
            n = 2 * numSteps - 1 - n    '/* Reflect over n = N - 1/2. */
        Else
            Exit Do
        End If
    Loop While True
    
    inf_extension = n
    
End Function

'Handling of the left boundary for Alvarez-Mazorra
' - Original C version is copyright (c) 2012-2013, Pascal Getreuer <getreuer@cmla.ens-cachan.fr>
' - Used here under its original simplified BSD license <http://www.opensource.org/licenses/bsd-license.html>
' - Translated into VB6 by Tanner Helland in 2019
Private Function am_left_boundary(ByRef srcFloat() As Single, ByVal initOffset As Long, ByVal numSteps As Long, ByVal srcStride As Long, ByVal nu As Double, ByVal numTerms As Long) As Double
    
    Dim h As Double, accum As Double
    h = 1#
    accum = srcFloat(initOffset)
    
    Dim m As Long
    
    'Pre-calculate terms table only as necessary
    If (numTerms <> m_GaussTermCount) Then
        ReDim m_GaussTerms(0 To numTerms - 1) As Long
        m_GaussTermCount = numTerms
        For m = 1 To numTerms - 1
            m_GaussTerms(m) = inf_extension(numSteps, -m)
        Next m
    End If
    
    For m = 1 To numTerms - 1
        h = h * nu
        accum = accum + (h * srcFloat(initOffset + srcStride * m_GaussTerms(m)))
    Next m
    
    am_left_boundary = accum
    
End Function

'Implements the fast approximate Gaussian convolution algorithm of Alvarez and Mazorra,
' where the Gaussian is approximated by the heat equation and each timestep is performed
' with an efficient recursive computation.  Using more steps yields a more accurate approximation
' of the Gaussian. Reasonable values for the parameters are `numSteps` = 4, `tol` = 1e-3.
' - Original C version is copyright (c) 2012-2013, Pascal Getreuer <getreuer@cmla.ens-cachan.fr>
' - Used here under its original simplified BSD license <http://www.opensource.org/licenses/bsd-license.html>
' - Translated into VB6 by Tanner Helland in 2019
Private Sub am_gaussian_conv(ByRef srcFloat() As Single, ByVal initOffset As Long, ByVal numElements As Long, ByVal srcStride As Long, ByVal sigma As Double, ByVal numSteps As Long, ByVal tol As Double, ByVal useAdjustedQ As Boolean)
    
    'To improve performance, we only calculate initial terms when sigma or numSteps changes.
    ' (Initial terms depend only on these and tolerance, but in PD, we do not vary tolerance
    ' so we never need to check it for changes.)
    If (sigma <> m_lastSigma) Or (m_lastNumSteps <> numSteps) Then
        
        m_lastSigma = sigma
        m_lastNumSteps = numSteps
        
        '/* Use a regression on q for improved accuracy. */
        Dim q As Double
        If useAdjustedQ Then
            q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
        
        '/* Use q = sigma as in the original A-M method. */
        Else
            q = sigma
        End If
    
        '/* Precompute the filter coefficient nu. */
        Dim lambda As Double, dnu As Double
        lambda = (q * q) / (2# * numSteps)
        dnu = (1# + 2# * lambda - Sqr(1# + 4# * lambda)) / (2# * lambda)
        m_nu = dnu
        
        'Exists only as an optimization, to skip division in the inner loop
        m_invNu = 1# / (1# - m_nu)
    
        '/* For handling the left boundary, determine the number of terms needed to
        '   approximate the sum with accuracy tol. */
        m_numTerms = Int(Log((1# - dnu) * tol) / Log(dnu) + 1)
    
        '/* Precompute the constant scale factor. */
        m_preScale = (dnu / lambda) ^ numSteps
        
    End If
        
    '/* Copy src to dest and multiply by the constant scale factor. */
    Dim stride_N As Long
    stride_N = srcStride * numElements
    
    Dim i As Long
    For i = 0 To (stride_N - 1) Step srcStride
        srcFloat(initOffset + i) = srcFloat(initOffset + i) * m_preScale
    Next i
    
    Dim strideOffset As Long
    strideOffset = initOffset - srcStride
    
    '/* Perform K passes of filtering. */
    Dim pass As Long
    For pass = 0 To numSteps - 1
    
        '/* Initialize the recursive filter on the left boundary. */
        srcFloat(initOffset) = am_left_boundary(srcFloat, initOffset, numSteps, srcStride, m_nu, m_numTerms)
        
        '/* This loop applies the causal filter, implementing the pseudocode
        '
        '   For n = 1, ..., N - 1
        '       dest(n) = dest(n) + nu dest(n - 1)
        '
        '   Variable i = stride * n is the offset to the nth sample.  */
        For i = srcStride To (stride_N - 1) Step srcStride
            srcFloat(initOffset + i) = srcFloat(initOffset + i) + m_nu * srcFloat(strideOffset + i)
        Next i
        
        '/* Handle the right boundary. */
        i = i - srcStride
        srcFloat(initOffset + i) = srcFloat(initOffset + i) * m_invNu
        
        '/* Similarly, this loop applies the anticausal filter as
        '
        '   For n = N - 1, ..., 1
        '       dest(n - 1) = dest(n - 1) + nu dest(n) */
        Do While (i > 0)
            srcFloat(strideOffset + i) = srcFloat(strideOffset + i) + m_nu * srcFloat(initOffset + i)
            i = i - srcStride
        Loop
    
    Next pass
    
End Sub

'Gaussian blur filter, using an approximation originally by Alvarez and Mazorra, as implemented by
' Pascal Getreuer.
' - Original C version is copyright (c) 2012-2013, Pascal Getreuer <getreuer@cmla.ens-cachan.fr>
' - Used here under its original simplified BSD license <http://www.opensource.org/licenses/bsd-license.html>
' - Translated into VB6 by Tanner Helland in 2019
Public Function GaussianBlur_IIRImplementation(ByRef srcDIB As pdDIB, ByVal radius As Double, ByVal numSteps As Long, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'First comes a mathematical fudge.  This particular gaussian approximation tends to produce a
    ' slightly "genter" blur than an identical radius in Photoshop.  To try and bring the two methods
    ' into (rough) alignment, I slightly increase the radius used by this method.  There's no obvious
    ' mathematical explanation for this, alas - I just determined this value experimentally and
    ' plug it in to better unify the results.  (Note that PD's 3x iterative box blur produces nearly
    ' identical results to Photoshop, so it's likely that Adobe uses some variation on that technique
    ' as well, at least in older versions of their software.)
    radius = radius * 1.075
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim pxImgWidth As Long, pxImgHeight As Long
    pxImgWidth = srcDIB.GetDIBWidth
    pxImgHeight = srcDIB.GetDIBHeight
    
    'These values will help us access locations in the array more quickly.
    ' (pxSizeBytes is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim pxSizeBytes As Long
    pxSizeBytes = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (modifyProgBarMax = -1) Then modifyProgBarMax = 4 * pxSizeBytes
    If (Not suppressMessages) Then SetProgBarMax modifyProgBarMax
    
    'Finally, a bunch of variables used in color calculation
    Dim origValue As Long
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    'Calculate sigma from the radius, using a similar formula to ImageJ (per this link:
    ' http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur)
    ' The idea here is to convert the radius to a sigma of sufficient magnitude where the outer edges
    ' of the gaussian no longer represent meaningful values on a [0, 255] scale.
    Dim sigma As Double
    Const LOG_255_BASE_10 As Double = 2.40654018043395
    sigma = (radius + 1#) / Sqr(2# * LOG_255_BASE_10)
    
    'Make sure sigma and steps are not so small as to produce errors or invisible results
    If (sigma <= 0#) Then sigma = 0.001
    If (numSteps < 1) Then
        numSteps = 1
    ElseIf (numSteps > 5) Then
        numSteps = 5
    End If
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate
    ' lambda calculation is proposed.  This adjustment doesn't affect running time at all, and it could
    ' potentially reduce errors relative to a pure Gaussian - but only at small radii.
    Dim useModifiedQ As Boolean
    useModifiedQ = True
    
    'To ensure ideal edge-handling, we also need to calculate how many terms are required to
    ' approximate to a given tolerance.  (The tolerance value used here, 1e-3, comes from
    ' this URL: http://www.ipol.im/pub/art/2013/87/?utm_source=doi)
    Const INF_SUM_TOLERANCE As Double = 0.001
    
    'Because this technique requires conversion to/from [0, 1] floats for *all* source channels,
    ' it can potentially consume a ton of memory (e.g. 16x an image's original size in bytes -
    ' 4x channels * 4x bytes per float).  To mitigate this, we process each channel individually,
    ' sharing a single buffer across channels.  This is slightly slower but much lighter on memory.
    Dim numPixels As Long
    numPixels = pxImgWidth * pxImgHeight
    
    Dim tmpFloat() As Single
    ReDim tmpFloat(0 To numPixels - 1) As Single
    
    'If requested, progress events are raised as discrete steps
    Dim progressTracker As Long
    progressTracker = 0
    
    Dim curChannel As Long
    For curChannel = 0 To pxSizeBytes - 1
        
        'Copy the contents of the current image into the float arrays and apply pre-scaling
        Dim xOffset As Long
        For y = 0 To finalY
            srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
            xOffset = y * pxImgWidth
        For x = 0 To finalX
            tmpFloat(x + xOffset) = imageData(x * pxSizeBytes + curChannel)
        Next x
        Next y
        
        If (Not suppressMessages) Then
            If Interface.UserPressedESC() Then Exit For
            progressTracker = progressTracker + 1
            SetProgBarVal progressTracker
        End If
        
        'All subsequent handling is provided by a separate, dedicated function
        For y = 0 To finalY
            am_gaussian_conv tmpFloat, y * pxImgWidth, pxImgWidth, 1, sigma, numSteps, INF_SUM_TOLERANCE, useModifiedQ
        Next y
        
        If (Not suppressMessages) Then
            If Interface.UserPressedESC() Then Exit For
            progressTracker = progressTracker + 1
            SetProgBarVal progressTracker
        End If
        
        'Next, filter all columns
        For x = 0 To finalX
            am_gaussian_conv tmpFloat, x, pxImgHeight, pxImgWidth, sigma, numSteps, INF_SUM_TOLERANCE, useModifiedQ
        Next x
        
        If (Not suppressMessages) Then
            If Interface.UserPressedESC() Then Exit For
            progressTracker = progressTracker + 1
            SetProgBarVal progressTracker
        End If
        
        'Apply final post-scaling
        For y = 0 To finalY
            srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
            xOffset = y * pxImgWidth
        For x = 0 To finalX
            
            'Round the finished result, perform failsafe clipping, then assign
            origValue = Int(tmpFloat(xOffset + x) + 0.5)
            If (origValue > 255) Then origValue = 255
            imageData(x * pxSizeBytes + curChannel) = origValue
            
        Next x
        Next y
        
        If (Not suppressMessages) Then
            If Interface.UserPressedESC() Then Exit For
            progressTracker = progressTracker + 1
            SetProgBarVal progressTracker
        End If
        
    Next curChannel
    
    If (Not suppressMessages) Then ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax
    
    'Regardless of success/failure, safely deallocate our fake pixel wrapper
    srcDIB.UnwrapArrayFromDIB imageData
    
    If g_cancelCurrentAction Then GaussianBlur_IIRImplementation = 0 Else GaussianBlur_IIRImplementation = 1

End Function

'Horizontal blur filter, using an IIR (Infininte Impulse Response) approach.
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
Public Function HorizontalBlur_IIR(ByRef srcDIB As pdDIB, ByVal radius As Double, ByVal numSteps As Long, Optional ByVal blurSymmetric As Boolean = True, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim iWidth As Long, iHeight As Long
    iWidth = srcDIB.GetDIBWidth
    iHeight = srcDIB.GetDIBHeight
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickX As Long, quickX2 As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'Determine if alpha handling is necessary for this image
    Dim imgHasAlpha As Boolean
    imgHasAlpha = (qvDepth = 4)
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If modifyProgBarMax = -1 Then modifyProgBarMax = srcDIB.GetDIBWidth
    If Not suppressMessages Then SetProgBarMax modifyProgBarMax
    
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Calculate sigma from the radius, using the same formula we do for PD's pure gaussian blur
    Dim sigma As Double
    radius = radius * 1.075
    Const LOG_255_BASE_10 As Double = 2.40654018043395
    sigma = (radius + 1#) / Sqr(2# * LOG_255_BASE_10)
    
    'Make sure sigma and steps are not so small as to produce errors or invisible results
    If (sigma <= 0#) Then sigma = 0.001
    If (numSteps < 1) Then
        numSteps = 1
    ElseIf (numSteps > 5) Then
        numSteps = 5
    End If
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate lambda calculation
    ' is proposed.  This adjustment doesn't affect running time at all, and should reduce errors relative to a pure Gaussian.
    ' The behavior could be toggled by the caller, but for now, I've hard-coded use of the modified formula.
    Dim useModifiedQ As Boolean, q As Double
    useModifiedQ = True
    
    If useModifiedQ Then
        q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
    Else
        q = sigma
    End If
    
    'Prep some IIR-specific values next
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    Dim step As Long
    lambda = (q * q) / (2# * numSteps)
    dnu = (1# + 2# * lambda - Sqr(1# + 4 * lambda)) / (2# * lambda)
    nu = dnu
    boundaryScale = (1# / (1# - dnu))
    If blurSymmetric Then
        postScale = Sqr((dnu / lambda) ^ (2# * numSteps)) * 255#
    Else
        postScale = Sqr((dnu / lambda) ^ numSteps) * 255#
    End If
    
    'Intermediate float arrays are required, so this technique consumes a *lot* of memory.
    Dim rFloat() As Single, gFloat() As Single, bFloat() As Single, aFloat() As Single
    ReDim rFloat(0 To finalX) As Single
    ReDim gFloat(0 To finalX) As Single
    ReDim bFloat(0 To finalX) As Single
    If imgHasAlpha Then ReDim aFloat(0 To finalX) As Single
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    '/* Filter horizontally along each row */
    For y = 0 To finalY
        
        'Wrap an array around the current scanline of pixels
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
        
        'Populate the float arrays
        For x = 0 To finalX
        
            quickX = x * qvDepth
                
            b = imageData(quickX)
            g = imageData(quickX + 1)
            r = imageData(quickX + 2)
            
            rFloat(x) = r * ONE_DIV_255
            gFloat(x) = g * ONE_DIV_255
            bFloat(x) = b * ONE_DIV_255
            
            If imgHasAlpha Then
                a = imageData(quickX + 3)
                aFloat(x) = a * ONE_DIV_255
            End If

        Next x
        
        'Apply the blur
        For step = 0 To numSteps - 1
            
            'Set initial values
            rFloat(0) = rFloat(0) * boundaryScale
            gFloat(0) = gFloat(0) * boundaryScale
            bFloat(0) = bFloat(0) * boundaryScale
            
            'Filter right
            For x = 1 To finalX
                quickX2 = (x - 1)
                rFloat(x) = rFloat(x) + nu * rFloat(quickX2)
                gFloat(x) = gFloat(x) + nu * gFloat(quickX2)
                bFloat(x) = bFloat(x) + nu * bFloat(quickX2)
            Next x
            
            'Filter left only if symmetric
            If blurSymmetric Then
                            
                'Fix closing row
                rFloat(finalX) = rFloat(finalX) * boundaryScale
                gFloat(finalX) = gFloat(finalX) * boundaryScale
                bFloat(finalX) = bFloat(finalX) * boundaryScale
                
                For x = finalX To 1 Step -1
                    quickX = (x - 1)
                    rFloat(quickX) = rFloat(quickX) + nu * rFloat(x)
                    gFloat(quickX) = gFloat(quickX) + nu * gFloat(x)
                    bFloat(quickX) = bFloat(quickX) + nu * bFloat(x)
                Next x
                
            End If
            
            'Apply alpha separately
            If imgHasAlpha Then
                
                aFloat(0) = aFloat(0) * boundaryScale
                
                For x = 1 To finalX
                    aFloat(x) = aFloat(x) + nu * aFloat(x - 1)
                Next x
                
                If blurSymmetric Then
                    aFloat(finalX) = aFloat(finalX) * boundaryScale
                    For x = finalX To 1 Step -1
                        quickX = (x - 1)
                        aFloat(quickX) = aFloat(quickX) + nu * aFloat(x)
                    Next x
                End If
            
            End If
            
        Next step
        
        'Apply final post-scaling
        For x = 0 To finalX
            
            r = rFloat(x) * postScale
            g = gFloat(x) * postScale
            b = bFloat(x) * postScale
            
            'Perform failsafe clipping
            If (r > 255) Then r = 255
            If (g > 255) Then g = 255
            If (b > 255) Then b = 255
            
            quickX = x * qvDepth
            imageData(quickX) = b
            imageData(quickX + 1) = g
            imageData(quickX + 2) = r
            
            'Handle alpha separately
            If imgHasAlpha Then
                a = aFloat(x) * postScale
                If (a > 255) Then a = 255
                imageData(quickX + 3) = a
            End If
        
        Next x
        
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
        
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    If g_cancelCurrentAction Then HorizontalBlur_IIR = 0 Else HorizontalBlur_IIR = 1

End Function
