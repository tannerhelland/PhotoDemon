Attribute VB_Name = "Filters_Area"
'***************************************************************************
'Filter (Area) Interface
'Copyright 2001-2017 by Tanner Helland
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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
    Dim dstSA As SAFEARRAY2D
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
    Dim dstSA As SAFEARRAY2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixel values from corrupting subsequent calculations.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
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
    Dim tmpSA As SAFEARRAY2D
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

'Gaussian blur filter, using an IIR (Infininte Impulse Response) approach
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
Public Function GaussianBlur_IIRImplementation(ByRef srcDIB As pdDIB, ByVal radius As Double, ByVal numSteps As Long, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim iWidth As Long, iHeight As Long
    iWidth = srcDIB.GetDIBWidth
    iHeight = srcDIB.GetDIBHeight
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickX As Long, quickX2 As Long, QuickY As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'Determine if alpha handling is necessary for this image
    Dim hasAlpha As Boolean
    hasAlpha = CBool(qvDepth = 4)
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If modifyProgBarMax = -1 Then modifyProgBarMax = srcDIB.GetDIBWidth + srcDIB.GetDIBHeight
    If Not suppressMessages Then SetProgBarMax modifyProgBarMax
    
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Prep some IIR-specific values next
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    Dim step As Long
    
    'Calculate sigma from the radius, using the same formula we do for PD's pure gaussian blur
    Dim sigma As Double
    sigma = Sqr(-(radius * radius) / (2# * Log(1# / 255#)))
    
    'Another possible sigma formula, per this link (http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur):
    'sigma = (radius + 1) / Sqr(2 * (Log(255) / Log(10)))
    
    'Make sure sigma and steps are valid
    If (sigma <= 0) Then sigma = 0.01
    If (numSteps <= 0) Then numSteps = 1
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate lambda calculation
    ' is proposed.  This adjustment doesn't affect running time at all, and should reduce errors relative to a pure Gaussian.
    ' The behavior could be toggled by the caller, but for now, I've hard-coded use of the modified formula.
    Dim useModifiedQ As Boolean, q As Single
    useModifiedQ = True
    
    If useModifiedQ Then
        q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
    Else
        q = sigma
    End If
    
    'Calculate IIR values
    lambda = (q * q) / (2# * numSteps)
    dnu = (1# + 2# * lambda - Sqr(1# + 4# * lambda)) / (2# * lambda)
    nu = dnu
    boundaryScale = (1# / (1# - dnu))
    postScale = ((dnu / lambda) ^ (2# * numSteps)) * 255#
    
    'Intermediate float arrays are required, so this technique consumes a *lot* of memory.
    Dim rFloat() As Single, gFloat() As Single, bFloat() As Single, aFloat() As Single
    ReDim rFloat(initX To finalX, initY To finalY) As Single
    ReDim gFloat(initX To finalX, initY To finalY) As Single
    ReDim bFloat(initX To finalX, initY To finalY) As Single
    
    If hasAlpha Then ReDim aFloat(initX To finalX, initY To finalY) As Single
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Copy the contents of the current image into the float arrays
    For x = initX To finalX
        quickX = x * qvDepth
    For y = initY To finalY
        
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        rFloat(x, y) = r * ONE_DIV_255
        gFloat(x, y) = g * ONE_DIV_255
        bFloat(x, y) = b * ONE_DIV_255
        
        If hasAlpha Then
            a = imageData(quickX + 3, y)
            aFloat(x, y) = a * ONE_DIV_255
        End If

    Next y
    Next x
    
    '/* Filter horizontally along each row */
    For y = initY To finalY
    
        For step = 0 To numSteps - 1
            
            'Set initial values
            rFloat(initX, y) = rFloat(initX, y) * boundaryScale
            gFloat(initX, y) = gFloat(initX, y) * boundaryScale
            bFloat(initX, y) = bFloat(initX, y) * boundaryScale
            
            'Filter right
            For x = initX + 1 To finalX
                quickX2 = (x - 1)
                rFloat(x, y) = rFloat(x, y) + nu * rFloat(quickX2, y)
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(quickX2, y)
                bFloat(x, y) = bFloat(x, y) + nu * bFloat(quickX2, y)
            Next x
            
            'Fix closing row
            rFloat(finalX, y) = rFloat(finalX, y) * boundaryScale
            gFloat(finalX, y) = gFloat(finalX, y) * boundaryScale
            bFloat(finalX, y) = bFloat(finalX, y) * boundaryScale
            
            'Filter left
            For x = finalX To 1 Step -1
                quickX = (x - 1)
                rFloat(quickX, y) = rFloat(quickX, y) + nu * rFloat(x, y)
                gFloat(quickX, y) = gFloat(quickX, y) + nu * gFloat(x, y)
                bFloat(quickX, y) = bFloat(quickX, y) + nu * bFloat(x, y)
            Next x
            
            'Apply alpha separately
            If hasAlpha Then
                
                aFloat(initX, y) = aFloat(initX, y) * boundaryScale
                
                For x = initX + 1 To finalX
                    aFloat(x, y) = aFloat(x, y) + nu * aFloat(x - 1, y)
                Next x
                
                aFloat(finalX, y) = aFloat(finalX, y) * boundaryScale
                
                For x = finalX To 1 Step -1
                    quickX = (x - 1)
                    aFloat(quickX, y) = aFloat(quickX, y) + nu * aFloat(x, y)
                Next x
            
            End If
            
        Next step
        
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
        
    Next y
    
    'Now repeat all the above steps, but filtering vertically along each column, instead
    If Not g_cancelCurrentAction Then
    
        For x = initX To finalX
            
            For step = 0 To numSteps - 1
                
                'Set initial values
                rFloat(x, initY) = rFloat(x, initY) * boundaryScale
                gFloat(x, initY) = gFloat(x, initY) * boundaryScale
                bFloat(x, initY) = bFloat(x, initY) * boundaryScale
                
                'Filter down
                For y = initY + 1 To finalY
                    QuickY = (y - 1)
                    rFloat(x, y) = rFloat(x, y) + nu * rFloat(x, QuickY)
                    gFloat(x, y) = gFloat(x, y) + nu * gFloat(x, QuickY)
                    bFloat(x, y) = bFloat(x, y) + nu * bFloat(x, QuickY)
                Next y
                
                'Fix closing column values
                rFloat(x, finalY) = rFloat(x, finalY) * boundaryScale
                gFloat(x, finalY) = gFloat(x, finalY) * boundaryScale
                bFloat(x, finalY) = bFloat(x, finalY) * boundaryScale
                
                'Filter up
                For y = finalY To 1 Step -1
                    QuickY = y - 1
                    rFloat(x, QuickY) = rFloat(x, QuickY) + nu * rFloat(x, y)
                    gFloat(x, QuickY) = gFloat(x, QuickY) + nu * gFloat(x, y)
                    bFloat(x, QuickY) = bFloat(x, QuickY) + nu * bFloat(x, y)
                Next y
                
                'Handle alpha separately
                If hasAlpha Then
                    
                    aFloat(x, initY) = aFloat(x, initY) * boundaryScale
                    
                    For y = initY + 1 To finalY
                        aFloat(x, y) = aFloat(x, y) + nu * aFloat(x, y - 1)
                    Next y
                    
                    aFloat(x, finalY) = aFloat(x, finalY) * boundaryScale
                    
                    For y = finalY To 1 Step -1
                        QuickY = y - 1
                        aFloat(x, QuickY) = aFloat(x, QuickY) + nu * aFloat(x, y)
                    Next y
                    
                End If
                
            Next step
            
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + iHeight + modifyProgBarOffset
                End If
            End If
        
        Next x
        
    End If
    
    'Apply final post-scaling
    If Not g_cancelCurrentAction Then
        
        For x = initX To finalX
            quickX = x * qvDepth
        For y = initY To finalY
        
            r = rFloat(x, y) * postScale
            g = gFloat(x, y) * postScale
            b = bFloat(x, y) * postScale
            
            'Perform failsafe clipping
            If (r > 255) Then r = 255
            If (g > 255) Then g = 255
            If (b > 255) Then b = 255
            
            imageData(quickX, y) = b
            imageData(quickX + 1, y) = g
            imageData(quickX + 2, y) = r
            
            'Handle alpha separately
            If hasAlpha Then
                a = aFloat(x, y) * postScale
                If (a > 255) Then a = 255
                imageData(quickX + 3, y) = a
            End If
        
        Next y
        Next x
        
    End If
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then GaussianBlur_IIRImplementation = 0 Else GaussianBlur_IIRImplementation = 1

End Function

'Horizontal blur filter, using an IIR (Infininte Impulse Response) approach.
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
Public Function HorizontalBlur_IIR(ByRef srcDIB As pdDIB, ByVal radius As Double, ByVal numSteps As Long, Optional ByVal blurSymmetric As Boolean = True, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
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
    Dim hasAlpha As Boolean
    hasAlpha = CBool(qvDepth = 4)
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If modifyProgBarMax = -1 Then modifyProgBarMax = srcDIB.GetDIBWidth
    If Not suppressMessages Then SetProgBarMax modifyProgBarMax
    
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Prep some IIR-specific values next
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    Dim step As Long
    
    'Calculate sigma from the radius, using the same formula we do for PD's pure gaussian blur
    Dim sigma As Double
    sigma = Sqr(-(radius * radius) / (2 * Log(1# / 255#)))
    
    'Another possible sigma formula, per this link (http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur):
    'sigma = (radius + 1) / Sqr(2 * (Log(255) / Log(10)))
    
    'Make sure sigma and steps are valid
    If (sigma <= 0#) Then sigma = 0.01
    If (numSteps <= 0) Then numSteps = 1
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate lambda calculation
    ' is proposed.  This adjustment doesn't affect running time at all, and should reduce errors relative to a pure Gaussian.
    ' The behavior could be toggled by the caller, but for now, I've hard-coded use of the modified formula.
    Dim useModifiedQ As Boolean, q As Single
    useModifiedQ = True
    
    If useModifiedQ Then
        q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
    Else
        q = sigma
    End If
    
    'Calculate IIR values
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
    ReDim rFloat(initX To finalX, initY To finalY) As Single
    ReDim gFloat(initX To finalX, initY To finalY) As Single
    ReDim bFloat(initX To finalX, initY To finalY) As Single
    
    If hasAlpha Then ReDim aFloat(initX To finalX, initY To finalY) As Single
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Copy the contents of the current image into float arrays
    For x = initX To finalX
        quickX = x * qvDepth
    For y = initY To finalY
        
        b = imageData(quickX, y)
        g = imageData(quickX + 1, y)
        r = imageData(quickX + 2, y)
        
        rFloat(x, y) = r * ONE_DIV_255
        gFloat(x, y) = g * ONE_DIV_255
        bFloat(x, y) = b * ONE_DIV_255
        
        If hasAlpha Then
            a = imageData(quickX + 3, y)
            aFloat(x, y) = a * ONE_DIV_255
        End If

    Next y
    Next x
    
    '/* Filter horizontally along each row */
    For y = initY To finalY
    
        For step = 0 To numSteps - 1
            
            'Set initial values
            rFloat(initX, y) = rFloat(initX, y) * boundaryScale
            gFloat(initX, y) = gFloat(initX, y) * boundaryScale
            bFloat(initX, y) = bFloat(initX, y) * boundaryScale
            
            'Filter right
            For x = initX + 1 To finalX
                quickX2 = (x - 1)
                rFloat(x, y) = rFloat(x, y) + nu * rFloat(quickX2, y)
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(quickX2, y)
                bFloat(x, y) = bFloat(x, y) + nu * bFloat(quickX2, y)
            Next x
            
            'Filter left only if symmetric
            If blurSymmetric Then
                            
                'Fix closing row
                rFloat(finalX, y) = rFloat(finalX, y) * boundaryScale
                gFloat(finalX, y) = gFloat(finalX, y) * boundaryScale
                bFloat(finalX, y) = bFloat(finalX, y) * boundaryScale
                
                For x = finalX To 1 Step -1
                    quickX = (x - 1)
                    rFloat(quickX, y) = rFloat(quickX, y) + nu * rFloat(x, y)
                    gFloat(quickX, y) = gFloat(quickX, y) + nu * gFloat(x, y)
                    bFloat(quickX, y) = bFloat(quickX, y) + nu * bFloat(x, y)
                Next x
                
            End If
            
            'Apply alpha separately
            If hasAlpha Then
                
                aFloat(initX, y) = aFloat(initX, y) * boundaryScale
                
                For x = initX + 1 To finalX
                    aFloat(x, y) = aFloat(x, y) + nu * aFloat(x - 1, y)
                Next x
                
                If blurSymmetric Then
                    aFloat(finalX, y) = aFloat(finalX, y) * boundaryScale
                    For x = finalX To 1 Step -1
                        quickX = (x - 1)
                        aFloat(quickX, y) = aFloat(quickX, y) + nu * aFloat(x, y)
                    Next x
                End If
            
            End If
            
        Next step
        
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
        
    Next y
    
    'Apply final post-scaling
    If Not g_cancelCurrentAction Then
        
        For y = initY To finalY
        For x = initX To finalX
            
            r = rFloat(x, y) * postScale
            g = gFloat(x, y) * postScale
            b = bFloat(x, y) * postScale
            
            'Perform failsafe clipping
            If (r > 255) Then r = 255
            If (g > 255) Then g = 255
            If (b > 255) Then b = 255
            
            quickX = x * qvDepth
            imageData(quickX, y) = b
            imageData(quickX + 1, y) = g
            imageData(quickX + 2, y) = r
            
            'Handle alpha separately
            If hasAlpha Then
                a = aFloat(x, y) * postScale
                If (a > 255) Then a = 255
                imageData(quickX + 3, y) = a
            End If
        
        Next x
        Next y
        
    End If
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then HorizontalBlur_IIR = 0 Else HorizontalBlur_IIR = 1

End Function
