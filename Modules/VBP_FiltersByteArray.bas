Attribute VB_Name = "Filters_ByteArray"
'***************************************************************************
'Byte Array Filters Module
'Copyright 2014-2015 by Tanner Helland
'Created: 02/April/15
'Last updated: 02/April/15
'Last update: start assembling byte-array-specific filter collection
'
'After version 6.6 released, I started work on a number of modified PD filters.  Unlike most filters in the project
' - which explicitly operate on pdDIB objects - these filters are modified to run on single-channel 2D byte arrays.
' In some places in the project, byte arrays contain sufficient detail for things like bumpmaps, and they consume
' less memory and are faster to process than a three- or four-channel DIB.
'
'Going forward, it would be nice to further expand this function collection, but in the meantime, just know that
' the functions in this module cannot directly operate on images.  Also, they do not generally support progress
' bar reports (by design) as their emphasis is on raw performance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Gaussian blur filter, using an IIR (Infininte Impulse Response) approach
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
'
'IIR provides many benefits over a naive Gaussian Blur implementation:
' - It's performed in-place, meaning a second array is not required.  (That said, the function requires floating-point data,
'    so an intermediate float-type array *is* currently needed.)
' - It approaches a true Gaussian over multiple iterations, but at low iterations, it provides a closer estimate than the
'    corresponding box blur filter would.
' - Floating-point radii are supported.
' - Most importantly, it's much faster to calculate!
'
'Note that the incoming arrayWidth and arrayHeight parameters are 1-based, so this function will automatically subtract 1 to arrive
' at an actual UBound value.
Public Function GaussianBlur_IIR_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal radius As Double, ByVal numSteps As Long) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here.
    ' (In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.)
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim iWidth As Long, iHeight As Long
    iWidth = arrayWidth
    iHeight = arrayHeight
    
    'Prep some IIR-specific values next
    Dim g As Long
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    Dim i As Long, step As Long
    
    'Calculate sigma from the radius, using the same formula we do for PD's pure gaussian blur
    Dim sigma As Double
    sigma = Sqr(-(radius * radius) / (2 * Log(1# / 255#)))
    
    'Another possible sigma formula, per this link (http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur):
    'sigma = (radius + 1) / Sqr(2 * (Log(255) / Log(10)))
    
    'Make sure sigma and steps are valid
    If sigma <= 0 Then sigma = 0.01
    If numSteps <= 0 Then numSteps = 1
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate lambda calculation
    ' is proposed.  This adjustment doesn't affect running time at all, and should reduce errors relative to a pure Gaussian,
    ' allowing for better results with fewer iterations.
    
    'This behavior could theroetically be toggled by the caller, but for now, I've hard-coded use of the modified formula.
    Dim useModifiedQ As Boolean, q As Single
    useModifiedQ = True
    
    If useModifiedQ Then
        q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
    Else
        q = sigma
    End If
    
    'Calculate IIR values
    lambda = (q * q) / (2 * numSteps)
    dnu = (1 + 2 * lambda - Sqr(1 + 4 * lambda)) / (2 * lambda)
    nu = dnu
    boundaryScale = (1 / (1 - dnu))
    postScale = ((dnu / lambda) ^ (2 * numSteps)) * 255
    
    'Intermediate float arrays are required for an IIR transform.
    Dim gFloat() As Single
    ReDim gFloat(initX To finalX, initY To finalY) As Single
    
    'Copy the contents of the incoming byte array into the float array
    For x = initX To finalX
    For y = initY To finalY
        g = srcArray(x, y)
        gFloat(x, y) = g / 255
    Next y
    Next x
    
    'Filter horizontally along each row
    For y = initY To finalY
    
        For step = 0 To numSteps - 1
            
            'Set initial values
            gFloat(initX, y) = gFloat(initX, y) * boundaryScale
            
            'Filter right
            For x = initX + 1 To finalX
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(x - 1, y)
            Next x
            
            'Fix closing row
            gFloat(finalX, y) = gFloat(finalX, y) * boundaryScale
            
            'Filter left
            For x = finalX To 1 Step -1
                gFloat(x - 1, y) = gFloat(x - 1, y) + nu * gFloat(x, y)
            Next x
            
        Next step
        
    Next y
    
    'Now repeat all the above steps, but filtering vertically along each column, instead
    For x = initX To finalX
        
        For step = 0 To numSteps - 1
            
            'Set initial values
            gFloat(x, initY) = gFloat(x, initY) * boundaryScale
            
            'Filter down
            For y = initY + 1 To finalY
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(x, y - 1)
            Next y
                
            'Fix closing column values
            gFloat(x, finalY) = gFloat(x, finalY) * boundaryScale
                
            'Filter up
            For y = finalY To 1 Step -1
                gFloat(x, y - 1) = gFloat(x, y - 1) + nu * gFloat(x, y)
            Next y
            
        Next step
        
    Next x
    
    'Apply final post-scaling
    For x = initX To finalX
    For y = initY To finalY
    
        'Apply post-scaling and perform failsafe clipping.  (Shouldn't technically be necessary, but better safe than sorry.)
        g = gFloat(x, y) * postScale
        If g > 255 Then g = 255
        
        'Store the final value back into the source array
        srcArray(x, y) = g
        
    Next y
    Next x
    
    GaussianBlur_IIR_ByteArray = True
    
End Function

'Horizontal box blur; single-pass only.  Left and right amounts can be independently specified.
Public Function HorizontalBlur_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal lRadius As Long, ByVal rRadius As Long) As Boolean
    
    'A second copy of the array is required, to prevent already blurred values from screwing up future calculations
    Dim srcArrayCopy() As Byte
    ReDim srcArrayCopy(0 To arrayWidth - 1, 0 To arrayHeight - 1) As Byte
    CopyMemory ByVal VarPtr(srcArrayCopy(0, 0)), ByVal VarPtr(srcArray(0, 0)), arrayWidth * arrayHeight
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here.
    ' (In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.)
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim xRadius As Long
    xRadius = finalX - initX
    
    'Limit the left and right offsets to the width of the image
    If lRadius > xRadius Then lRadius = xRadius
    If rRadius > xRadius Then rRadius = xRadius
        
    'The number of pixels in the current horizontal line are tracked dynamically.
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbX As Long, ubX As Long
    Dim obuX As Boolean
    
    'This horizontal blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each horizontal line, and only update it when we move one column to the right.
    Dim gTotals() As Long
    ReDim gTotals(initY To finalY) As Long
    
    'Populate the initial arrays.  We can ignore the left offset at this point, as we are starting at column 0 (and there are no
    ' pixels left of that!)
    If rRadius > 0 Then
    
        For x = initX To initX + rRadius - 1
        For y = initY To finalY
            gTotals(y) = gTotals(y) + srcArrayCopy(x, y)
        Next y
            
            'Increase the pixel tally on a per-column basis
            NumOfPixels = NumOfPixels + 1
            
        Next x
        
    End If
                
    'Loop through each column in the image, tallying blur values as we go
    For x = initX To finalX
        
        'Determine the loop bounds of the current blur box in the X direction
        lbX = x - lRadius
        If lbX < 0 Then lbX = 0
        ubX = x + rRadius
        
        If ubX > finalX Then
            obuX = True
            ubX = finalX
        Else
            obuX = False
        End If
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        If lbX > 0 Then
            
            For y = initY To finalY
                gTotals(y) = gTotals(y) - srcArrayCopy(lbX - 1, y)
            Next y
            
            NumOfPixels = NumOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        If Not obuX Then
            
            For y = initY To finalY
                gTotals(y) = gTotals(y) + srcArrayCopy(ubX, y)
            Next y
            
            NumOfPixels = NumOfPixels + 1
            
        End If
            
        'Process the current column.  This simply involves calculating blur values, and applying them to the destination array
        For y = initY To finalY
            srcArray(x, y) = gTotals(y) \ NumOfPixels
        Next y
        
    Next x
    
    HorizontalBlur_ByteArray = True
    
End Function

'Vertical box blur; single-pass only.  Up and down amounts can be independently specified.
Public Function VerticalBlur_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal uRadius As Long, ByVal dRadius As Long) As Boolean
    
    'A second copy of the array is required, to prevent already blurred values from screwing up future calculations
    Dim srcArrayCopy() As Byte
    ReDim srcArrayCopy(0 To arrayWidth - 1, 0 To arrayHeight - 1) As Byte
    CopyMemory ByVal VarPtr(srcArrayCopy(0, 0)), ByVal VarPtr(srcArray(0, 0)), arrayWidth * arrayHeight
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here.
    ' (In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.)
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim yRadius As Long
    yRadius = finalY - initY
    
    'Limit the up and down offsets to the height of the image
    If uRadius > yRadius Then uRadius = yRadius
    If dRadius > yRadius Then dRadius = yRadius
        
    'The number of pixels in the current vertical line are tracked dynamically.
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbY As Long, ubY As Long
    Dim obuY As Boolean
        
    'This vertical blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each vertical line, and only update it when we move one row down.
    Dim gTotals() As Long
    ReDim gTotals(initX To finalX) As Long
    
    'Populate the initial array.  We can ignore the up offset at this point, as we are starting at row 0 (and there are no
    ' pixels above that!)
    If dRadius > 0 Then
        
        For y = initY To initY + dRadius - 1
        For x = initX To finalX
            gTotals(x) = gTotals(x) + srcArrayCopy(x, y)
        Next x
        
            'Increase the pixel tally on a per-column basis
            NumOfPixels = NumOfPixels + 1
            
        Next y
        
    End If
                
    'Loop through each row in the image, tallying blur values as we go
    For y = initY To finalY
        
        'Determine the loop bounds of the current blur box in the Y direction
        lbY = y - uRadius
        If lbY < 0 Then lbY = 0
        ubY = y + dRadius
        
        If ubY > finalY Then
            obuY = True
            ubY = finalY
        Else
            obuY = False
        End If
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        If lbY > 0 Then
            
            For x = initX To finalX
                gTotals(x) = gTotals(x) - srcArrayCopy(x, lbY - 1)
            Next x
            
            NumOfPixels = NumOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        If Not obuY Then
        
            For x = initX To finalX
                gTotals(x) = gTotals(x) + srcArrayCopy(x, ubY)
            Next x
            
            NumOfPixels = NumOfPixels + 1
            
        End If
            
        'Process the current row.  This simply involves calculating blur values, and applying them to the destination image.
        For x = initX To finalX
            srcArray(x, y) = gTotals(x) \ NumOfPixels
        Next x
        
    Next y
    
    VerticalBlur_ByteArray = True
    
End Function

'Given a 2D byte array, normalize the contents to guarantee a full stretch on the range [0, 255]
Public Function normalizeByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim g As Long, minVal As Long, maxVal As Long
    minVal = 255
    maxVal = 0
    
    'Start by finding max/min values in the current array
    For x = initX To finalX
    For y = initY To finalY
                
        g = srcArray(x, y)
        If g < minVal Then
            minVal = g
        ElseIf g > maxVal Then
            maxVal = g
        End If
    
    Next y
    Next x
        
    'If the data isn't normalized, normalize it now
    If (minVal > 0) Or (maxVal < 255) Then
        
        Dim curRange As Long
        curRange = maxVal - minVal
        
        If curRange = 0 Then curRange = 1
        
        'Build a normalization lookup table
        Dim normalizedLookup() As Byte
        ReDim normalizedLookup(0 To 255) As Byte
        
        Dim newValue As Long
        
        For x = 0 To 255
        
            newValue = (CDbl(x - minVal) / CDbl(curRange)) * 255
            
            If newValue < 0 Then
                newValue = 0
            ElseIf newValue > 255 Then
                newValue = 255
            End If
            
            normalizedLookup(x) = newValue
            
        Next x
            
        'Apply normalization
        For x = initX To finalX
        For y = initY To finalY
            srcArray(x, y) = normalizedLookup(srcArray(x, y))
        Next y
        Next x
    
    End If
    
    normalizeByteArray = True
    
End Function

'Add noise to a byte array.  noiseAmount is on the range [0, 255]; it will be auto-converted to [-255, 255] when applying changes
' to the image.
Public Function addNoiseByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal noiseAmount As Long) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Double noise and calculate a half value
    Dim halfNoise As Long
    halfNoise = noiseAmount
    noiseAmount = noiseAmount * 2
    
    Dim oldValue As Long, newValue As Long
    
    'Add noise to each point in the array
    For x = initX To finalX
    For y = initY To finalY
                
        oldValue = srcArray(x, y)
        
        newValue = oldValue + ((Rnd * noiseAmount) - halfNoise)
        
        If newValue < 0 Then
            newValue = 0
        ElseIf newValue > 255 Then
            newValue = 255
        End If
        
        srcArray(x, y) = newValue
    
    Next y
    Next x
    
    addNoiseByteArray = True
    
End Function

'Given a byte array, convert all values to 0 or 255 using a user-supplied threshold value.  If the autoCalculateThreshold value is TRUE,
' the array will be scanned and the median value of the array will be used as the threshold.  This should result in an image with a
' relatively even split between white and black pixels.
Public Function thresholdByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal thresholdValue As Long = 127, Optional ByVal autoCalculateThreshold As Boolean = False) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim curValue As Long
    
    'If auto-calculate was specified, find the array's mean value now.
    If autoCalculateThreshold Then
    
        Dim gHistogram() As Long
        ReDim gHistogram(0 To 255) As Long
        
        For x = initX To finalX
        For y = initY To finalY
            
            curValue = srcArray(x, y)
            gHistogram(curValue) = gHistogram(curValue) + 1
            
        Next y
        Next x
        
        'Find the median value of the histogram
        Dim halfTotalValues As Long
        halfTotalValues = (arrayWidth * arrayHeight) \ 2
        
        curValue = 0
        For x = 0 To 255
        
            curValue = curValue + gHistogram(x)
            If curValue >= halfTotalValues Then
                thresholdValue = x
                Exit For
            End If
        
        Next x
        
        Erase gHistogram
    
    End If
    
    'Threshold the array
    For x = initX To finalX
    For y = initY To finalY
        
        If srcArray(x, y) >= thresholdValue Then
            srcArray(x, y) = 255
        Else
            srcArray(x, y) = 0
        End If
        
    Next y
    Next x
    
    thresholdByteArray = True
    
End Function

'Given a byte array, convert all values to 0 or 255 using a user-supplied threshold value and a given dithering type (currently
' Floyd-Steinberg, but nothing prevents the use of other kernels).
'
'If the autoCalculateThreshold value is TRUE, the array will be scanned and the median value of the array will be used as the threshold.
' (This should result in an image with a relatively even split between white and black pixels.)
Public Function thresholdPlusDither_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal thresholdValue As Long = 127, Optional ByVal autoCalculateThreshold As Boolean = False) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim curValue As Long
    
    'If auto-calculate was specified, find the array's mean value now.
    If autoCalculateThreshold Then
    
        Dim gHistogram() As Long
        ReDim gHistogram(0 To 255) As Long
        
        For x = initX To finalX
        For y = initY To finalY
            
            curValue = srcArray(x, y)
            gHistogram(curValue) = gHistogram(curValue) + 1
            
        Next y
        Next x
        
        'Find the median value of the histogram
        Dim halfTotalValues As Long
        halfTotalValues = (arrayWidth * arrayHeight) \ 2
        
        curValue = 0
        For x = 0 To 255
        
            curValue = curValue + gHistogram(x)
            If curValue >= halfTotalValues Then
                thresholdValue = x
                Exit For
            End If
        
        Next x
        
        Erase gHistogram
    
    End If
    
    'Prep a dither table.  Note that any tables from the masterBlackWhiteConversion function could be used, but at present,
    ' I've hard-coded this function against Floyd-Steinberg dithering.
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    
    Dim DitherTable() As Byte
    ReDim DitherTable(-1 To 1, 0 To 1) As Byte
            
    DitherTable(1, 0) = 7
    DitherTable(-1, 1) = 3
    DitherTable(0, 1) = 5
    DitherTable(1, 1) = 1
    
    dDivisor = 16

    'Next, mark the relevant size of the dither table in the left, right, and down directions
    xLeft = -1
    xRight = 1
    yDown = 1
    
    'Next, create a dithering table the same size as the source array.  We make it of Single type to prevent rounding errors.
    ' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
    Dim dErrors() As Single
    ReDim dErrors(initX To finalX, initY To finalY) As Single
    
    Dim i As Long, j As Long
    Dim g As Long, newG As Long
    Dim QuickX As Long, QuickY As Long
    
    'Now loop through the array, calculating errors as we go
    For x = initX To finalX
    For y = initY To finalY
        
        g = srcArray(x, y)
        
        'Add in the current running error for this pixel
        newG = g + dErrors(x, y)
        
        'Check our modified luminance value against the threshold, and set new values accordingly
        If newG >= thresholdValue Then
            errorVal = newG - 255
            srcArray(x, y) = 255
        Else
            errorVal = newG
            srcArray(x, y) = 0
        End If
        
        'If there is an error, spread it according to the dither table formula
        If errorVal <> 0 Then
        
            For i = xLeft To xRight
            For j = 0 To yDown
            
                'First, ignore already processed pixels
                If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                
                'Second, ignore pixels that have a zero in the dither table
                If DitherTable(i, j) = 0 Then GoTo NextDitheredPixel
                
                QuickX = x + i
                QuickY = y + j
                
                'Next, ignore target pixels that are off the image boundary
                If QuickX < initX Then GoTo NextDitheredPixel
                If QuickX > finalX Then GoTo NextDitheredPixel
                If QuickY > finalY Then GoTo NextDitheredPixel
                
                'If we've made it all the way here, we are able to actually spread the error to this location
                dErrors(QuickX, QuickY) = dErrors(QuickX, QuickY) + (errorVal * (CSng(DitherTable(i, j)) / dDivisor))
            
NextDitheredPixel:
            Next j
            Next i
        
        End If
            
    Next y
    Next x
    
    thresholdPlusDither_ByteArray = True
    
End Function

'Given a byte array, reduce the number of available values to some number specified by the user (e.g. "2 shades" = monochrome,
' "4 shades" equals black, dark gray, light gray, white), with dithering.  (Currently dithering is limited to Floyd-Steinberg
' Floyd-Steinberg, but nothing prevents the use of other kernels).
Public Function Dither_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal numOfShades As Long = 4, Optional ByVal autoCalculateThreshold As Boolean = False) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build a look-up table for our custom conversion
    Dim g As Long, newG As Long
    Dim grayLookup() As Byte
    ReDim grayLookup(0 To 255) As Byte
    
    For x = 0 To 255
        g = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If g > 255 Then g = 255
        grayLookup(x) = g
    Next x
    
    'Prep a dither table.  Note that any tables from the masterBlackWhiteConversion function could be used, but at present,
    ' I've hard-coded this function against Floyd-Steinberg dithering.
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    
    Dim DitherTable() As Byte
    ReDim DitherTable(-1 To 1, 0 To 1) As Byte
            
    DitherTable(1, 0) = 7
    DitherTable(-1, 1) = 3
    DitherTable(0, 1) = 5
    DitherTable(1, 1) = 1
    
    dDivisor = 16

    'Next, mark the relevant size of the dither table in the left, right, and down directions
    xLeft = -1
    xRight = 1
    yDown = 1
    
    'Next, create a dithering table the same size as the source array.  We make it of Single type to prevent rounding errors.
    ' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
    Dim dErrors() As Single
    ReDim dErrors(initX To finalX, initY To finalY) As Single
    
    Dim i As Long, j As Long
    Dim QuickX As Long, QuickY As Long
    
    'Now loop through the array, calculating errors as we go
    For x = initX To finalX
    For y = initY To finalY
        
        g = srcArray(x, y)
        
        'Add in the current running error for this pixel
        g = g + dErrors(x, y)
        
        'Convert to a lookup-table safe value
        If g >= 255 Then
            newG = 255
        ElseIf g < 0 Then
            newG = 0
        Else
            newG = g
        End If
        
        'Write out the new luminance
        srcArray(x, y) = grayLookup(newG)
        
        'Calculate an error
        errorVal = g - grayLookup(newG)
        
        'If there is an error, spread it according to the dither table formula
        If errorVal <> 0 Then
        
            For i = xLeft To xRight
            For j = 0 To yDown
            
                'First, ignore already processed pixels
                If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                
                'Second, ignore pixels that have a zero in the dither table
                If DitherTable(i, j) = 0 Then GoTo NextDitheredPixel
                
                QuickX = x + i
                QuickY = y + j
                
                'Next, ignore target pixels that are off the image boundary
                If QuickX < initX Then GoTo NextDitheredPixel
                If QuickX > finalX Then GoTo NextDitheredPixel
                If QuickY > finalY Then GoTo NextDitheredPixel
                
                'If we've made it all the way here, we are able to actually spread the error to this location
                dErrors(QuickX, QuickY) = dErrors(QuickX, QuickY) + (errorVal * (CSng(DitherTable(i, j)) / dDivisor))
            
NextDitheredPixel:
            Next j
            Next i
        
        End If
            
    Next y
    Next x
    
    Dither_ByteArray = True
    
End Function

'Contrast-correct a byte array.  (This function is based off PD's white balance algorithm.)
Public Function ContrastCorrect_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal percentIgnore As Double) As Long

    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Color values
    Dim g As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim lMax As Byte, lMin As Byte
    lMax = 0
    lMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    percentIgnore = percentIgnore / 100
    
    'Prepare a histogram array
    Dim lCount() As Long
    ReDim lCount(0 To 255) As Long
    
    'Build an initial histogram
    For x = initX To finalX
    For y = initY To finalY
    
        'Increment the histogram at this position
        g = srcArray(x, y)
        lCount(g) = lCount(g) + 1
        
    Next y
    Next x
    
     'With the histogram complete, we can now figure out how to stretch the gray map. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    Dim foundYet As Boolean
    foundYet = False
    
    Dim NumOfPixels As Long
    NumOfPixels = arrayWidth * arrayHeight
    
    Dim wbThreshold As Long
    wbThreshold = NumOfPixels * percentIgnore
    
    g = 0
    
    Dim lTally As Long
    lTally = 0
    
    'Find minimum and maximum luminance values in the current image
    Do
        If lCount(g) + lTally < wbThreshold Then
            g = g + 1
            lTally = lTally + lCount(g)
        Else
            lMin = g
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
    
    g = 255
    lTally = 0
    
    Do
        If lCount(g) + lTally < wbThreshold Then
            g = g - 1
            lTally = lTally + lCount(g)
        Else
            lMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Calculate the difference between max and min
    Dim lDif As Long
    lDif = CLng(lMax) - CLng(lMin)
    
    'Build a final set of look-up tables that contain the results of the requisite luminance transformation
    Dim lFinal() As Byte
    ReDim lFinal(0 To 255) As Byte
    
    For x = 0 To 255
    
        If lDif <> 0 Then g = 255 * ((x - lMin) / lDif) Else g = x
        
        If g < 0 Then
            g = 0
        ElseIf g > 255 Then
            g = 255
        End If
        
        lFinal(x) = g
        
    Next x
    
    'Now we can loop through each entry in the array, converting values as we go
    For x = initX To finalX
    For y = initY To finalY
        srcArray(x, y) = lFinal(srcArray(x, y))
    Next y
    Next x
    
    ContrastCorrect_ByteArray = True
    
End Function
