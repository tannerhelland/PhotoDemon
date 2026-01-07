Attribute VB_Name = "Filters_ByteArray"
'***************************************************************************
'Byte Array Filters Module
'Copyright 2014-2026 by Tanner Helland
'Created: 02/April/15
'Last updated: 23/August/22
'Last update: minor perf improvements to dilate/erode
'
'This module contains various image filters, but rewritten to work on single-channel byte arrays
' instead of three- or four-channel RGB/A images.  (In some cases, PD can use these filters for
' meaningful perf improvements over multichannel implementations - like operations on a mask.)
'
'Going forward, it would be nice to further expand this collection, but for now just know that
' the functions in this module cannot directly operate on images.  Also, they do not generally
' support progress bar reports (by design) as their emphasis is on raw performance.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Gaussian blur filter, using an IIR (Infininte Impulse Response) approach
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
'
'This function is a stripped-down version of the full RGBA implementation in the
' Filters_Area.GaussianBlur_AM() function.  Please look there for full implementation details
' (and comments).
Public Function GaussianBlur_AM_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal radius As Double, ByVal numSteps As Long) As Boolean
    
    'Fudge to produce results closer to an iterative box blur estimation
    radius = radius * 1.075
    
    'In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
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
    
    'Prep some IIR-specific values next
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    lambda = (q * q) / (2# * numSteps)
    dnu = (1# + 2# * lambda - Sqr(1# + 4# * lambda)) / (2# * lambda)
    nu = dnu
    boundaryScale = (1# / (1# - dnu))
    postScale = ((dnu / lambda) ^ (2# * numSteps))
    
    Dim step As Long
    
    Dim numPixels As Long
    numPixels = arrayWidth * arrayHeight
    
    Dim tmpFloat() As Single
    ReDim tmpFloat(0 To numPixels - 1) As Single
    
    Dim origValue As Long, xOffset As Long
    
    'Convert the image to floats
    For y = 0 To finalY
        xOffset = y * arrayWidth
        For x = 0 To finalX
            tmpFloat(x + xOffset) = srcArray(x, y)
        Next x
    Next y
    
    'Filter horizontally along each row
    For y = 0 To finalY
    
        xOffset = y * arrayWidth
    
        For step = 0 To numSteps - 1
            
            'Set initial values
            tmpFloat(xOffset) = tmpFloat(xOffset) * boundaryScale
            
            'Filter right
            For x = 1 To finalX
                tmpFloat(xOffset + x) = tmpFloat(xOffset + x) + nu * tmpFloat(xOffset + x - 1)
            Next x
            
            'Fix closing row
            tmpFloat(xOffset + finalX) = tmpFloat(xOffset + finalX) * boundaryScale
            
            'Filter left
            For x = finalX To 1 Step -1
                tmpFloat(xOffset + x - 1) = tmpFloat(xOffset + x - 1) + nu * tmpFloat(xOffset + x)
            Next x
            
        Next step
        
    Next y
    
    'Now repeat all the above steps, but filtering vertically along each column, instead
    For step = 0 To numSteps - 1
        
        'Set initial values
        For x = 0 To finalX
            tmpFloat(x) = tmpFloat(x) * boundaryScale
        Next x
        
        'Filter down
        For y = 1 To finalY
            xOffset = y * arrayWidth
            For x = 0 To finalX
                tmpFloat(xOffset + x) = tmpFloat(xOffset + x) + nu * tmpFloat(xOffset + x - arrayWidth)
            Next x
        Next y
            
        'Fix closing column values
        xOffset = finalY * arrayWidth
        For x = 0 To finalX
            tmpFloat(xOffset + x) = tmpFloat(xOffset + x) * boundaryScale
        Next x
            
        'Filter up
        For y = finalY To 1 Step -1
            xOffset = (y - 1) * arrayWidth
            For x = 0 To finalX
                tmpFloat(xOffset + x) = tmpFloat(xOffset + x) + nu * tmpFloat(xOffset + arrayWidth + x)
            Next x
        Next y
        
    Next step
        
    'Apply final post-scaling
    For y = 0 To finalY
        xOffset = y * arrayWidth
    For x = 0 To finalX
        
        'Round the finished result, perform failsafe clipping, then assign
        origValue = Int(tmpFloat(xOffset + x) * postScale + 0.5)
        If (origValue > 255) Then origValue = 255
        srcArray(x, y) = origValue
        
    Next x
    Next y
    
    GaussianBlur_AM_ByteArray = True
    
End Function

'Horizontal box blur; single-pass only.  Left and right amounts can be independently specified.
Public Function HorizontalBlur_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal lRadius As Long, ByVal rRadius As Long) As Boolean
    
    'A second copy of the array is required, to prevent already blurred values from screwing up future calculations
    Dim srcArrayCopy() As Byte
    ReDim srcArrayCopy(0 To arrayWidth * arrayHeight - 1) As Byte
    CopyMemoryStrict VarPtr(srcArrayCopy(0)), VarPtr(srcArray(0, 0)), arrayWidth * arrayHeight
    
    'In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Limit the left and right offsets to the width of the image
    If (lRadius > finalX) Then lRadius = finalX
    If (rRadius > finalX) Then rRadius = finalX
        
    'The number of pixels in the current horizontal line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = lRadius + rRadius + 1
    
    'To achieve better results, we want to round final blur totals.  Cache the equivalent of
    ' 0.5 for the current pixel count.
    Dim halfNumPixels As Long
    halfNumPixels = Int(numOfPixels \ 2)
    
    'This horizontal blur algorithm is based on the principle of "not redoing work that's
    ' already been done."  To that end, we store the accumulated blur total for the current line,
    ' and only update it when we move to the next column.
    Dim lbX As Long, ubX As Long, lineOffset As Long
    Dim gTotal As Long, gInit As Byte, gFinal As Byte
    
    'Populate the initial trackers.  We can ignore the left offset at this point, as we are starting
    ' at column 0 (and there are no pixels left of that!)
    For y = 0 To finalY
        
        'Reset all line trackers
        gTotal = 0
        
        'Populate the initial accumulators
        lineOffset = y * arrayWidth
        
        'Make a note of the first r/g/b/a values in the line; this allows us to skip
        ' (relatively expensive) array accesses for these values.
        gInit = srcArrayCopy(lineOffset)
        gFinal = srcArrayCopy(lineOffset + finalX)
        
        'First, add copies of the left-most pixel (effectively clamping the edges of the blur).
        ' Note that we also add an *extra* copy of the left-most pixel; this allows us to skip a
        ' boundary check on the inner loop.
        gTotal = gTotal + gInit * (lRadius + 1)
        
        'Next, add all pixels in the initial radius
        For x = 0 To rRadius - 1
            gTotal = gTotal + srcArrayCopy(lineOffset + x)
        Next x
        
        'Loop through each column in this row, updating the accumulator as we go
        For x = 0 To finalX
            
            'Remove trailing values from the blur collection if they lie outside the processing radius
            lbX = x - lRadius
            If (lbX > 0) Then
                gTotal = gTotal - srcArrayCopy(lineOffset + lbX - 1)
            Else
                gTotal = gTotal - gInit
            End If
            
            'Add leading values to the blur box if they lie inside the processing radius
            ubX = x + rRadius
            If (ubX <= finalX) Then
                gTotal = gTotal + srcArrayCopy(lineOffset + ubX)
            Else
                gTotal = gTotal + gFinal
            End If
            
            'Apply the blurred value to the destination image (with rounding).
            srcArray(x, y) = (gTotal + halfNumPixels) \ numOfPixels
            
        Next x
        
    Next y
    
    HorizontalBlur_ByteArray = True
    
End Function

'Vertical box blur; single-pass only.  Up and down amounts can be independently specified.
Public Function VerticalBlur_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal uRadius As Long, ByVal dRadius As Long) As Boolean
    
    'A second copy of the array is required, to prevent already blurred values from screwing up future calculations
    Dim srcArrayCopy() As Byte
    ReDim srcArrayCopy(0 To arrayWidth * arrayHeight - 1) As Byte
    CopyMemoryStrict VarPtr(srcArrayCopy(0)), VarPtr(srcArray(0, 0)), arrayWidth * arrayHeight
    
    'In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Limit the up and down offsets to the height of the image
    If (uRadius > finalY) Then uRadius = finalY
    If (dRadius > finalY) Then dRadius = finalY
        
    'The number of pixels in the current vertical line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = uRadius + dRadius + 1
    
    'To achieve better results, we want to round final blur totals.  Cache the equivalent of
    ' 0.5 for the current pixel count.
    Dim halfNumPixels As Long
    halfNumPixels = Int(numOfPixels \ 2)
    
    'This vertical blur algorithm is based on the principle of "not redoing work that's already been done."
    ' To that end, we store the accumulated blur total for each vertical line, and only update it when we
    ' move one row down.
    Dim lbY As Long, ubY As Long, xOffset As Long
    Dim gTotal() As Long, gInit() As Byte, gFinal() As Byte
    ReDim gTotal(0 To finalX) As Long
    ReDim gInit(0 To finalX) As Byte
    ReDim gFinal(0 To finalX) As Byte
    
    'Make a note of the first and last r/g/b/a values in each line;
    ' this allows us to skip (relatively expensive) array accesses for these values.
    For x = 0 To finalX
        gInit(x) = srcArrayCopy(x)
    Next x
    
    xOffset = finalY * arrayWidth
    For x = 0 To finalX
        gFinal(x) = srcArrayCopy(xOffset + x)
    Next x
    
    'Next, add copies of the top-most pixel (effectively clamping the edges of the blur).
    ' Note that we also add an *extra* copy of the top-most pixel; this allows us to skip a
    ' boundary check on the inner loop.
    For x = 0 To finalX
        gTotal(x) = gInit(x) * (uRadius + 1)
    Next x
    
    'Next, add all pixels in the initial radius
    For y = 0 To dRadius - 1
        xOffset = y * arrayWidth
        For x = 0 To finalX
            gTotal(x) = gTotal(x) + srcArrayCopy(xOffset + x)
        Next x
    Next y
    
    'Loop through each row in the image, tallying blur values as we go
    For y = 0 To finalY
        
        'Remove trailing values from the blur collection if they lie outside the processing radius
        lbY = y - uRadius
        If (lbY > 0) Then
            xOffset = (lbY - 1) * arrayWidth
            For x = 0 To finalX
                gTotal(x) = gTotal(x) - srcArrayCopy(xOffset + x)
            Next x
        Else
            For x = 0 To finalX
                gTotal(x) = gTotal(x) - gInit(x)
            Next x
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        ubY = y + dRadius
        If (ubY <= finalY) Then
            xOffset = ubY * arrayWidth
            For x = 0 To finalX
                gTotal(x) = gTotal(x) + srcArrayCopy(xOffset + x)
            Next x
        Else
            For x = 0 To finalX
                gTotal(x) = gTotal(x) + gFinal(x)
            Next x
        End If
        
        'Apply blurred values to the destination image (with rounding).
        For x = 0 To finalX
            srcArray(x, y) = (gTotal(x) + halfNumPixels) \ numOfPixels
        Next x
        
    Next y
    
    VerticalBlur_ByteArray = True
    
End Function

'Given a 2D byte array, normalize the contents to guarantee a full stretch on the range [0, 255]
Public Function NormalizeByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
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
    
    NormalizeByteArray = True
    
End Function

'Pad the edges of a byte array by some arbitrary amount.  The padded edges will be filled with
' clamped copies of boundary values, making this extremely helpful for area operators.
Public Sub PadByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByRef dstArray() As Byte, Optional ByVal padHorizontal As Long = 0, Optional ByVal padVertical As Long = 0)
    
    'Ensure valid inputs
    If (arrayWidth <= 0) Or (arrayHeight <= 0) Or (padHorizontal < 0) Or (padVertical < 0) Or ((padHorizontal = 0) And (padVertical = 0)) Then Exit Sub
    
    'Calculate new dimensions and prep the destination array
    Dim newWidth As Long, newHeight As Long
    newWidth = arrayWidth + padHorizontal * 2
    newHeight = arrayHeight + padVertical * 2
    
    ReDim dstArray(0 To newWidth - 1, 0 To newHeight - 1) As Byte
    
    'Start with vertical stripes, as we can use memcpy to move the existing bytes into place
    ' (and pad the vertical stripes while we're at it).
    Dim yBound As Long
    yBound = arrayHeight - 1
    
    Dim y As Long
    For y = 0 To newHeight - 1
        If (y <= padVertical) Then
            CopyMemoryStrict VarPtr(dstArray(padHorizontal, y)), VarPtr(srcArray(0, 0)), arrayWidth
        ElseIf (y < arrayHeight + padVertical) Then
            CopyMemoryStrict VarPtr(dstArray(padHorizontal, y)), VarPtr(srcArray(0, y - padVertical)), arrayWidth
        Else
            CopyMemoryStrict VarPtr(dstArray(padHorizontal, y)), VarPtr(srcArray(0, yBound)), arrayWidth
        End If
    Next y
    
    'We now need to pad horizontal bytes, if any.  Start with left padding.
    For y = 0 To newHeight - 1
        If (y <= padVertical) Then
            VBHacks.FillMemory VarPtr(dstArray(0, y)), padHorizontal, srcArray(0, 0)
        ElseIf (y < arrayHeight + padVertical) Then
            VBHacks.FillMemory VarPtr(dstArray(0, y)), padHorizontal, srcArray(0, y - padVertical)
        Else
            VBHacks.FillMemory VarPtr(dstArray(0, y)), padHorizontal, srcArray(0, yBound)
        End If
    Next y
    
    'Repeat above steps, but for right padding.  (And pre-calculate some offsets for perf reasons.)
    Dim xBound As Long, xOffset As Long
    xBound = arrayWidth - 1
    xOffset = padHorizontal + xBound
    
    For y = 0 To newHeight - 1
        If (y <= padVertical) Then
            VBHacks.FillMemory VarPtr(dstArray(xOffset, y)), padHorizontal, srcArray(xBound, 0)
        ElseIf (y < arrayHeight + padVertical) Then
            VBHacks.FillMemory VarPtr(dstArray(xOffset, y)), padHorizontal, srcArray(xBound, y - padVertical)
        Else
            VBHacks.FillMemory VarPtr(dstArray(xOffset, y)), padHorizontal, srcArray(xBound, yBound)
        End If
    Next y
    
End Sub

'Pad the edges of a byte array by some arbitrary amount.  The padded edges will be filled with zeroes.
Public Sub PadByteArray_NoClamp(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByRef dstArray() As Byte, Optional ByVal padHorizontal As Long = 0, Optional ByVal padVertical As Long = 0)
    
    'Ensure valid inputs
    If (arrayWidth <= 0) Or (arrayHeight <= 0) Or (padHorizontal < 0) Or (padVertical < 0) Or ((padHorizontal = 0) And (padVertical = 0)) Then Exit Sub
    
    'Calculate new dimensions and prep the destination array
    Dim newWidth As Long, newHeight As Long
    newWidth = arrayWidth + padHorizontal * 2
    newHeight = arrayHeight + padVertical * 2
    
    ReDim dstArray(0 To newWidth - 1, 0 To newHeight - 1) As Byte
    
    Dim y As Long
    For y = 0 To arrayHeight - 1
        CopyMemoryStrict VarPtr(dstArray(padHorizontal, y + padVertical)), VarPtr(srcArray(0, y)), arrayWidth
    Next y
    
End Sub

'Sister function to various PadByteArray() functions, above.
' NOTE: pass identical values to both functions!  If you don't, this will break!
Public Sub UnPadByteArray(ByRef dstArray() As Byte, ByVal dstArrayWidth As Long, ByVal dstArrayHeight As Long, ByRef srcArray() As Byte, Optional ByVal padHorizontal As Long = 0, Optional ByVal padVertical As Long = 0, Optional ByVal dstIsAlreadySized As Boolean = True)
    
    'Ensure valid inputs
    If (dstArrayWidth <= 0) Or (dstArrayHeight <= 0) Or (padHorizontal < 0) Or (padVertical < 0) Or ((padHorizontal = 0) And (padVertical = 0)) Then Exit Sub
    
    'Prep the destination array and copy scanlines one-at-a-time
    If (Not dstIsAlreadySized) Then ReDim dstArray(0 To dstArrayWidth - 1, 0 To dstArrayHeight - 1) As Byte
    
    Dim y As Long
    For y = 0 To dstArrayHeight - 1
        CopyMemoryStrict VarPtr(dstArray(0, y)), VarPtr(srcArray(padHorizontal, y + padVertical)), dstArrayWidth
    Next y
    
End Sub

'Add noise to a byte array.  noiseAmount is on the range [0, 255]; it will be auto-converted to [-255, 255] when applying changes
' to the image.
Public Function AddNoiseByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal noiseAmount As Long) As Boolean
    
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
        
        If (newValue < 0) Then
            newValue = 0
        ElseIf (newValue > 255) Then
            newValue = 255
        End If
        
        srcArray(x, y) = newValue
    
    Next y
    Next x
    
    AddNoiseByteArray = True
    
End Function

'Given a byte array, convert all values to 0 or 255 using a user-supplied threshold value.  If the autoCalculateThreshold value is TRUE,
' the array will be scanned and the median value of the array will be used as the threshold.  This should result in an image with a
' relatively even split between white and black pixels.
Public Function ThresholdByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal thresholdValue As Long = 127, Optional ByVal autoCalculateThreshold As Boolean = False) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim curValue As Long
    
    'If auto-calculate was specified, find the array's mean value now.
    If autoCalculateThreshold Then
    
        Dim gHistogram(0 To 255) As Long
        
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
        If (srcArray(x, y) >= thresholdValue) Then
            srcArray(x, y) = 255
        Else
            srcArray(x, y) = 0
        End If
    Next y
    Next x
    
    ThresholdByteArray = True
    
End Function

'Given a byte array, convert all values to 0 or 255 using a user-supplied threshold value and
' a given dithering type (currently Floyd-Steinberg, but nothing prevents the use of other kernels).
'
'If the autoCalculateThreshold value is TRUE, the array will be scanned and the median value of the
' array will be used as the threshold.  (This should result in an image with a relatively even split
' between white and black pixels.)
Public Function ThresholdPlusDither_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal thresholdValue As Long = 127, Optional ByVal autoCalculateThreshold As Boolean = False) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim curValue As Long
    
    'If auto-calculate was specified, find the array's mean value now.
    If autoCalculateThreshold Then
    
        Dim gHistogram(0 To 255) As Long
        
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
    
    'Prep a dither table.  Note that any dithering table will work, but at present,
    ' I've hard-coded this function against Floyd-Steinberg dithering.
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    
    Dim ditherTable() As Byte
    ReDim ditherTable(-1 To 1, 0 To 1) As Byte
            
    ditherTable(1, 0) = 7
    ditherTable(-1, 1) = 3
    ditherTable(0, 1) = 5
    ditherTable(1, 1) = 1
    
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
    Dim xStride As Long, quickY As Long
    
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
                If ditherTable(i, j) = 0 Then GoTo NextDitheredPixel
                
                xStride = x + i
                quickY = y + j
                
                'Next, ignore target pixels that are off the image boundary
                If xStride < initX Then GoTo NextDitheredPixel
                If xStride > finalX Then GoTo NextDitheredPixel
                If quickY > finalY Then GoTo NextDitheredPixel
                
                'If we've made it all the way here, we are able to actually spread the error to this location
                dErrors(xStride, quickY) = dErrors(xStride, quickY) + (errorVal * (CSng(ditherTable(i, j)) / dDivisor))
            
NextDitheredPixel:
            Next j
            Next i
        
        End If
            
    Next y
    Next x
    
    ThresholdPlusDither_ByteArray = True
    
End Function

'Given a byte array, reduce the number of available values to some number specified by the user (e.g. "2 shades" = monochrome,
' "4 shades" equals black, dark gray, light gray, white), with dithering.  (Currently dithering is limited to Floyd-Steinberg
' Floyd-Steinberg, but nothing prevents the use of other kernels).
Public Function Dither_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal numOfShades As Long = 4) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = 255 / (numOfShades - 1)
    
    'Build a look-up table for our custom conversion
    Dim g As Long, newG As Long
    'Dim grayLookUp() As Byte
    Dim grayLookUp(0 To 255) As Byte
    
    For x = 0 To 255
        g = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If (g > 255) Then g = 255
        grayLookUp(x) = g
    Next x
    
    'Prep a dither table.  Note that any dithering table will work, but at present
    ' I've hard-coded this function against Floyd-Steinberg dithering.
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    
    'Dim ditherTable() As Byte
    Dim ditherTable(-1 To 1, 0 To 1) As Byte
            
    ditherTable(1, 0) = 7
    ditherTable(-1, 1) = 3
    ditherTable(0, 1) = 5
    ditherTable(1, 1) = 1
    
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
    Dim xStride As Long, quickY As Long
    
    'Now loop through the array, calculating errors as we go
    For x = initX To finalX
    For y = initY To finalY
        
        g = srcArray(x, y)
        
        'Add in the current running error for this pixel
        g = g + dErrors(x, y)
        
        'Convert to a lookup-table safe value
        If (g >= 255) Then
            newG = 255
        ElseIf (g < 0) Then
            newG = 0
        Else
            newG = g
        End If
        
        'Write out the new luminance
        srcArray(x, y) = grayLookUp(newG)
        
        'Calculate an error
        errorVal = g - grayLookUp(newG)
        
        'If there is an error, spread it according to the dither table formula
        If (errorVal <> 0) Then
        
            For i = xLeft To xRight
            For j = 0 To yDown
            
                'First, ignore already processed pixels
                If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                
                'Second, ignore pixels that have a zero in the dither table
                If ditherTable(i, j) = 0 Then GoTo NextDitheredPixel
                
                xStride = x + i
                quickY = y + j
                
                'Next, ignore target pixels that are off the image boundary
                If (xStride < initX) Then GoTo NextDitheredPixel
                If (xStride > finalX) Then GoTo NextDitheredPixel
                If (quickY > finalY) Then GoTo NextDitheredPixel
                
                'If we've made it all the way here, we are able to actually spread the error to this location
                dErrors(xStride, quickY) = dErrors(xStride, quickY) + (errorVal * (CSng(ditherTable(i, j)) / dDivisor))
            
NextDitheredPixel:
            Next j
            Next i
        
        End If
            
    Next y
    Next x
    
    Dither_ByteArray = True
    
End Function

'Contrast-correct a byte array.  (This function is based off PD's white balance algorithm.)
Public Function ContrastCorrect_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal percentIgnore As Double) As Boolean
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long
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
    Dim lCount(0 To 255) As Long
    
    'Build an initial histogram
    For y = 0 To finalY
    For x = 0 To finalX
    
        'Increment the histogram at this position
        g = srcArray(x, y)
        lCount(g) = lCount(g) + 1
        
    Next x
    Next y
    
     'With the histogram complete, we can now figure out how to stretch the gray map. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    Dim foundYet As Boolean
    foundYet = False
    
    Dim numOfPixels As Long
    numOfPixels = arrayWidth * arrayHeight
    
    Dim wbThreshold As Long
    wbThreshold = numOfPixels * percentIgnore
    
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
    For y = 0 To finalY
    For x = 0 To finalX
        srcArray(x, y) = lFinal(srcArray(x, y))
    Next x
    Next y
    
    ContrastCorrect_ByteArray = True
    
End Function

'Find the range-based median of each entry in a given byte array.  pdPixelIterator is used.
Public Function Median_ByteArray(ByVal mRadius As Long, ByVal mPercent As Double, ByVal kernelShape As PD_PixelRegionShape, ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If (mRadius > (finalY - initY)) Then mRadius = finalY - initY
    Else
        If (mRadius > (finalX - initX)) Then mRadius = finalX - initX
    End If
    
    If (mRadius < 1) Then mRadius = 1
        
    mPercent = mPercent / 100
    If (mPercent < 0.01) Then mPercent = 0.01
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalX Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim lValues(0 To 255) As Long
    
    Dim cutoffTotal As Long
    Dim l As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator_ByteArray(srcArray, arrayWidth, arrayHeight, mRadius, mRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_ByteArray(lValues)
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = initX To finalX
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
            
                'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
                ' actual median value.
                
                'Loop through each color component histogram, until we've passed the desired percentile of pixels
                l = 0
                cutoffTotal = (mPercent * numOfPixels)
                If (cutoffTotal = 0) Then cutoffTotal = 1
        
                i = -1
                Do
                    i = i + 1
                    l = l + lValues(i)
                Loop Until (l >= cutoffTotal)
                l = i
                
                'Finally, apply the results to the destination array.
                dstArray(x, y) = l
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown_Byte
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp_Byte
                End If
        
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight_Byte
            
            'Update the progress bar every (progBarCheck) lines
            If (Not suppressMessages) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
            
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_ByteArray lValues
          
        Median_ByteArray = (Not g_cancelCurrentAction)
    
    Else
        Median_ByteArray = True
    End If
    
End Function

'Find the range-based maximum value of each segment of a given byte array.  pdPixelIterator is used.
Public Function Dilate_ByteArray(ByVal mRadius As Long, ByVal kernelShape As PD_PixelRegionShape, ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If (mRadius > (finalY - initY)) Then mRadius = finalY - initY
    Else
        If (mRadius > (finalX - initX)) Then mRadius = finalX - initX
    End If
    
    If (mRadius < 1) Then mRadius = 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim lValues(0 To 255) As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator_ByteArray(srcArray, arrayWidth, arrayHeight, mRadius, mRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_ByteArray(lValues)
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = initX To finalX
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
            
                'With a local histogram successfully built for the area surrounding this pixel,
                ' we can now rapidly find the local maximum
                
                'Loop through histogram entries until we reach a non-zero value (meaning there is
                ' at least one pixel in this region with that value)
                i = 255
                Do While (lValues(i) = 0)
                    i = i - 1
                Loop
                
                'Finally, apply the results to the destination array.
                dstArray(x, y) = i
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown_Byte
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp_Byte
                End If
        
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight_Byte
            
            'Update the progress bar every (progBarCheck) lines
            If (Not suppressMessages) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
            
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_ByteArray lValues
          
        Dilate_ByteArray = (Not g_cancelCurrentAction)
    
    Else
        Dilate_ByteArray = True
    End If
    
End Function

'Find the minimum value of each block of a given byte array.  pdPixelIterator is used.
Public Function Erode_ByteArray(ByVal mRadius As Long, ByVal kernelShape As PD_PixelRegionShape, ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If (mRadius > (finalY - initY)) Then mRadius = finalY - initY
    Else
        If (mRadius > (finalX - initX)) Then mRadius = finalX - initX
    End If
    
    If (mRadius < 1) Then mRadius = 1
    
    'To keep processing quick, only update the progress bar periodically.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'The number of pixels in the current box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'We use an optimized histogram technique for tracking pixel values
    Dim lValues(0 To 255) As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator_ByteArray(srcArray, arrayWidth, arrayHeight, mRadius, mRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_ByteArray(lValues)
        
        'Loop through each pixel in the image, updating our sliding window box as we go
        For x = initX To finalX
            
            'Based on the direction we're traveling, reverse the interior loop boundaries at edges
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above,
            ' but in a vertical direction
            For y = startY To stopY Step yStep
            
                'With a local histogram successfully built for the area surrounding this pixel,
                ' we can now rapidly find the local minimum
                
                'Loop through histogram entries until we reach a non-zero value (meaning there is
                ' at least one pixel in this region with that value)
                i = 0
                Do While (lValues(i) = 0)
                    i = i + 1
                Loop
                
                'Finally, apply the results to the destination array.
                dstArray(x, y) = i
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown_Byte
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp_Byte
                End If
        
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight_Byte
            
            'Update the progress bar every (progBarCheck) lines
            If (Not suppressMessages) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
            
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_ByteArray lValues
          
        Erode_ByteArray = (Not g_cancelCurrentAction)
    
    Else
        Erode_ByteArray = True
    End If
    
End Function

'Given a byte array, invert all values (e.g. Value = (255 - Value)).
Public Function InvertByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    'Invert the array
    For y = 0 To finalY
    For x = 0 To finalX
        srcArray(x, y) = 255 - srcArray(x, y)
    Next x
    Next y
    
    InvertByteArray = True
    
End Function

