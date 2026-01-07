VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFilterLUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Filter LUT (Look Up Table) Support Class
'Copyright 2014-2026 by Tanner Helland
'Created: 15/January/15
'Last updated: 26/March/20
'Last update: performance improvements
'
'Per its name, this class provides support routines for LUT-based image filters.  The goal is to make LUT-based filters
' easy to create and chain together, without needing to know the details of how the various filters are implemented.
'
'The class is presently designed to simplify three tasks:
' 1) Generating LUTs for you, based on various input parameters.  These functions are generally color channel agnostic,
'     meaning you simply provide input parameters and a blank byte array, and the LUT generator will fill the byte array
'     for you.  What you do with that array - including what channels you apply it to - is up to you.
' 2) Chaining together existing LUTs.  Say you want to apply both brightness and contrast to an image, in one fell swoop.
'     The MergeLUT function takes two source LUTs and a destination LUT, and passes the second table through the first,
'     resulting in a single LUT that represents both modifications.
'     *IMPORTANT NOTE*  Order is important when generating LUTs, as most LUT transfers are unlikely to be associative.
'     Take care when using this function.
' 3) Applying LUTs to an image.
'
'Though this class is initially focused on producing LUTs from Curve data, my goal is to move many other LUT-based filters
' (levels, gamma, brightness/contrast, etc) to this class.  This would not only simplify a lot of PD's dialogs, but would
' also make it much easier to add LUT-based remapping to future filters, without needing to add a bunch of custom code.
'
'Another future goal for this class - on Vista+, anyway - is to try and wrap the "ApplyToDIB" class function around the GDI+
' ColorLUT effects interface.  (See http://msdn.microsoft.com/en-us/library/ms534061%28v=vs.85%29.aspx for details.)
' On later Windows versions, GDI+ is supposedly SIMD accelerated, which would make it capable of applying LUT-based filters
' much faster than we can natively achieve in VB.  Unfortunately, the GDI+ effects flat APIs are poorly documented, so it
' could be a headache to reverse-engineer this.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with its corresponding default LUT value.
Friend Sub FillLUT_Default(ByRef srcArray() As Byte)
    ReDim srcArray(0 To 255) As Byte
    Dim i As Long
    For i = 0 To 255
        srcArray(i) = i
    Next i
End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding brightness-adjusted LUT value
' (standard additive calculation).  The passed adjustment value should be on the range [-255, 255].
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Friend Sub FillLUT_Brightness(ByRef srcArray() As Byte, Optional ByVal brightnessAdjustment As Long = 0)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (brightnessAdjustment < -255) Or (brightnessAdjustment > 255) Then PDDebug.LogAction "WARNING: Invalid adjustment passed to pdFilterLut.FillLUT_Brightness().  Fix it!"
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        newValue = i + brightnessAdjustment
        
        'Clamp to byte range, as necessary
        If (newValue > 255) Then newValue = 255
        If (newValue < 0) Then newValue = 0
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding brightness-adjusted LUT value,
' using non-standard multiplicative scaling (e.g. New_Value = Old_Value * Input_Parameter).  The passed adjustment value should be
' on the range [0, 2 or 3], but the upper bound isn't actually checked - just don't make it enormous.
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Friend Sub FillLUT_BrightnessMultiplicative(ByRef srcArray() As Byte, Optional ByVal brightnessScale As Double = 1#)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (brightnessScale < 0) Then Exit Sub
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        newValue = i * brightnessScale
        
        'Clamp to byte range, as necessary
        If (newValue > 255) Then newValue = 255
        If (newValue < 0) Then newValue = 0
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding contrast-adjusted LUT value.
' The passed adjustment value should be on the range [-100, 100].
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Friend Sub FillLUT_Contrast(ByRef srcArray() As Byte, Optional ByVal contrastAdjustment As Long = 0)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (contrastAdjustment < -100) Or (contrastAdjustment > 100) Then Exit Sub
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        'Calculate contrast
        newValue = i + (((i - 127) * contrastAdjustment) \ 100)
        
        'Clamp to byte range, as necessary
        If (newValue > 255) Then newValue = 255
        If (newValue < 0) Then newValue = 0
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding gamma-adjusted LUT value.
' The passed adjustment value should be on the range [>0, ~4 or 5], but the upper bound isn't actually checked - just don't
' make it enormous, since it's used as part of an exponent function.  (Similarly, don't make it 0.)
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Friend Sub FillLUT_Gamma(ByRef srcArray() As Byte, Optional ByVal gammaAdjustment As Double = 1#)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (gammaAdjustment <= 0) Then Exit Sub
    
    Dim i As Long, newValue As Long, tmpGammaCalc As Double
        
    'Generate the LUT
    For i = 0 To 255
        
        'Calculate gamma
        tmpGammaCalc = i / 255#
        tmpGammaCalc = tmpGammaCalc ^ (1# / gammaAdjustment)
        tmpGammaCalc = tmpGammaCalc * 255#
        
        newValue = Int(tmpGammaCalc)
        
        'Clamp to byte range, as necessary
        If (newValue > 255) Then newValue = 255
        If (newValue < 0) Then newValue = 0
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array and an array of curve points, generate a Curve LUT using the same formula as PD's Curve tool.
'
' IMPORTANT DETAILS:
' 1) The input curvePoints() array is assumed to be 0-based, dimmed to its exact size, with at least two points present,
'     and all points on the range 0-255, presorted by their x-coordinate.  (While I could rig this function to sort the
'     input points for the caller, that's a hassle and a lot of extra code, so please - sort your points in advance.)
'
' 2) This function will temporarily change the upper bound of the curvePoints array.  This is required to prevent OOB issues
'     when calculating values near the start and end points in the curve.   This function will restore the original bound
'     before exiting, but I mention this because it will cause any cached unsafe array pointers (e.g. VarPtr) to be invalid.
'
' 3) srcArray doesn't have to be initialized.  This function will do it for you.
Friend Sub FillLUT_Curve(ByRef srcArray() As Byte, ByRef curvePoints() As PointFloat)

    'Start by setting up default parameters for the source array
    FillLUT_Default srcArray
    
    'Next, prepare some basic analysis values for the curve function
    Dim cResults() As Double
    ReDim cResults(-1 To 256) As Double
    
    Dim numOfPoints As Long
    numOfPoints = UBound(curvePoints) + 1
    
    'To prevent errors in the curve generation function, we will actually pad the curve point array with some dummy entries.
    ' These entries will be removed at the end of the function.
    ReDim Preserve curvePoints(0 To numOfPoints + 1) As PointFloat
    
    Dim i As Long
    Dim p() As Double, u() As Double
    ReDim p(0 To numOfPoints) As Double
    ReDim u(0 To numOfPoints) As Double
        
    'Initialize the curve generator
    SetPandU numOfPoints, curvePoints, p, u
    
    'Iterate through the point array, calculating splines as we go
    Dim xPos As Long, yPos As Single
    Dim minX As Long, maxX As Long
    minX = 256
    maxX = -1
    
    For i = 0 To numOfPoints - 1
        For xPos = curvePoints(i).x To curvePoints(i + 1).x
            yPos = GetCurvePoint(i, xPos, curvePoints, p, u)
            If (xPos < minX) Then minX = xPos
            If (xPos > maxX) Then maxX = xPos
            If (yPos > 255) Then yPos = 255
            If (yPos < 0) Then yPos = 0
            cResults(xPos) = yPos
        Next xPos
    Next i
        
    'cResults() now contains the y-coordinate of the spline for every x-coordinate that falls between the initial
    ' point and the final point.  Points outside this range are treated as flat lines with values matching the nearest
    ' end point, and we fill those values now.
    For i = -1 To minX - 1
        cResults(i) = cResults(minX)
    Next i
    
    For i = 256 To maxX + 1 Step -1
        cResults(i) = cResults(maxX)
    Next i
    
    'cResults is now complete.  Its primary dimension can be treated as [0, 255], and each entry in the array
    ' contains the y-value of the spline at that x-position.  This can be used to easily render the spline on-screen,
    ' or in our case, treat the curve data as remap instructions for pixel data.
    For i = 0 To 255
        srcArray(i) = Int(cResults(i))
    Next i
    
    'Restore the original array bounds before exiting
    ReDim Preserve curvePoints(0 To numOfPoints - 1) As PointFloat
    
End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding inverted LUT value.
Friend Sub FillLUT_Invert(ByRef srcArray() As Byte)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Generate the LUT
    Dim i As Long
    For i = 0 To 255
        srcArray(i) = 255 - i
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding range-adjusted LUT value.
' This LUT function is similar to Levels, but without the pesky midtone operator.  Note that both input and output ranges
' can be specified, but either will default to [0, 255] if unsupplied.
'
'The Optional keyword is only used to demonstrate default values (e.g. the values that return the Identity LUT).
' You should pass actual values if you want this function to do anything useful!
Friend Sub FillLUT_RemappedRange(ByRef srcArray() As Byte, Optional ByVal inputMin As Long = 0, Optional ByVal inputMax As Long = 255, Optional ByVal outputMin As Long = 0, Optional ByVal outputMax As Long = 255)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply bounds-checking
    Dim boundFail As Boolean
    boundFail = False
    
    If (inputMin < 0) Or (inputMax < 0) Or (outputMin < 0) Or (outputMax < 0) Then boundFail = True
    If (inputMin > 255) Or (inputMax > 255) Or (outputMin > 255) Or (outputMax > 255) Then boundFail = True
    If (inputMin >= inputMax) Or (outputMin >= outputMax) Then boundFail = True
    
    If boundFail Then
        Debug.Print "WARNING: Invalid adjustment passed to fillLUT_RemappedRange().  Fix it!"
        Exit Sub
    End If
    
    Dim i As Long, tmpCalc As Double
    Dim inRange As Long, outRange As Long
    
    inRange = inputMax - inputMin
    outRange = outputMax - outputMin
        
    'Generate the LUT
    For i = 0 To 255
        
        'Remap the input value to the [0, 1] range
        tmpCalc = (i - inputMin) / inRange
        
        'Now, map the [0, 1] value to the desired output range
        tmpCalc = tmpCalc * outRange + outputMin
        
        'Failsafe range clamping
        If (tmpCalc > 255) Then tmpCalc = 255
        If (tmpCalc < 0) Then tmpCalc = 0
        srcArray(i) = tmpCalc
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding threshold-adjusted LUT value.
' Values LESS THAN the cutoff will be set to 0.  Values GREATER THAN OR EQUAL TO the cutoff will be set to 255.
' The passed adjustment value should be on the range [0, 256], but this function will actually accept any input.
'
'The Optional keyword is only used to demonstrate the default value (which in this case, does NOT return an Identity LUT).
Friend Sub FillLUT_Threshold(ByRef srcArray() As Byte, Optional ByVal thresholdCutoff As Long = 127)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Note that this function does not require bounds-checking, but if it did, we would apply it here.
    Dim i As Long
    
    'Generate the LUT
    For i = 0 To 255
        If (i < thresholdCutoff) Then
            srcArray(i) = 0
        Else
            srcArray(i) = 255
        End If
    Next i

End Sub

'Spline initialization function used by the Curve LUT generator.
' Call this function once prior to generating a curve.
Private Sub SetPandU(ByVal numOfPoints As Long, ByRef srcPoints() As PointFloat, ByRef p() As Double, ByRef u() As Double)
    
    Dim i As Long
    Dim d() As Double, w() As Double
    ReDim d(0 To numOfPoints) As Double
    ReDim w(0 To numOfPoints) As Double
    
    'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
    'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
    'at knots they share.  General algorithm taken from a book which has long since been lost.
    
    'The math that derived this stuff is pretty messy...  expressions are isolated and put into
    'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
    'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
    'later, we use the p's and u's to calculate curve points...
    
    For i = 2 To numOfPoints - 1
        d(i) = 2 * (srcPoints(i + 1).x - srcPoints(i - 1).x)
    Next
    
    For i = 0 To numOfPoints - 1
        u(i) = srcPoints(i + 1).x - srcPoints(i).x
    Next
    
    For i = 2 To numOfPoints - 1
        w(i) = 6# * ((srcPoints(i + 1).y - srcPoints(i).y) / u(i) - (srcPoints(i).y - srcPoints(i - 1).y) / u(i - 1))
    Next
    
    For i = 2 To numOfPoints - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    
    p(1) = 0#
    For i = numOfPoints - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    
    p(numOfPoints) = 0#
            
End Sub

'Spline retrieval functions
Private Function GetCurvePoint(ByVal i As Long, ByVal v As Double, ByRef srcPoints() As PointFloat, ByRef p() As Double, ByRef u() As Double) As Double
    Dim t As Double
    t = (v - srcPoints(i).x) / u(i)
    GetCurvePoint = t * srcPoints(i + 1).y + (1 - t) * srcPoints(i).y + u(i) * u(i) * (f(t) * p(i + 1) + f(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function f(ByVal x As Double) As Double
    f = x * x * x - x
End Function

'Helper function for generating curves to be used as LUT input parameters.  Given a param array of points (in x/y order), this function
' will return an assembled POINTFLOAT array (as required by the fillLUT_Curve function, above).
'
'Sample usage would be something like this, for a dramatic S-curve:
' Dim curvePoints() As POINTFLOAT
' pdFilterLUTInstance.helper_QuickCreateCurveArray curvePoints, 0, 0, 96, 72, 192, 160, 255, 255
Friend Sub Helper_QuickCreateCurveArray(ByRef curvePoints() As PointFloat, ParamArray listOfPoints() As Variant)

    If UBound(listOfPoints) >= LBound(listOfPoints) Then
                    
        Dim i As Long, numOfPoints As Long
        
        numOfPoints = (UBound(listOfPoints) - LBound(listOfPoints) + 1) \ 2
        ReDim curvePoints(0 To numOfPoints - 1) As PointFloat
        
        For i = 0 To numOfPoints - 1
            curvePoints(i).x = listOfPoints(i * 2)
            curvePoints(i).y = listOfPoints(i * 2 + 1)
        Next i
        
    Else
        Debug.Print "No points provided - helper_QuickCreateCurveArray function failed!"
    End If

End Sub

'Given two source LUTs and a destination LUT, merge the two sources and place the composite result
' into the destination.
'
'IMPORTANT NOTE: order is crucial when using this function!  Consider the standard case of applying
' different LUTs to each individual RGB channel, then applying a second, luminance-based LUT to all
' channels equally.  In a situation like this, you'd want to specify the channel-specific LUTs as
' the FIRST LUT, then the uniform luminance LUT as the SECOND LUT.  (The primary concern is always
' clamping, so in general terms, try to save the "most prone to cause clamping" LUT as the final
' LUT in the merge.)
'
' ANOTHER IMPORTANT NOTE: the destination LUT must NEVER be one of the source LUTs.  Otherwise, you
' risk source entries being overwritten early in the copy phase, which can result in screwed up
' subsequent values (because they get remapped according to look-up values that have already been
' remapped!).
'
'Finally, note that the destination LUT will be automatically sized for you, but the source LUTs
' must (obviously) be already sized and filled appropriately.
Friend Sub MergeLUTs(ByRef firstSourceLUT() As Byte, ByRef secondSourceLUT() As Byte, ByRef dstLUT() As Byte)
    
    ReDim dstLUT(0 To 255) As Byte
    
    'Fill the destination LUT with the values of firstSourceLUT, as passed through secondSourceLUT
    Dim i As Long
    For i = 0 To 255
        dstLUT(i) = secondSourceLUT(firstSourceLUT(i))
    Next i

End Sub

'Given one LUT, apply it to each pixel in the image, using the pixel's calculated GRAY VALUE as the LUT lookup parameter.
'
' RETURNS: 1 if successful, 0 if unsuccessful or canceled by the user.
Friend Function ApplyLUTsToDIB_Gray(ByRef srcDIB As pdDIB, ByRef gLUT() As Byte) As Long
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.GetDIBColorDepth = 32) And srcDIB.GetAlphaPremultiplication Then
        srcWasPremultiplied = True
        srcDIB.SetAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long, grayVal As Long, grayByte As Byte
        
    'Now we can loop through each pixel in the image, converting values as we go
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (218 * r + 732 * g + 74 * b) \ 1024
        grayByte = gLUT(grayVal)
        
        'Assign the look-up table values
        imageData(x) = grayByte
        imageData(x + 1) = grayByte
        imageData(x + 2) = grayByte
        
    Next x
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.SetAlphaPremultiplication True
    
    If g_cancelCurrentAction Then ApplyLUTsToDIB_Gray = 0 Else ApplyLUTsToDIB_Gray = 1
    
End Function

'Apply a single LUT equally to each color channel.  If alpha is present, it is ignored.
'
' TODO: try wrapping GDI+ on Vista and later, and do some performance profiling to compare results.
'
' RETURNS: 1 if successful, 0 if unsuccessful or canceled by the user.
Friend Function ApplyLUTToAllColorChannels(ByRef srcDIB As pdDIB, ByRef cLUT() As Byte) As Long
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.GetDIBColorDepth = 32) And srcDIB.GetAlphaPremultiplication Then
        srcWasPremultiplied = True
        srcDIB.SetAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0 * 4
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Assign the look-up table values
        imageData(x) = cLUT(b)
        imageData(x + 1) = cLUT(g)
        imageData(x + 2) = cLUT(r)
        
    Next x
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.SetAlphaPremultiplication True
    
    If g_cancelCurrentAction Then ApplyLUTToAllColorChannels = 0 Else ApplyLUTToAllColorChannels = 1
    
End Function
