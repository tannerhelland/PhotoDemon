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
'Copyright 2014-2015 by Tanner Helland
'Created: 15/January/15
'Last updated: 16/January/15
'Last update: continue fleshing out features
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with its corresponding default LUT value.
Public Sub fillLUT_Default(ByRef srcArray() As Byte)

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
Public Sub fillLUT_Brightness(ByRef srcArray() As Byte, Optional ByVal brightnessAdjustment As Long = 0)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (brightnessAdjustment < -255) Or (brightnessAdjustment > 255) Then
        Debug.Print "WARNING: Invalid adjustment passed to fillLUT_Brightness().  Fix it!"
        Exit Sub
    End If
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        newValue = i + brightnessAdjustment
        
        'Clamp to byte range, as necessary
        If newValue > 255 Then
            newValue = 255
        ElseIf newValue < 0 Then
            newValue = 0
        End If
        
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding brightness-adjusted LUT value,
' using non-standard multiplicative scaling (e.g. New_Value = Old_Value * Input_Parameter).  The passed adjustment value should be
' on the range [0, 2 or 3], but the upper bound isn't actually checked - just don't make it enormous.
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Public Sub fillLUT_BrightnessMultiplicative(ByRef srcArray() As Byte, Optional ByVal brightnessScale As Double = 1#)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (brightnessScale < 0) Then
        Debug.Print "WARNING: Invalid adjustment passed to fillLUT_BrightnessMultiplicative().  Fix it!"
        Exit Sub
    End If
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        newValue = i * brightnessScale
        
        'Clamp to byte range, as necessary
        If newValue > 255 Then
            newValue = 255
        ElseIf newValue < 0 Then
            newValue = 0
        End If
        
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding contrast-adjusted LUT value.
' The passed adjustment value should be on the range [-100, 100].
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Public Sub fillLUT_Contrast(ByRef srcArray() As Byte, Optional ByVal contrastAdjustment As Long = 0)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (contrastAdjustment < -100) Or (contrastAdjustment > 100) Then
        Debug.Print "WARNING: Invalid adjustment passed to fillLUT_Contrast().  Fix it!"
        Exit Sub
    End If
    
    Dim i As Long, newValue As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        'Calculate contrast
        newValue = i + (((i - 127) * contrastAdjustment) \ 100)
        
        'Clamp to byte range, as necessary
        If newValue > 255 Then
            newValue = 255
        ElseIf newValue < 0 Then
            newValue = 0
        End If
        
        srcArray(i) = newValue
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding gamma-adjusted LUT value.
' The passed adjustment value should be on the range [>0, ~4 or 5], but the upper bound isn't actually checked - just don't
' make it enormous, since it's used as part of an exponent function.  (Similarly, don't make it 0.)
'
'The Optional keyword is only used to demonstrate the default value (e.g. the value that returns the Identity LUT).
' You should pass an actual value if you want this function to do anything useful!
Public Sub fillLUT_Gamma(ByRef srcArray() As Byte, Optional ByVal gammaAdjustment As Double = 1#)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Apply basic bounds-checking to the input parameter
    If (gammaAdjustment <= 0) Then
        Debug.Print "WARNING: Invalid adjustment passed to fillLUT_Gamma().  Fix it!"
        Exit Sub
    End If
    
    Dim i As Long, newValue As Long, tmpGammaCalc As Double
        
    'Generate the LUT
    For i = 0 To 255
        
        'Calculate gamma
        tmpGammaCalc = i / 255
        tmpGammaCalc = tmpGammaCalc ^ (1 / gammaAdjustment)
        tmpGammaCalc = tmpGammaCalc * 255
        
        newValue = Int(tmpGammaCalc)
        
        'Clamp to byte range, as necessary
        If newValue > 255 Then
            newValue = 255
        ElseIf newValue < 0 Then
            newValue = 0
        End If
        
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
Friend Sub fillLUT_Curve(ByRef srcArray() As Byte, ByRef curvePoints() As POINTFLOAT)

    'Start by setting up default parameters for the source array
    fillLUT_Default srcArray
    
    'Next, prepare some basic analysis values for the curve function
    Dim cResults() As Double
    ReDim cResults(-1 To 256) As Double
    
    Dim numOfPoints As Long
    numOfPoints = UBound(curvePoints) + 1
    
    'To prevent errors in the curve generation function, we will actually pad the curve point array with some dummy entries.
    ' These entries will be removed at the end of the function.
    ReDim Preserve curvePoints(0 To numOfPoints + 1) As POINTFLOAT
    
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
            yPos = getCurvePoint(i, xPos, curvePoints, p, u)
            If xPos < minX Then minX = xPos
            If xPos > maxX Then maxX = xPos
            If yPos > 255 Then yPos = 255
            If yPos < 0 Then yPos = 0
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
    ReDim Preserve curvePoints(0 To numOfPoints - 1) As POINTFLOAT
    
End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding inverted LUT value.
Public Sub fillLUT_Invert(ByRef srcArray() As Byte)
    
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
Public Sub fillLUT_RemappedRange(ByRef srcArray() As Byte, Optional ByVal inputMin As Long = 0, Optional ByVal inputMax As Long = 255, Optional ByVal outputMin As Long = 0, Optional ByVal outputMax As Long = 255)
    
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
        If tmpCalc < 0 Then
            tmpCalc = 0
        ElseIf tmpCalc > 255 Then
            tmpCalc = 255
        End If
        
        srcArray(i) = tmpCalc
        
    Next i

End Sub

'Given an arbitrary byte array, ReDim it to [0, 255] and fill each index with a corresponding threshold-adjusted LUT value.
' Values LESS THAN the cutoff will be set to 0.  Values GREATER THAN OR EQUAL TO the cutoff will be set to 255.
' The passed adjustment value should be on the range [0, 256], but this function will actually accept any input.
'
'The Optional keyword is only used to demonstrate the default value (which in this case, does NOT return an Identity LUT).
Public Sub fillLUT_Threshold(ByRef srcArray() As Byte, Optional ByVal thresholdCutoff As Long = 127)
    
    'Prep the source array
    ReDim srcArray(0 To 255) As Byte
    
    'Note that this function does not require bounds-checking, but if it did, we would apply it here.
    
    Dim i As Long
    
    'Generate the LUT
    For i = 0 To 255
        
        If i < thresholdCutoff Then
            srcArray(i) = 0
        Else
            srcArray(i) = 255
        End If
        
    Next i

End Sub

'Spline initialization function used by the Curve LUT generator.  Call this function once prior to generating a curve.
Private Sub SetPandU(ByVal numOfPoints As Long, ByRef srcPoints() As POINTFLOAT, ByRef p() As Double, ByRef u() As Double)
    
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
Private Function getCurvePoint(ByVal i As Long, ByVal v As Double, ByRef srcPoints() As POINTFLOAT, ByRef p() As Double, ByRef u() As Double) As Double
    Dim t As Double
    t = (v - srcPoints(i).x) / u(i)
    getCurvePoint = t * srcPoints(i + 1).y + (1 - t) * srcPoints(i).y + u(i) * u(i) * (f(t) * p(i + 1) + f(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function f(ByRef x As Double) As Double
        f = x * x * x - x
End Function

'Helper function for generating curves to be used as LUT input parameters.  Given a param array of points (in x/y order), this function
' will return an assembled POINTFLOAT array (as required by the fillLUT_Curve function, above).
'
'Sample usage would be something like this, for a dramatic S-curve:
' Dim curvePoints() As POINTFLOAT
' pdFilterLUTInstance.helper_QuickCreateCurveArray curvePoints, 0, 0, 96, 72, 192, 160, 255, 255
Friend Sub helper_QuickCreateCurveArray(ByRef curvePoints() As POINTFLOAT, ParamArray listOfPoints() As Variant)

    If UBound(listOfPoints) >= LBound(listOfPoints) Then
                    
        Dim i As Long, numOfPoints As Long
        
        numOfPoints = (UBound(listOfPoints) - LBound(listOfPoints) + 1) \ 2
        ReDim curvePoints(0 To numOfPoints - 1) As POINTFLOAT
        
        For i = 0 To numOfPoints - 1
            curvePoints(i).x = listOfPoints(i * 2)
            curvePoints(i).y = listOfPoints(i * 2 + 1)
        Next i
        
    Else
        Debug.Print "No points provided - helper_QuickCreateCurveArray function failed!"
    End If

End Sub

'Given two source LUTs and a destination LUT, merge the two sources and place the composite result into the destination.
' IMPORTANT NOTE: order is crucial when using this function.  Consider the standard case of applying different LUTs to each individual
'                 RGB channel, then applying a second, luminance-based LUT to all channels equally.  In a situation like this, you'd
'                 want to specify the channel-specific LUTs as the FIRST LUT, then the uniform luminance LUT as the SECOND LUT.
'                 (The primary concern is always clamping, so in general terms, try to save the "most prone to cause clamping" LUT
'                  as the final LUT in the merge.)
'
' ANOTHER IMPORTANT NOTE: the destination LUT must NEVER be one of the source LUTs.  Otherwise, you risk source entries being
'                         overwritten early in the copy phase, which can result in screwed up subsequent values (because they get
'                         remapped according to look-up values that have already been remapped!).
'
'Finally, note that the destination LUT will be automatically sized for you, but the source LUTs must (obviously) be already sized
' and filled appropriately.
Public Sub MergeLUTs(ByRef firstSourceLUT() As Byte, ByRef secondSourceLUT() As Byte, ByRef dstLUT() As Byte)
    
    'Prep the source array
    ReDim dstLUT(0 To 255) As Byte
    
    Dim i As Long, newVal As Long
    
    'Fill the destination LUT with the values of firstSourceLUT, as passed through secondSourceLUT
    For i = 0 To 255
        dstLUT(i) = secondSourceLUT(firstSourceLUT(i))
    Next i

End Sub

'Given three LUTs - one for each color channel - apply them to a given DIB.
' IMPORTANT NOTE: alpha LUTs are not currently implemented.  What I might do is create a separate function that takes four LUT inputs.
'                  That would keep this function speedy, since alpha could be address on an as-needed basis.
'
' TODO: try wrapping GDI+ on Vista and later, and do some performance profiling to compare results.
'
' RETURNS: 1 if successful, 0 if unsuccessful or canceled by the user.
Public Function applyLUTsToDIB_Color(ByRef srcDIB As pdDIB, ByRef rLUT() As Byte, ByRef gLUT() As Byte, ByRef bLUT() As Byte, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.getDIBColorDepth = 32) And srcDIB.getAlphaPremultiplication Then
        srcWasPremultiplied = True
        srcDIB.setAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the look-up table values
        ImageData(QuickVal + 2, y) = rLUT(r)
        ImageData(QuickVal + 1, y) = gLUT(g)
        ImageData(QuickVal, y) = bLUT(b)
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.setAlphaPremultiplication True
    
    If cancelCurrentAction Then applyLUTsToDIB_Color = 0 Else applyLUTsToDIB_Color = 1
    
End Function

'Given one LUT, apply it to each pixel in the image, using the pixel's calculated GRAY VALUE as the LUT lookup parameter.
'
' RETURNS: 1 if successful, 0 if unsuccessful or canceled by the user.
Public Function applyLUTsToDIB_Gray(ByRef srcDIB As pdDIB, ByRef gLUT() As Byte, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.getDIBColorDepth = 32) And srcDIB.getAlphaPremultiplication Then
        srcWasPremultiplied = True
        srcDIB.setAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long, grayVal As Long, grayByte As Byte
        
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If grayVal > 255 Then grayVal = 255
        grayByte = gLUT(grayVal)
        
        'Assign the look-up table values
        ImageData(QuickVal + 2, y) = grayByte
        ImageData(QuickVal + 1, y) = grayByte
        ImageData(QuickVal, y) = grayByte
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.setAlphaPremultiplication True
    
    If cancelCurrentAction Then applyLUTsToDIB_Gray = 0 Else applyLUTsToDIB_Gray = 1
    
End Function

'Apply a single LUT equally to each color channel.  If alpha is present, it is ignored.
'
' TODO: try wrapping GDI+ on Vista and later, and do some performance profiling to compare results.
'
' RETURNS: 1 if successful, 0 if unsuccessful or canceled by the user.
Public Function applyLUTToAllColorChannels(ByRef srcDIB As pdDIB, ByRef cLUT() As Byte, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Unpremultiply the source DIB's alpha, as necessary
    Dim srcWasPremultiplied As Boolean
    If (srcDIB.getDIBColorDepth = 32) And srcDIB.getAlphaPremultiplication Then
        srcWasPremultiplied = True
        srcDIB.setAlphaPremultiplication False
    Else
        srcWasPremultiplied = False
    End If

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values are used, but I plan on doing some testing to see if direct LUT assignment is faster.
    Dim r As Long, g As Long, b As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the look-up table values
        ImageData(QuickVal + 2, y) = cLUT(r)
        ImageData(QuickVal + 1, y) = cLUT(g)
        ImageData(QuickVal, y) = cLUT(b)
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Reset alpha premultiplication, as necessary
    If srcWasPremultiplied Then srcDIB.setAlphaPremultiplication True
    
    If cancelCurrentAction Then applyLUTToAllColorChannels = 0 Else applyLUTToAllColorChannels = 1
    
End Function
