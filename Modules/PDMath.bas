Attribute VB_Name = "PDMath"
'***************************************************************************
'Specialized Math Routines
'Copyright 2013-2017 by Tanner Helland
'Created: 13/June/13
'Last updated: 12/January/17
'Last update: added two optimized Atan2() variants, each with trade-offs between accuracy and performance.
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June '13.  This module is now used to store all the random bits of specialized math required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Many drawing features lean on various geometry functions
Public Const PI As Double = 3.14159265358979
Public Const PI_HALF As Double = 1.5707963267949
Public Const PI_DOUBLE As Double = 6.28318530717958
Public Const PI_DIV_180 As Double = 0.017453292519943
Public Const PI_14 As Double = 0.785398163397448
Public Const PI_34 As Double = 2.35619449019234

Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRectL Lib "user32" Alias "PtInRect" (ByRef lpRect As RECTL, ByVal x As Long, ByVal y As Long) As Long

'See if a point lies inside a rect (integer)
Public Function IsPointInRect(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECT) As Boolean
    IsPointInRect = (PtInRect(srcRect, ptX, ptY) <> 0)
End Function

'See if a point lies inside a RectL struct
Public Function IsPointInRectL(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECTL) As Boolean
    IsPointInRectL = (PtInRectL(srcRect, ptX, ptY) <> 0)
End Function

'See if a point lies inside a rect (float)
Public Function IsPointInRectF(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECTF) As Boolean
    With srcRect
        If (ptX >= .Left) And (ptX <= (.Left + .Width)) Then
            IsPointInRectF = ((ptY >= .Top) And (ptY <= (.Top + .Height)))
        Else
            IsPointInRectF = False
        End If
    End With
End Function

'Find the union rect of two floating-point rects.  (This is the smallest rect that contains both rects.)
Public Sub UnionRectF(ByRef dstRect As RECTF, ByRef srcRect As RECTF, ByRef srcRect2 As RECTF, Optional ByVal widthAndHeightAreReallyRightAndBottom As Boolean = False)

    'Union rects are easy: find the min top/left, and the max bottom/right
    With dstRect
        
        If (srcRect.Left < srcRect2.Left) Then
            .Left = srcRect.Left
        Else
            .Left = srcRect2.Left
        End If
        
        If (srcRect.Top < srcRect2.Top) Then
            .Top = srcRect.Top
        Else
            .Top = srcRect2.Top
        End If
        
        'Next, determine right bounds.  Note that the caller can stuff right bounds into a floating-point rect, and this function will handle that
        ' case contingent on the (very long-named) widthAndHeightAreReallyRightAndBottom parameter.
        Dim srcRight As Single, srcRight2 As Single
        
        If widthAndHeightAreReallyRightAndBottom Then
            srcRight = srcRect.Width
            srcRight2 = srcRect2.Width
        Else
            srcRight = srcRect.Left + srcRect.Width
            srcRight2 = srcRect2.Left + srcRect2.Width
        End If
        
        'Find the max value and store it in srcRight
        If (srcRight < srcRight2) Then srcRight = srcRight2
        
        'Account for widthAndHeightAreReallyRightAndBottom (again)
        If widthAndHeightAreReallyRightAndBottom Then
            .Width = srcRight
        Else
            .Width = srcRight - .Left
        End If
        
        'Repeat the above steps for the bottom bound
        Dim srcBottom As Single, srcBottom2 As Single
        
        If widthAndHeightAreReallyRightAndBottom Then
            srcBottom = srcRect.Height
            srcBottom2 = srcRect2.Height
        Else
            srcBottom = srcRect.Top + srcRect.Height
            srcBottom2 = srcRect2.Top + srcRect2.Height
        End If
        
        If (srcBottom < srcBottom2) Then srcBottom = srcBottom2
        
        If widthAndHeightAreReallyRightAndBottom Then
            .Height = srcBottom
        Else
            .Height = srcBottom - .Top
        End If
        
    End With

End Sub

Public Function AreRectFsEqual(ByRef srcRectF1 As RECTF, ByRef srcRectf2 As RECTF) As Boolean
    AreRectFsEqual = VBHacks.MemCmp(VarPtr(srcRectF1), VarPtr(srcRectf2), LenB(srcRectF1))
End Function

Public Function Frac(ByVal srcValue As Double) As Double
    Frac = srcValue - Int(srcValue)
End Function

'Convert a decimal to a near-identical fraction using vector math.
' This excellent function comes courtesy of VB6 coder LaVolpe.  I have modified it slightly to suit PhotoDemon's unique needs.
' You can download the original at this link (good as of 13 June 2014): http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=61596&lngWId=1
Public Sub ConvertToFraction(ByVal v As Double, w As Double, n As Double, d As Double, Optional ByVal maxDenomDigits As Byte, Optional ByVal accuracy As Double = 100#)

    Const MAX_TERMS As Long = 50          'Limit to prevent infinite loop
    Const MIN_DIVISOR As Double = 1E-16      'Limit to prevent divide by zero
    Const MAX_ERROR As Double = 1E-50        'How close is enough
    Dim f As Double                         'Fraction being converted
    Dim a As Double     'Current term in continued fraction
    Dim n1 As Double    'Numerator, denominator of last approx
    Dim d1 As Double
    Dim n2 As Double    'Numerator, denominator of previous approx
    Dim d2 As Double
    Dim i As Integer
    Dim t As Double
    Dim maxDenom As Double
    Dim bIsNegative As Boolean
    
    If maxDenomDigits = 0 Or maxDenomDigits > 17 Then maxDenomDigits = 17
    maxDenom = 10 ^ maxDenomDigits
    If accuracy > 100 Or accuracy < 1 Then accuracy = 100
    accuracy = accuracy / 100#
    
    bIsNegative = (v < 0)
    w = Abs(Fix(v))
    'V = Abs(V) - W << subtracting doubles can change the decimal portion by adding more numeral at end
    'TANNER'S NOTE: the original version of this included +1 to the string length, which gave me consistent errors.  So I removed it.
    v = CDbl(Mid$(CStr(Abs(v)), Len(CStr(w))))
    
    ' check for no decimal or zero
    If v = 0 Then GoTo RtnResult
    
    f = v                       'Initialize fraction being converted
    
    n1 = 1                      'Initialize fractions with 1/0, 0/1
    d1 = 0
    n2 = 0
    d2 = 1

    On Error GoTo RtnResult
    For i = 0 To MAX_TERMS
        a = Fix(f)              'Get next term
        f = f - a               'Get new divisor
        n = n1 * a + n2         'Calculate new fraction
        d = d1 * a + d2
        n2 = n1                 'Save last two fractions
        d2 = d1
        n1 = n
        d1 = d
                                'Quit if dividing by zero
        If f < MIN_DIVISOR Then Exit For

                                'Quit if close enough
        t = n / d               ' A=zero indicates exact match or extremely close
        a = Abs(v - t)          ' Difference btwn actual V and calculated T
        If a < MAX_ERROR Then Exit For
                                'Quit if max denominator digits encountered
        If d > maxDenom Then Exit For
                                ' Quit if preferred accuracy accomplished
        If n Then
            If t > v Then t = v / t Else t = t / v
            If t >= accuracy And Abs(t) < 1 Then Exit For
        End If
        f = 1# / f               'Take reciprocal
    Next i

RtnResult:
    If Err Or d > maxDenom Then
        ' in above case, use the previous best N & D
        If d2 = 0 Then
            n = n1
            d = d1
        Else
            d = d2
            n = n2
        End If
    End If
    
    ' correct for negative values
    If bIsNegative Then
        If w Then w = -w Else n = -n
    End If
    
    'TANNER'S NOTE: the original function included some simple code here to generate a user-friendly string of the result.
    ' PhotoDemon does this itself (so translations can be applied) so I have removed that section of code.

End Sub

'Convert a width and height pair to a new max width and height, while preserving aspect ratio
' NOTE: by default, inclusive fitting is assumed, but the user can set that parameter to false.  That can be used to
'        fit an image into a new size with no blank space, but cropping overhanging edges as necessary.)
Public Sub ConvertAspectRatio(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long, Optional ByVal fitInclusive As Boolean = True)
    
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0) And (dstHeight > 0) Then
        srcAspect = srcWidth / srcHeight
        dstAspect = dstWidth / dstHeight
    Else
        Exit Sub
    End If
    
    Dim aspectLarger As Boolean
    aspectLarger = (srcAspect > dstAspect)
    
    'Exclusive fitting fits the opposite dimension, so simply reverse the way the dimensions are calculated
    If (Not fitInclusive) Then aspectLarger = Not aspectLarger
    
    If aspectLarger Then
        newWidth = dstWidth
        newHeight = CDbl(srcHeight / srcWidth) * newWidth
    Else
        newHeight = dstHeight
        newWidth = CDbl(srcWidth / srcHeight) * newHeight
    End If
    
End Sub

'Return the distance between two values on the same line
Public Function DistanceOneDimension(ByVal x1 As Double, ByVal x2 As Double) As Double
    DistanceOneDimension = Sqr((x1 - x2) * (x1 - x2))
End Function

'Return the distance between two points
Public Function DistanceTwoPoints(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    DistanceTwoPoints = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

'Return the distance between two points, but ignores the square root function; if calculating something simple, like "minimum distance only",
' we only need relative values - not absolute ones - so we can skip that step for an extra performance boost.
Public Function DistanceTwoPointsShortcut(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    DistanceTwoPointsShortcut = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)
End Function

'Return the distance between two 3D points
Public Function DistanceThreeDimensions(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
    DistanceThreeDimensions = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
End Function

Public Function Distance3D_FastFloat(ByRef x1 As Single, ByRef y1 As Single, ByRef z1 As Single, ByRef x2 As Single, ByRef y2 As Single, ByRef z2 As Single) As Single
    Distance3D_FastFloat = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
End Function

'Given two intersecting lines, return the angle between them (e.g. the inner product: https://en.wikipedia.org/wiki/Inner_product_space)
Public Function AngleBetweenTwoIntersectingLines(ByRef ptIntersect As POINTFLOAT, ByRef pt1 As POINTFLOAT, ByRef pt2 As POINTFLOAT, Optional ByVal returnResultInDegrees As Boolean = True) As Double
    
    Dim dx1i As Double, dy1i As Double, dx2i As Double, dy2i As Double
    dx1i = pt1.x - ptIntersect.x
    dy1i = pt1.y - ptIntersect.y
    
    dx2i = pt2.x - ptIntersect.x
    dy2i = pt2.y - ptIntersect.y
    
    Dim m12 As Double, m13 As Double
    m12 = Sqr(dx1i * dx1i + dy1i * dy1i)
    m13 = Sqr(dx2i * dx2i + dy2i * dy2i)
    
    AngleBetweenTwoIntersectingLines = Acos((dx1i * dx2i + dy1i * dy2i) / (m12 * m13))
    
    If returnResultInDegrees Then
        AngleBetweenTwoIntersectingLines = AngleBetweenTwoIntersectingLines / PI_DIV_180
    End If
    
End Function

'Fast arctangent estimation.  Max error 0.0015 radians (0.085944 degrees), first found here: http://nghiaho.com/?p=997
' IMPORTANT NOTE: only works for (x) values on the range [-1, 1]; as such, it should only be used with normalized values.
' Because many PD functions do not normalize prior to calling Atn(), I've commented this out for now to reduce confusion.
'Public Function Atn_Fast(ByVal x As Double) As Double
'    Atn_Fast = PI_14 * x - x * (Abs(x) - 1.0) * (0.2447 + 0.0663 * Abs(x))
'End Function

'Return the arctangent of two values (rise / run); unlike VB's integrated Atn() function, this return is quadrant-specific.
' (It also circumvents potential DBZ errors when horizontal.)
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
 
    If (y = 0#) And (x = 0#) Then
        Atan2 = 0#
        Exit Function
    End If
 
    If (y > 0#) Then
        If (x >= y) Then
            Atan2 = Atn(y / x)
        ElseIf (x <= -y) Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = PI_HALF - Atn(x / y)
        End If
    Else
        If (x >= -y) Then
            Atan2 = Atn(y / x)
        ElseIf (x <= y) Then
            Atan2 = Atn(y / x) - PI
        Else
            Atan2 = -Atn(x / y) - PI_HALF
        End If
    End If
 
End Function

'Estimation optimization of Atan2, using Hastings optimizations (https://lists.apple.com/archives/perfoptimization-dev/2005/Jan/msg00051.html)
' Stated absolute error is expected to be < 0.005, which is more than good enough for most PD tasks.
' This function is reliably faster than the "perfect" Atan2() function, above, and valid for all quadrants.
Public Function Atan2_Faster(ByVal y As Double, ByVal x As Double) As Double
    
    If (x = 0#) Then
       If (y > 0#) Then
           Atan2_Faster = PI_HALF
       ElseIf (y = 0#) Then
           Atan2_Faster = 0#
       Else
           Atan2_Faster = -PI_HALF
       End If
    Else
       Dim z As Double
       z = y / x
       If (Abs(z) < 1#) Then
           Atan2_Faster = z / (1# + 0.28 * z * z)
           If (x < 0#) Then
               If (y < 0#) Then
                   Atan2_Faster = Atan2_Faster - PI
               Else
                   Atan2_Faster = Atan2_Faster + PI
               End If
           End If
       Else
           Atan2_Faster = PI_HALF - z / (z * z + 0.28)
           If (y < 0#) Then Atan2_Faster = Atan2_Faster - PI
       End If
    End If
    
End Function

'Attempted estimation optimization of Atan2, using self-normalization (https://web.archive.org/web/20090519203600/http://www.dspguru.com:80/comp.dsp/tricks/alg/fxdatan2.htm)
' Stated worst-case error is expected to be < 0.07, which is good enough for certain PD tasks (e.g. image distort filters).
' This function is reliably faster than the "perfect" Atan2() function, above, as well as the Atan2_Faster() function,
' while remaining valid for all quadrants.
Public Function Atan2_Fastest(ByVal y As Double, ByVal x As Double) As Double
    
    'Cheap non-branching workaround for the case y = 0.0
    Dim absY As Double
    absY = Abs(y) + 0.0000000001
    
    If (x >= 0#) Then
        Atan2_Fastest = PI_14 - PI_14 * (x - absY) / (x + absY)
    Else
        Atan2_Fastest = PI_34 - PI_14 * (x + absY) / (absY - x)
    End If
    
    If (y < 0#) Then Atan2_Fastest = -Atan2_Fastest
    
End Function

'Arcsine function
Public Function Asin(ByVal x As Double) As Double
    If (x > 1#) Or (x < -1#) Then x = 1#
    Asin = Atan2(x, Sqr(1# - x * x))
End Function

'Arccosine function
Public Function Acos(ByVal x As Double) As Double
    If (x > 1#) Or (x < -1#) Then x = 1#
    Acos = Atan2(Sqr(1# - x * x), x)
End Function

'Given a list of floating-point values, convert each to its integer equivalent *furthest* from 0.
' Said another way, round negative numbers down, and positive numbers up.  This is often relevant in PD when performing
' coordinate conversions that are ultimately mapped to pixel locations, and we need to bounds-check corner coordinates
' in advance and push them away from 0, so any partially-covered pixels are converted to fully-covered ones.
Public Function ConvertArbitraryListToFurthestRoundedInt(ParamArray listOfValues() As Variant)
    
    If UBound(listOfValues) >= LBound(listOfValues) Then
        
        Dim i As Long
        For i = LBound(listOfValues) To UBound(listOfValues)
            If (listOfValues(i) < 0) Then
                listOfValues(i) = Int(listOfValues(i))
            Else
                If (listOfValues(i) = Int(listOfValues(i))) Then
                    listOfValues(i) = Int(listOfValues(i))
                Else
                    listOfValues(i) = Int(listOfValues(i)) + 1
                End If
            End If
        Next i
        
    Else
        Debug.Print "No points provided - convertArbitraryFListToRoundedInt() function failed!"
    End If

End Function

Public Sub ConvertCartesianToPolar(ByVal srcX As Double, ByVal srcY As Double, ByRef dstRadius As Double, ByRef dstAngle As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)
    dstRadius = Sqr((srcX - centerX) * (srcX - centerX) + (srcY - centerY) * (srcY - centerY))
    dstAngle = PDMath.Atan2_Faster((srcY - centerY), (srcX - centerX))
End Sub

Public Sub ConvertPolarToCartesian(ByVal srcAngle As Double, ByVal srcRadius As Double, ByRef dstX As Double, ByRef dstY As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)

    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub

Public Sub ConvertPolarToCartesian_Sng(ByVal srcAngle As Single, ByVal srcRadius As Single, ByRef dstX As Single, ByRef dstY As Single, Optional ByVal centerX As Single = 0#, Optional ByVal centerY As Single = 0#)

    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub
'Given an array of points, find the closest one to a target location.  If none fall below a minimum distance threshold, return -1.
' (This function is used by many bits of mouse interaction code, to see if the user has clicked on something interesting.)
Public Function FindClosestPointInArray(ByVal targetX As Double, ByVal targetY As Double, ByVal minAllowedDistance As Double, ByRef poiArray() As POINTAPI) As Long

    Dim curMinDistance As Double, curMinIndex As Long
    curMinDistance = &HFFFFFFF
    curMinIndex = -1
    
    Dim tmpDistance As Double
    
    'From the array of supplied points, find the one closest to the target point
    Dim i As Long
    For i = LBound(poiArray) To UBound(poiArray)
        tmpDistance = DistanceTwoPoints(targetX, targetY, poiArray(i).x, poiArray(i).y)
        If (tmpDistance < curMinDistance) Then
            curMinDistance = tmpDistance
            curMinIndex = i
        End If
    Next i
    
    'If the distance of the closest point falls below the allowed threshold, return that point's index.
    If (curMinDistance < minAllowedDistance) Then
        FindClosestPointInArray = curMinIndex
    Else
        FindClosestPointInArray = -1
    End If

End Function

'Given an array of points (in floating-point format), find the closest one to a target location.  If none fall below a minimum distance threshold,
' return -1.  (This function is used by many bits of mouse interaction code, to see if the user has clicked on something interesting.)
Public Function FindClosestPointInFloatArray(ByVal targetX As Single, ByVal targetY As Single, ByVal minAllowedDistance As Single, ByRef poiArray() As POINTFLOAT) As Long

    Dim curMinDistance As Double, curMinIndex As Long
    curMinDistance = &HFFFFFFF
    curMinIndex = -1
    
    Dim tmpDistance As Double
    
    'From the array of supplied points, find the one closest to the target point
    Dim i As Long
    For i = LBound(poiArray) To UBound(poiArray)
        tmpDistance = DistanceTwoPoints(targetX, targetY, poiArray(i).x, poiArray(i).y)
        If (tmpDistance < curMinDistance) Then
            curMinDistance = tmpDistance
            curMinIndex = i
        End If
    Next i
    
    'If the distance of the closest point falls below the allowed threshold, return that point's index.
    If (curMinDistance < minAllowedDistance) Then
        FindClosestPointInFloatArray = curMinIndex
    Else
        FindClosestPointInFloatArray = -1
    End If

End Function

'Retrieve the low-word value from a Long-type variable.  With thanks to Randy Birch for this function (http://vbnet.mvps.org/index.html?code/subclass/activation.htm)
Public Function LoWord(ByRef dw As Long) As Integer
   If (dw And &H8000&) Then
      LoWord = &H8000& Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

'Max/min functions
Public Function Max2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 > l2) Then Max2Int = l1 Else Max2Int = l2
End Function

Public Function Min2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 < l2) Then Min2Int = l1 Else Min2Int = l2
End Function

Public Function Max2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 > f2) Then Max2Float_Single = f1 Else Max2Float_Single = f2
End Function

Public Function Min2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 < f2) Then Min2Float_Single = f1 Else Min2Float_Single = f2
End Function

'Return the maximum of three floating point values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Max3Float(ByVal rR As Double, ByVal rG As Double, ByVal rB As Double) As Double
    If (rR > rG) Then
        If (rR > rB) Then
            Max3Float = rR
        Else
            Max3Float = rB
        End If
    ElseIf (rB > rG) Then
        Max3Float = rB
    Else
        Max3Float = rG
    End If
End Function

'Return the minimum of three floating point values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Min3Float(ByVal rR As Double, ByVal rG As Double, ByVal rB As Double) As Double
    If (rR < rG) Then
        If (rR < rB) Then
            Min3Float = rR
        Else
            Min3Float = rB
        End If
    ElseIf (rB < rG) Then
        Min3Float = rB
    Else
        Min3Float = rG
    End If
End Function

'Return the maximum of three integer values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Max3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR > rG) Then
        If (rR > rB) Then
            Max3Int = rR
        Else
            Max3Int = rB
        End If
    ElseIf (rB > rG) Then
        Max3Int = rB
    Else
        Max3Int = rG
    End If
End Function

'Return the minimum of three integer values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Min3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR < rG) Then
        If (rR < rB) Then
            Min3Int = rR
        Else
            Min3Int = rB
        End If
    ElseIf (rB < rG) Then
        Min3Int = rB
    Else
        Min3Int = rG
    End If
End Function

'Return the maximum value from an arbitrary list of floating point values
Public Function MaxArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
    If UBound(listOfValues) >= LBound(listOfValues) Then
                    
        Dim i As Long, numOfPoints As Long
        numOfPoints = (UBound(listOfValues) - LBound(listOfValues)) + 1
        
        Dim maxValue As Double
        maxValue = listOfValues(0)
        
        If numOfPoints > 1 Then
            For i = 1 To numOfPoints - 1
                If listOfValues(i) > maxValue Then maxValue = listOfValues(i)
            Next i
        End If
        
        MaxArbitraryListF = maxValue
        
    Else
        Debug.Print "No points provided - maxArbitraryListF() function failed!"
    End If
    
End Function

'Return the minimum value from an arbitrary list of floating point values
Public Function MinArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
    If UBound(listOfValues) >= LBound(listOfValues) Then
                    
        Dim i As Long, numOfPoints As Long
        numOfPoints = (UBound(listOfValues) - LBound(listOfValues)) + 1
        
        Dim minValue As Double
        minValue = listOfValues(0)
        
        If numOfPoints > 1 Then
            For i = 1 To numOfPoints - 1
                If listOfValues(i) < minValue Then minValue = listOfValues(i)
            Next i
        End If
        
        MinArbitraryListF = minValue
        
    Else
        Debug.Print "No points provided - minArbitraryListF() function failed!"
    End If
        
End Function

'This is a modified modulo function; it handles negative values specially to ensure they work with certain distort functions
Public Function Modulo(ByVal Quotient As Double, ByVal Divisor As Double) As Double
    Modulo = Quotient - Fix(Quotient / Divisor) * Divisor
    If (Modulo < 0#) Then Modulo = Modulo + Divisor
End Function

Public Function NearestInt(ByVal srcFloat As Single) As Long
    NearestInt = Int(srcFloat + 0.5)
End Function

'Given an array of points (in floating-point format), calculate the center point of the bounding rect.
Public Sub FindCenterOfFloatPoints(ByRef dstPoint As POINTFLOAT, ByRef srcPoints() As POINTFLOAT)

    Dim curMinX As Single, curMinY As Single, curMaxX As Single, curMaxY As Single
    curMinX = 9999999:   curMaxX = -9999999:   curMinY = 9999999:   curMaxY = -9999999
    
    'From the array of supplied points, find minimum and maximum (x, y) values
    Dim i As Long
    For i = LBound(srcPoints) To UBound(srcPoints)
        With srcPoints(i)
            If (.x < curMinX) Then curMinX = .x
            If (.x > curMaxX) Then curMaxX = .x
            If (.y < curMinY) Then curMinY = .y
            If (.y > curMaxY) Then curMaxY = .y
        End With
    Next i
    
    dstPoint.x = (curMaxX + curMinX) * 0.5
    dstPoint.y = (curMaxY + curMinY) * 0.5
    
End Sub

'Given a rectangle (as defined by width and height, not position), calculate the bounding rect required by a rotation of that rectangle.
Public Sub FindBoundarySizeOfRotatedRect(ByVal srcWidth As Double, ByVal srcHeight As Double, ByVal rotateAngle As Double, ByRef dstWidth As Double, ByRef dstHeight As Double, Optional ByVal padToIntegerValues As Boolean = True)

    'Convert the rotation angle to radians
    rotateAngle = rotateAngle * (PI_DIV_180)
    
    'Find the cos and sin of this angle and store the values
    Dim cosTheta As Double, sinTheta As Double
    cosTheta = Cos(rotateAngle)
    sinTheta = Sin(rotateAngle)
    
    'Create source and destination points
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim x11 As Double, x21 As Double, x31 As Double, x41 As Double
    
    Dim y1 As Double, y2 As Double, y3 As Double, y4 As Double
    Dim y11 As Double, y21 As Double, y31 As Double, y41 As Double
    
    'Position the points around (0, 0) to simplify the rotation code
    x1 = -srcWidth / 2
    x2 = srcWidth / 2
    x3 = srcWidth / 2
    x4 = -srcWidth / 2
    y1 = srcHeight / 2
    y2 = srcHeight / 2
    y3 = -srcHeight / 2
    y4 = -srcHeight / 2

    'Apply the rotation to each point
    x11 = x1 * cosTheta + y1 * sinTheta
    y11 = -x1 * sinTheta + y1 * cosTheta
    x21 = x2 * cosTheta + y2 * sinTheta
    y21 = -x2 * sinTheta + y2 * cosTheta
    x31 = x3 * cosTheta + y3 * sinTheta
    y31 = -x3 * sinTheta + y3 * cosTheta
    x41 = x4 * cosTheta + y4 * sinTheta
    y41 = -x4 * sinTheta + y4 * cosTheta
        
    'If the caller is using this for something like determining bounds of a rotated image, we need to convert all points to
    ' their "furthest from 0" integer amount.  Int() works on negative numbers, but a modified Ceiling()-type functions is
    ' required as VB oddly does not provide one.
    If padToIntegerValues Then ConvertArbitraryListToFurthestRoundedInt x11, x21, x31, x41, y11, y21, y31, y41
    
    'Find max/min values
    Dim xMin As Double, xMax As Double
    xMin = MinArbitraryListF(x11, x21, x31, x41)
    xMax = MaxArbitraryListF(x11, x21, x31, x41)
    
    Dim yMin As Double, yMax As Double
    yMin = MinArbitraryListF(y11, y21, y31, y41)
    yMax = MaxArbitraryListF(y11, y21, y31, y41)
    
    'Return the max/min values
    dstWidth = xMax - xMin
    dstHeight = yMax - yMin
    
End Sub

'Given a rectangle (as defined by width and height, not position), calculate new corner positions as an array of PointF structs.
Public Sub FindCornersOfRotatedRect(ByVal srcWidth As Double, ByVal srcHeight As Double, ByVal rotateAngle As Double, ByRef dstPoints() As POINTFLOAT, Optional ByVal arrayAlreadyDimmed As Boolean = False)

    'Convert the rotation angle to radians
    rotateAngle = rotateAngle * (PI_DIV_180)
    
    'Find the cos and sin of this angle and store the values
    Dim cosTheta As Double, sinTheta As Double
    cosTheta = Cos(rotateAngle)
    sinTheta = Sin(rotateAngle)
    
    'Create source and destination points
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim x11 As Double, x21 As Double, x31 As Double, x41 As Double
    
    Dim y1 As Double, y2 As Double, y3 As Double, y4 As Double
    Dim y11 As Double, y21 As Double, y31 As Double, y41 As Double
    
    'Position the points around (0, 0) to simplify the rotation code
    Dim halfWidth As Double, halfHeight As Double
    halfWidth = srcWidth / 2
    halfHeight = srcHeight / 2
    
    x1 = -halfWidth
    x2 = halfWidth
    x3 = halfWidth
    x4 = -halfWidth
    y1 = -halfHeight
    y2 = -halfHeight
    y3 = halfHeight
    y4 = halfHeight

    'Apply the rotation to each point
    x11 = x1 * cosTheta + y1 * sinTheta
    y11 = -x1 * sinTheta + y1 * cosTheta
    x21 = x2 * cosTheta + y2 * sinTheta
    y21 = -x2 * sinTheta + y2 * cosTheta
    x31 = x3 * cosTheta + y3 * sinTheta
    y31 = -x3 * sinTheta + y3 * cosTheta
    x41 = x4 * cosTheta + y4 * sinTheta
    y41 = -x4 * sinTheta + y4 * cosTheta
    
    'Fill the destination array with the rotated points, translated back into the original coordinate space for convenience
    If (Not arrayAlreadyDimmed) Then ReDim dstPoints(0 To 3) As POINTFLOAT
    dstPoints(0).x = x11 + halfWidth
    dstPoints(0).y = y11 + halfHeight
    dstPoints(1).x = x21 + halfWidth
    dstPoints(1).y = y21 + halfHeight
    dstPoints(3).x = x31 + halfWidth
    dstPoints(3).y = y31 + halfHeight
    dstPoints(2).x = x41 + halfWidth
    dstPoints(2).y = y41 + halfHeight
    
End Sub

Public Function RadiansToDegrees(ByVal srcRadian As Double) As Double
    Const ONE_DIV_PI As Double = 1# / PI
    RadiansToDegrees = (srcRadian * 180#) * ONE_DIV_PI
End Function

Public Function DegreesToRadians(ByVal srcDegrees As Double) As Double
    Const ONE_DIV_180 As Double = 1# / 180#
    DegreesToRadians = (srcDegrees * PI) * ONE_DIV_180
End Function

'Helper function to rotate one arbitrary point around another arbitrary point.
Public Function RotatePointAroundPoint(ByVal rotateX As Single, ByVal rotateY As Single, ByVal centerX As Single, ByVal centerY As Single, ByVal angleInRadians As Single, ByRef newX As Single, ByRef newY As Single)

    'For performance reasons, it's easier to cache the cos and sin of the angle in question
    Dim sinAngle As Double, cosAngle As Double
    sinAngle = Sin(angleInRadians)
    cosAngle = Cos(angleInRadians)
    
    'Rotation works the same way as it does around (0, 0), except we transform the center position before and
    ' after rotation.
    newX = cosAngle * (rotateX - centerX) - sinAngle * (rotateY - centerY) + centerX
    newY = cosAngle * (rotateY - centerY) + sinAngle * (rotateX - centerX) + centerY
    
End Function

'Given a RectF object, enlarge the boundaries to produce an integer-only RectF that is guaranteed
' to encompass the entire original rect.  (Said another way, the modified rect will *never* be smaller
' than the original rect.)
Public Sub GetIntClampedRectF(ByRef srcRectF As RECTF)
    Dim xOffset As Single, yOffset As Single
    xOffset = srcRectF.Left - Int(srcRectF.Left)
    yOffset = srcRectF.Top - Int(srcRectF.Top)
    srcRectF.Left = Int(srcRectF.Left)
    srcRectF.Top = Int(srcRectF.Top)
    srcRectF.Width = Int(srcRectF.Width + xOffset + 0.9999)
    srcRectF.Height = Int(srcRectF.Height + yOffset + 0.9999)
End Sub

'Given a RectF object, adjust the boundaries to produce an integer-only RectF that approximates the
' original rect as closely as possible.  (This rect *may* be smaller than the original; for a rect
' guaranteed to encompass the entire original area, use GetIntClampedRectF(), above.)
Public Sub GetNearestIntRectF(ByRef srcRectF As RECTF)
    Dim xOffset As Single, yOffset As Single
    xOffset = PDMath.Frac(srcRectF.Left)
    yOffset = PDMath.Frac(srcRectF.Top)
    srcRectF.Left = Int(srcRectF.Left)
    srcRectF.Top = Int(srcRectF.Top)
    If (PDMath.Frac(srcRectF.Width + xOffset) >= 0.5) Then srcRectF.Width = Int(srcRectF.Width + 1#) Else srcRectF.Width = Int(srcRectF.Width)
    If (PDMath.Frac(srcRectF.Height + yOffset) >= 0.5) Then srcRectF.Height = Int(srcRectF.Height + 1#) Else srcRectF.Height = Int(srcRectF.Height)
End Sub

Public Function ClampL(ByVal srcL As Long, ByVal minL As Long, ByVal maxL As Long) As Long
    If (srcL < minL) Then
        ClampL = minL
    ElseIf (srcL > maxL) Then
        ClampL = maxL
    Else
        ClampL = srcL
    End If
End Function

Public Function ClampF(ByVal srcF As Double, ByVal minF As Double, ByVal maxF As Double) As Double
    If (srcF < minF) Then
        ClampF = minF
    ElseIf (srcF > maxF) Then
        ClampF = maxF
    Else
        ClampF = srcF
    End If
End Function

Public Function ConvertDPIToPels(ByVal srcDPI As Double) As Double
    ConvertDPIToPels = (srcDPI / 2.54) * 100#
End Function
