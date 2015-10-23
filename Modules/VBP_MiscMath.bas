Attribute VB_Name = "Math_Functions"
'***************************************************************************
'Specialized Math Routines
'Copyright 2013-2015 by Tanner Helland and Audioglider
'Created: 13/June/13
'Last updated: 22/May/14
'Last update: fixed convertToFraction() function to work with non-English locales
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June '13.  This module is now used to store all the random bits of specialized math required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRectL Lib "user32" Alias "PtInRect" (ByRef lpRect As RECTL, ByVal x As Long, ByVal y As Long) As Long

'See if a point lies inside a rect (integer)
Public Function isPointInRect(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECT) As Boolean

    If PtInRect(srcRect, ptX, ptY) = 0 Then
        isPointInRect = False
    Else
        isPointInRect = True
    End If

End Function

'See if a point lies inside a RectL struct
Public Function isPointInRectL(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECTL) As Boolean

    If PtInRectL(srcRect, ptX, ptY) = 0 Then
        isPointInRectL = False
    Else
        isPointInRectL = True
    End If

End Function

'See if a point lies inside a rect (float)
Public Function isPointInRectF(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECTF) As Boolean

    'There's no GDI function for floating-point rects, so we must do this manually
    With srcRect
    
        'Check x boundaries
        If (ptX >= .Left) And (ptX <= (.Left + .Width)) Then
        
            'Check y boundaries
            If (ptY >= .Top) And (ptY <= (.Top + .Height)) Then
                isPointInRectF = True
            Else
                isPointInRectF = False
            End If
        
        Else
            isPointInRectF = False
        End If
    
    End With

End Function

'Find the union rect of two floating-point rects.  (This is the smallest rect that contains both rects.)
Public Sub UnionRectF(ByRef dstRect As RECTF, ByRef srcRect As RECTF, ByRef srcRect2 As RECTF, Optional ByVal widthAndHeightAreReallyRightAndBottom As Boolean = False)

    'Union rects are easy: find the min top/left, and the max bottom/right
    With dstRect
        
        If srcRect.Left < srcRect2.Left Then
            .Left = srcRect.Left
        Else
            .Left = srcRect2.Left
        End If
        
        If srcRect.Top < srcRect2.Top Then
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
        If srcRight < srcRight2 Then srcRight = srcRight2
        
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
        
        If srcBottom < srcBottom2 Then srcBottom = srcBottom2
        
        If widthAndHeightAreReallyRightAndBottom Then
            .Height = srcBottom
        Else
            .Height = srcBottom - .Top
        End If
        
    End With

End Sub

'Given an arbitrary output range and input range, convert a value from the input range to the output range
' Thank you to expert coder audioglider for contributing this function.
Public Function convertRange(ByVal originalStart As Double, ByVal originalEnd As Double, ByVal newStart As Double, ByVal newEnd As Double, ByVal Value As Double) As Double
    Dim dScale As Double
    
    dScale = (newEnd - newStart) / (originalEnd - originalStart)
    convertRange = (newStart + ((Value - originalStart) * dScale))
End Function

'Convert a decimal to a near-identical fraction using vector math.
' This excellent function comes courtesy of VB6 coder LaVolpe.  I have modified it slightly to suit PhotoDemon's unique needs.
' You can download the original at this link (good as of 13 June 2014): http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=61596&lngWId=1
Public Sub convertToFraction(ByVal v As Double, w As Double, n As Double, d As Double, Optional ByVal maxDenomDigits As Byte, Optional ByVal Accuracy As Double = 100#)

    Const MaxTerms As Integer = 50          'Limit to prevent infinite loop
    Const MinDivisor As Double = 1E-16      'Limit to prevent divide by zero
    Const MaxError As Double = 1E-50        'How close is enough
    Dim f As Double                         'Fraction being converted
    Dim a As Double     'Current term in continued fraction
    Dim n1 As Double    'Numerator, denominator of last approx
    Dim D1 As Double
    Dim n2 As Double    'Numerator, denominator of previous approx
    Dim D2 As Double
    Dim i As Integer
    Dim t As Double
    Dim maxDenom As Double
    Dim bIsNegative As Boolean
    
    If maxDenomDigits = 0 Or maxDenomDigits > 17 Then maxDenomDigits = 17
    maxDenom = 10 ^ maxDenomDigits
    If Accuracy > 100 Or Accuracy < 1 Then Accuracy = 100
    Accuracy = Accuracy / 100#
    
    bIsNegative = (v < 0)
    w = Abs(Fix(v))
    'V = Abs(V) - W << subtracting doubles can change the decimal portion by adding more numeral at end
    'TANNER'S NOTE: the original version of this included +1 to the string length, which gave me consistent errors.  So I removed it.
    v = CDbl(Mid$(CStr(Abs(v)), Len(CStr(w))))
    
    ' check for no decimal or zero
    If v = 0 Then GoTo RtnResult
    
    f = v                       'Initialize fraction being converted
    
    n1 = 1                      'Initialize fractions with 1/0, 0/1
    D1 = 0
    n2 = 0
    D2 = 1

    On Error GoTo RtnResult
    For i = 0 To MaxTerms
        a = Fix(f)              'Get next term
        f = f - a               'Get new divisor
        n = n1 * a + n2         'Calculate new fraction
        d = D1 * a + D2
        n2 = n1                 'Save last two fractions
        D2 = D1
        n1 = n
        D1 = d
                                'Quit if dividing by zero
        If f < MinDivisor Then Exit For

                                'Quit if close enough
        t = n / d               ' A=zero indicates exact match or extremely close
        a = Abs(v - t)          ' Difference btwn actual V and calculated T
        If a < MaxError Then Exit For
                                'Quit if max denominator digits encountered
        If d > maxDenom Then Exit For
                                ' Quit if preferred accuracy accomplished
        If n Then
            If t > v Then t = v / t Else t = t / v
            If t >= Accuracy And Abs(t) < 1 Then Exit For
        End If
        f = 1# / f               'Take reciprocal
    Next i

RtnResult:
    If Err Or d > maxDenom Then
        ' in above case, use the previous best N & D
        If D2 = 0 Then
            n = n1
            d = D1
        Else
            d = D2
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
Public Sub convertAspectRatio(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long, Optional ByVal fitInclusive As Boolean = True)
    
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0) And (dstHeight > 0) Then
        srcAspect = srcWidth / srcHeight
        dstAspect = dstWidth / dstHeight
    Else
        Exit Sub
    End If
    
    Dim aspectLarger As Boolean
    If srcAspect > dstAspect Then aspectLarger = True Else aspectLarger = False
    
    'Exclusive fitting fits the opposite dimension, so simply reverse the way the dimensions are calculated
    If Not fitInclusive Then aspectLarger = Not aspectLarger
    
    If aspectLarger Then
        newWidth = dstWidth
        newHeight = CDbl(srcHeight / srcWidth) * newWidth
    Else
        newHeight = dstHeight
        newWidth = CDbl(srcWidth / srcHeight) * newHeight
    End If
    
End Sub

'Return the distance between two values on the same line
Public Function distanceOneDimension(ByVal x1 As Double, ByVal x2 As Double) As Double
    distanceOneDimension = Sqr((x1 - x2) ^ 2)
End Function

'Return the distance between two points
Public Function distanceTwoPoints(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    distanceTwoPoints = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'Return the distance between two points, but ignores the square root function; if calculating something simple, like "minimum distance only",
' we only need relative values - not absolute ones - so we can skip that step for an extra performance boost.
Public Function distanceTwoPointsShortcut(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    distanceTwoPointsShortcut = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)
End Function

'Return the distance between two 3D points
Public Function distanceThreeDimensions(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    distanceThreeDimensions = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
End Function

'Given two intersecting lines, return the angle between them (e.g. the inner product: https://en.wikipedia.org/wiki/Inner_product_space)
Public Function angleBetweenTwoIntersectingLines(ByRef ptIntersect As POINTFLOAT, ByRef pt1 As POINTFLOAT, ByRef pt2 As POINTFLOAT, Optional ByVal returnResultInDegrees As Boolean = True) As Double
    
    Dim dx1i As Double, dy1i As Double, dx2i As Double, dy2i As Double
    dx1i = pt1.x - ptIntersect.x
    dy1i = pt1.y - ptIntersect.y
    
    dx2i = pt2.x - ptIntersect.x
    dy2i = pt2.y - ptIntersect.y
    
    Dim m12 As Double, m13 As Double
    m12 = Sqr(dx1i * dx1i + dy1i * dy1i)
    m13 = Sqr(dx2i * dx2i + dy2i * dy2i)
    
    angleBetweenTwoIntersectingLines = Acos((dx1i * dx2i + dy1i * dy2i) / (m12 * m13))
    
    If returnResultInDegrees Then
        angleBetweenTwoIntersectingLines = angleBetweenTwoIntersectingLines / PI_DIV_180
    End If
    
End Function

'Return the arctangent of two values (rise / run)
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
 
    If (y = 0) And (x = 0) Then
        Atan2 = 0
        Exit Function
    End If
 
    If y > 0 Then
        If x >= y Then
            Atan2 = Atn(y / x)
        ElseIf x <= -y Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = PI_HALF - Atn(x / y)
        End If
    Else
        If x >= -y Then
            Atan2 = Atn(y / x)
        ElseIf x <= y Then
            Atan2 = Atn(y / x) - PI
        Else
            Atan2 = -Atn(x / y) - PI_HALF
        End If
    End If
 
End Function

'Arcsine function
Public Function Asin(ByVal x As Double) As Double
    If (x > 1) Or (x < -1) Then x = 1
    Asin = Atan2(x, Sqr(1 - x * x))
End Function

'Arccosine function
Public Function Acos(ByVal x As Double) As Double
    If (x > 1) Or (x < -1) Then x = 1
    Acos = Atan2(Sqr(1 - x * x), x)
End Function

'Return the maximum of two floating point values
Public Function Max2Float_Single(f1 As Single, f2 As Single) As Single
    If f1 > f2 Then
        Max2Float_Single = f1
    Else
        Max2Float_Single = f2
    End If
End Function

Public Function Min2Float_Single(f1 As Single, f2 As Single) As Single
    If f1 < f2 Then
        Min2Float_Single = f1
    Else
        Min2Float_Single = f2
    End If
End Function

'Return the maximum of three floating point values
Public Function Max3Float(rR As Double, rG As Double, rB As Double) As Double
   If (rR > rG) Then
      If (rR > rB) Then
         Max3Float = rR
      Else
         Max3Float = rB
      End If
   Else
      If (rB > rG) Then
         Max3Float = rB
      Else
         Max3Float = rG
      End If
   End If
End Function

'Return the minimum of three floating point values
Public Function Min3Float(rR As Double, rG As Double, rB As Double) As Double
   If (rR < rG) Then
      If (rR < rB) Then
         Min3Float = rR
      Else
         Min3Float = rB
      End If
   Else
      If (rB < rG) Then
         Min3Float = rB
      Else
         Min3Float = rG
      End If
   End If
End Function

'Return the maximum of three integer values
Public Function Max3Int(rR As Long, rG As Long, rB As Long) As Long
   If (rR > rG) Then
      If (rR > rB) Then
         Max3Int = rR
      Else
         Max3Int = rB
      End If
   Else
      If (rB > rG) Then
         Max3Int = rB
      Else
         Max3Int = rG
      End If
   End If
End Function

'Return the minimum of three integer values
Public Function Min3Int(rR As Long, rG As Long, rB As Long) As Long
   If (rR < rG) Then
      If (rR < rB) Then
         Min3Int = rR
      Else
         Min3Int = rB
      End If
   Else
      If (rB < rG) Then
         Min3Int = rB
      Else
         Min3Int = rG
      End If
   End If
End Function

'Return the maximum value from an arbitrary list of floating point values
Public Function maxArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
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
        
        maxArbitraryListF = maxValue
        
    Else
        Debug.Print "No points provided - maxArbitraryListF() function failed!"
    End If
        
End Function

'Return the minimum value from an arbitrary list of floating point values
Public Function minArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
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
        
        minArbitraryListF = minValue
        
    Else
        Debug.Print "No points provided - minArbitraryListF() function failed!"
    End If
        
End Function

'Given a list of floating-point values, convert each to its integer equivalent *furthest* from 0.
' Said another way, round negative numbers down, and positive numbers up.  This is often relevant in PD when performing
' coordinate conversions that are ultimately mapped to pixel locations, and we need to bounds-check corner coordinates
' in advance and push them away from 0, so any partially-covered pixels are converted to fully-covered ones.
Public Function convertArbitraryListToFurthestRoundedInt(ParamArray listOfValues() As Variant)
    
    If UBound(listOfValues) >= LBound(listOfValues) Then
        
        Dim i As Long
        For i = LBound(listOfValues) To UBound(listOfValues)
            If listOfValues(i) < 0 Then
                listOfValues(i) = Int(listOfValues(i))
            Else
                listOfValues(i) = IIf(listOfValues(i) = Int(listOfValues(i)), listOfValues(i), Int(listOfValues(i)) + 1)
            End If
        Next i
        
    Else
        Debug.Print "No points provided - convertArbitraryFListToRoundedInt() function failed!"
    End If

End Function

Public Sub convertPolarToCartesian(ByVal srcAngle As Double, ByVal srcRadius As Double, ByRef dstX As Double, ByRef dstY As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)

    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub

'This is a modified module function; it handles negative values specially to ensure they work with certain distort functions
Public Function Modulo(ByVal Quotient As Double, ByVal Divisor As Double) As Double
    Modulo = Quotient - Fix(Quotient / Divisor) * Divisor
    If Modulo < 0 Then Modulo = Modulo + Divisor
End Function

'Retrieve the low-word value from a Long-type variable.  With thanks to Randy Birch for this function (http://vbnet.mvps.org/index.html?code/subclass/activation.htm)
Public Function LoWord(dw As Long) As Integer
   If dw And &H8000& Then
      LoWord = &H8000& Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

'Given an array of points, find the closest one to a target location.  If none fall below a minimum distance threshold, return -1.
' (This function is used by many bits of mouse interaction code, to see if the user has clicked on something interesting.)
Public Function findClosestPointInArray(ByVal targetX As Double, ByVal targetY As Double, ByVal minAllowedDistance As Double, ByRef poiArray() As POINTAPI) As Long

    Dim curMinDistance As Double, curMinIndex As Long
    curMinDistance = &HFFFFFFF
    curMinIndex = -1
    
    Dim tmpDistance As Double
    
    'From the array of supplied points, find the one closest to the target point
    Dim i As Long
    For i = LBound(poiArray) To UBound(poiArray)
        tmpDistance = distanceTwoPoints(targetX, targetY, poiArray(i).x, poiArray(i).y)
        If tmpDistance < curMinDistance Then
            curMinDistance = tmpDistance
            curMinIndex = i
        End If
    Next i
    
    'If the distance of the closest point falls below the allowed threshold, return that point's index.
    If curMinDistance < minAllowedDistance Then
        findClosestPointInArray = curMinIndex
    Else
        findClosestPointInArray = -1
    End If

End Function

'Given an array of points (in floating-point format), find the closest one to a target location.  If none fall below a minimum distance threshold,
' return -1.  (This function is used by many bits of mouse interaction code, to see if the user has clicked on something interesting.)
Public Function findClosestPointInFloatArray(ByVal targetX As Double, ByVal targetY As Double, ByVal minAllowedDistance As Double, ByRef poiArray() As POINTFLOAT) As Long

    Dim curMinDistance As Double, curMinIndex As Long
    curMinDistance = &HFFFFFFF
    curMinIndex = -1
    
    Dim tmpDistance As Double
    
    'From the array of supplied points, find the one closest to the target point
    Dim i As Long
    For i = LBound(poiArray) To UBound(poiArray)
        tmpDistance = distanceTwoPoints(targetX, targetY, poiArray(i).x, poiArray(i).y)
        If tmpDistance < curMinDistance Then
            curMinDistance = tmpDistance
            curMinIndex = i
        End If
    Next i
    
    'If the distance of the closest point falls below the allowed threshold, return that point's index.
    If curMinDistance < minAllowedDistance Then
        findClosestPointInFloatArray = curMinIndex
    Else
        findClosestPointInFloatArray = -1
    End If

End Function

'Given a rectangle (as defined by width and height, not position), calculate the bounding rect required by a rotation of that rectangle.
Public Sub findBoundarySizeOfRotatedRect(ByVal srcWidth As Double, ByVal srcHeight As Double, ByVal rotateAngle As Double, ByRef dstWidth As Double, ByRef dstHeight As Double, Optional ByVal padToIntegerValues As Boolean = True)

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
    If padToIntegerValues Then convertArbitraryListToFurthestRoundedInt x11, x21, x31, x41, y11, y21, y31, y41
    
    'Find max/min values
    Dim xMin As Double, xMax As Double
    xMin = minArbitraryListF(x11, x21, x31, x41)
    xMax = maxArbitraryListF(x11, x21, x31, x41)
    
    Dim yMin As Double, yMax As Double
    yMin = minArbitraryListF(y11, y21, y31, y41)
    yMax = maxArbitraryListF(y11, y21, y31, y41)
    
    'Return the max/min values
    dstWidth = xMax - xMin
    dstHeight = yMax - yMin
    
End Sub

'Given a rectangle (as defined by width and height, not position), calculate new corner positions as an array of PointF structs.
Public Sub findCornersOfRotatedRect(ByVal srcWidth As Double, ByVal srcHeight As Double, ByVal rotateAngle As Double, ByRef dstPoints() As POINTFLOAT, Optional ByVal arrayAlreadyDimmed As Boolean = False)

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
    If Not arrayAlreadyDimmed Then ReDim dstPoints(0 To 3) As POINTFLOAT
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
    RadiansToDegrees = (srcRadian * 180) / PI
End Function

Public Function DegreesToRadians(ByVal srcDegrees As Double) As Double
    DegreesToRadians = (srcDegrees * PI) / 180
End Function

Public Function ClampL(ByVal srcL As Long, ByVal minL As Long, ByVal maxL As Long) As Long
    If srcL < minL Then
        ClampL = minL
    ElseIf srcL > maxL Then
        ClampL = maxL
    Else
        ClampL = srcL
    End If
End Function

Public Function ClampF(ByVal srcF As Double, ByVal minF As Double, ByVal maxF As Double) As Double
    If srcF < minF Then
        ClampF = minF
    ElseIf srcF > maxF Then
        ClampF = maxF
    Else
        ClampF = srcF
    End If
End Function
