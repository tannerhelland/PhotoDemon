Attribute VB_Name = "PDMath"
'***************************************************************************
'Specialized Math Routines
'Copyright 2013-2026 by Tanner Helland
'Created: 13/June/13
'Last updated: 06/May/22
'Last update: new function for direct path simplification, which handles all line extraction for you
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June '13.  This module is now used to store all the random bits of specialized math required by the program.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Const PI As Double = 3.14159265358979
Public Const PI_HALF As Double = 1.5707963267949
Public Const PI_DOUBLE As Double = 6.28318530717958
Public Const PI_DIV_180 As Double = 0.017453292519943
Public Const PI_14 As Double = 0.785398163397448
Public Const PI_34 As Double = 2.35619449019234

'Noise generators
Public Enum PD_NoiseGenerator
    ng_Perlin = 0
    ng_Simplex = 1
    ng_OpenSimplex = 2
End Enum

#If False Then
    Private Const ng_Perlin = 0, ng_Simplex = 1, ng_OpenSimplex = 2
#End If

'Marching squares line simplification (see SimplifyLinesFromMarchingSquares() function)
Private Enum PD_PointSimplify
    ps_Essential = 0
    ps_East = 1
    ps_West = 2
    ps_North = 4
    ps_South = 8
    ps_Remove = 16
End Enum

#If False Then
    Private Const ps_Essential = 0, ps_East = 1, ps_West = 2, ps_North = 4, ps_South = 8, ps_Remove = 16
#End If

Private Declare Function IntersectRect Lib "user32" (ByVal ptrDstRect As Long, ByVal ptrSrcRect1 As Long, ByVal ptrSrcRect2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRectL Lib "user32" Alias "PtInRect" (ByRef lpRect As RectL, ByVal x As Long, ByVal y As Long) As Long

'Rect intersect calculation; wraps IntersectRect API and returns VB boolean if rects intersect
Public Function IntersectRectL(ByRef dstRect As RectL, ByRef srcRect1 As RectL, ByRef srcRect2 As RectL) As Boolean
    IntersectRectL = (IntersectRect(VarPtr(dstRect), VarPtr(srcRect1), VarPtr(srcRect2)) <> 0)
End Function

'See if a point lies inside a rect (integer)
Public Function IsPointInRect(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RECT) As Boolean
    IsPointInRect = (PtInRect(srcRect, ptX, ptY) <> 0)
End Function

'See if a point lies inside a RectL struct
Public Function IsPointInRectL(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RectL) As Boolean
    IsPointInRectL = (PtInRectL(srcRect, ptX, ptY) <> 0)
End Function

'See if a point lies inside a rect (float)
Public Function IsPointInRectF(ByVal ptX As Long, ByVal ptY As Long, ByRef srcRect As RectF) As Boolean
    With srcRect
        If (ptX >= .Left) And (ptX <= (.Left + .Width)) Then
            IsPointInRectF = ((ptY >= .Top) And (ptY <= (.Top + .Height)))
        Else
            IsPointInRectF = False
        End If
    End With
End Function

'Is an arbitrary polygon convex?  Algorithm adapted from an original algorithm by Rory Daulton as shared
' at StackOverflow (link good as of April '22):
' https://stackoverflow.com/questions/471962/how-do-i-efficiently-determine-if-a-polygon-is-convex-non-convex-or-complex
'
'Note that this version is deliberately modified to allow points to sit on the same line (angle = 0);
' because PhotoDemon always passes points in clockwise order, we don't need to worry about how this 0
' would otherwise affect checks for directional changes in angle.
'
'Returns: TRUE if convex, FALSE otherwise
Public Function IsPolygonConvex(ByRef listOfPoints() As PointFloat, ByVal numOfPoints As Long) As Boolean
    
    'Assume failure by default; there are many obvious fail states that allow us to abandon early
    IsPolygonConvex = False
    
    'Errors are not expected, but if they occur FALSE will be returned
    On Error GoTo NotConvex
    
    'Perform some preliminary failsafe checks
    If (numOfPoints > UBound(listOfPoints) + 1) Then Exit Function
    If (numOfPoints < 3) Then Exit Function
    
    'Populate initial values
    Dim oldPoint As PointFloat, newPoint As PointFloat
    oldPoint = listOfPoints(numOfPoints - 2)
    newPoint = listOfPoints(numOfPoints - 1)
    
    Dim newDirection As Double, oldDirection As Double, angle As Double, angleSum As Double
    newDirection = PDMath.Atan2(newPoint.y - oldPoint.y, newPoint.x - oldPoint.x)
    angleSum = 0#
    
    Dim curOrientation As Double
    
    'Check each point (specifically, the polygon side that terminates at each point)
    ' along with both that point's angle *and* the accumulated angles in the full polygon.
    Dim i As Long
    For i = 0 To numOfPoints - 1
        
        'Update points
        oldDirection = newDirection
        oldPoint = newPoint
        newPoint = listOfPoints(i)
        newDirection = PDMath.Atan2(newPoint.y - oldPoint.y, newPoint.x - oldPoint.x)
        
        'If consecutive points match, reject the polygon
        If (oldPoint.x = newPoint.x) And (oldPoint.y = newPoint.y) Then Exit Function
        
        'Calculate and check direction-change angles.  (We will normalize the angle to ensure correct behavior.)
        angle = newDirection - oldDirection
        If (angle <= -PI) Then
            angle = angle + PI_DOUBLE
        ElseIf (angle > PI) Then
            angle = angle - PI_DOUBLE
        End If
        
        'First time through the loop, initialize orientation
        If (i = 0) Then
            If (angle >= 0#) Then curOrientation = 1# Else curOrientation = -1#
        
        'On subsequent passes, look for orientation changes
        Else
            If (curOrientation * angle <= 0#) Then Exit Function
        End If
        
        'Accumulate angles
        angleSum = angleSum + angle
        
    Next i
    
    'Check that the total number of full turns is +/- 1
    If Abs(Int(angleSum / PI_DOUBLE + 0.5)) = 1 Then IsPolygonConvex = True
    
    Exit Function
    
NotConvex:
    IsPolygonConvex = False
    
End Function

Public Function PopulateRectL(ByVal srcLeft As Long, ByVal srcTop As Long, ByVal srcRight As Long, ByVal srcBottom As Long) As RectL
    PopulateRectL.Left = srcLeft
    PopulateRectL.Top = srcTop
    PopulateRectL.Right = srcRight
    PopulateRectL.Bottom = srcBottom
End Function

'Translate values on the scale [0, 1] to an arbitrary new scale.
Public Function TranslateValue_AbsToRel(ByVal srcValue As Single, ByVal srcMax As Single) As Single
    If (srcMax <> 0!) Then TranslateValue_AbsToRel = srcValue / srcMax Else TranslateValue_AbsToRel = 0!
End Function

'Translate values from some arbitrary scale to [0, 1].
Public Function TranslateValue_RelToAbs(ByVal srcValue As Single, ByVal srcMax As Single) As Single
    TranslateValue_RelToAbs = srcValue * srcMax
End Function

'Find the union rect of two floating-point rects.  (This is the smallest rect that contains both rects.)
Public Sub UnionRectF(ByRef dstRect As RectF, ByRef srcRect As RectF, ByRef srcRect2 As RectF, Optional ByVal widthAndHeightAreReallyRightAndBottom As Boolean = False)

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

'Arccosine function
Public Function Acos(ByVal x As Double) As Double
    If (x > 1#) Or (x < -1#) Then x = 1#
    Acos = Atan2(Sqr(1# - x * x), x)
End Function

'Arcsine function
Public Function Asin(ByVal x As Double) As Double
    If (x > 1#) Or (x < -1#) Then x = 1#
    Asin = Atan2(x, Sqr(1# - x * x))
End Function

'Given two intersecting lines, return the angle between them (e.g. the inner product: https://en.wikipedia.org/wiki/Inner_product_space)
Public Function AngleBetweenTwoIntersectingLines(ByRef ptIntersect As PointFloat, ByRef pt1 As PointFloat, ByRef pt2 As PointFloat, Optional ByVal returnResultInDegrees As Boolean = True) As Double
    
    Dim dx1i As Double, dy1i As Double, dx2i As Double, dy2i As Double
    dx1i = pt1.x - ptIntersect.x
    dy1i = pt1.y - ptIntersect.y
    
    dx2i = pt2.x - ptIntersect.x
    dy2i = pt2.y - ptIntersect.y
    
    Dim m12 As Double, m13 As Double
    m12 = Sqr(dx1i * dx1i + dy1i * dy1i)
    m13 = Sqr(dx2i * dx2i + dy2i * dy2i)
    
    AngleBetweenTwoIntersectingLines = Acos((dx1i * dx2i + dy1i * dy2i) / (m12 * m13))
    
    If returnResultInDegrees Then AngleBetweenTwoIntersectingLines = AngleBetweenTwoIntersectingLines / PI_DIV_180
    
End Function

Public Function AreRectFsEqual(ByRef srcRectF1 As RectF, ByRef srcRectf2 As RectF) As Boolean
    AreRectFsEqual = VBHacks.MemCmp(VarPtr(srcRectF1), VarPtr(srcRectf2), LenB(srcRectF1))
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

'Fast and easy technique for converting an arbitrary floating-point value to a fraction.  Developed with thanks to
' multiple authors at: https://stackoverflow.com/questions/95727/how-to-convert-floats-to-human-readable-fractions
Public Sub ConvertToFraction(ByVal srcValue As Double, ByRef dstNumerator As Long, ByRef dstDenominator As Long, Optional ByVal epsilon As Double = 0.001)

    dstNumerator = 1
    dstDenominator = 1
    
    Dim fracTest As Double
    fracTest = 1#
    
    Do While (Abs(fracTest - srcValue) > epsilon)
    
        If (fracTest < srcValue) Then
            dstNumerator = dstNumerator + 1
        Else
            dstDenominator = dstDenominator + 1
            dstNumerator = Int(srcValue * dstDenominator + 0.5)
        End If
        
        fracTest = CDbl(dstNumerator) / CDbl(dstDenominator)
        
    Loop
    
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

'Return the perpendicular distance between an arbitrary point and a line
Public Function DistancePerpendicular(ByVal ptX As Single, ByVal ptY As Single, ByVal lineX1 As Single, ByVal lineY1 As Single, ByVal lineX2 As Single, ByVal lineY2 As Single) As Single
    DistancePerpendicular = Sqr((lineY2 - lineY1) * (lineY2 - lineY1) + (lineX2 - lineX1) * (lineX2 - lineX1))
    If (DistancePerpendicular <> 0!) Then DistancePerpendicular = ((lineY2 - lineY1) * ptX - (lineX2 - lineX1) * ptY + (lineX2 * lineY1) - (lineY2 * lineX1)) / DistancePerpendicular
    If (DistancePerpendicular < 0!) Then DistancePerpendicular = -1 * DistancePerpendicular
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
    DistanceThreeDimensions = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2) + (z1 - z2) * (z1 - z2))
End Function

Public Function Distance3D_FastFloat(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
    Distance3D_FastFloat = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2) + (z1 - z2) * (z1 - z2))
End Function

Public Function Frac(ByVal srcValue As Double) As Double
    Frac = srcValue - Int(srcValue)
End Function

'Given a list of floating-point values, convert each to its integer equivalent *furthest* from 0.
' Said another way, round negative numbers down, and positive numbers up.  This is often relevant in PD when performing
' coordinate conversions that are ultimately mapped to pixel locations, and we need to bounds-check corner coordinates
' in advance and push them away from 0, so any partially-covered pixels are converted to fully-covered ones.
Public Sub ConvertArbitraryListToFurthestRoundedInt(ParamArray listOfValues() As Variant)
    
    If (UBound(listOfValues) >= LBound(listOfValues)) Then
        
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
        PDDebug.LogAction "No points provided - ConvertArbitraryListToFurthestRoundedInt() function failed!"
    End If

End Sub

Public Sub ConvertCartesianToPolar(ByVal srcX As Double, ByVal srcY As Double, ByRef dstRadius As Double, ByRef dstAngle As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)
    srcX = srcX - centerX
    srcY = srcY - centerY
    dstRadius = Sqr(srcX * srcX + srcY * srcY)
    dstAngle = PDMath.Atan2(srcY, srcX)
End Sub

Public Sub ConvertPolarToCartesian(ByVal srcAngle As Double, ByVal srcRadius As Double, ByRef dstX As Double, ByRef dstY As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)

    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub

Public Sub ConvertPolarToCartesian_Sng(ByVal srcAngle As Single, ByVal srcRadius As Single, ByRef dstX As Single, ByRef dstY As Single, Optional ByVal centerX As Single = 0!, Optional ByVal centerY As Single = 0!)

    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub
'Given an array of points, find the closest one to a target location.  If none fall below a minimum distance threshold, return -1.
' (This function is used by many bits of mouse interaction code, to see if the user has clicked on something interesting.)
Public Function FindClosestPointInArray(ByVal targetX As Double, ByVal targetY As Double, ByVal minAllowedDistance As Double, ByRef poiArray() As PointAPI) As Long

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
Public Function FindClosestPointInFloatArray(ByVal targetX As Single, ByVal targetY As Single, ByVal minAllowedDistance As Single, ByRef poiArray() As PointFloat) As Long

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

'Log variants
Public Function Log10(ByVal srcValue As Double) As Double
    Const INV_LOG_OF_10 As Double = 1# / 2.30258509     'Ln(10) = 2.30258509
    Log10 = Log(srcValue) * INV_LOG_OF_10
End Function

Public Function LogBaseN(ByVal srcValue As Double, ByVal srcBase As Double) As Double
    LogBaseN = Log(srcValue) / Log(srcBase)
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
Public Function Max2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 > f2) Then Max2Float_Single = f1 Else Max2Float_Single = f2
End Function

Public Function Max2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 > l2) Then Max2Int = l1 Else Max2Int = l2
End Function

'Return the maximum of three floating point values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Max3Float(ByVal rR As Double, ByVal rG As Double, ByVal rB As Double) As Double
    If (rR > rG) Then
        If (rR > rB) Then Max3Float = rR Else Max3Float = rB
    Else
        If (rB > rG) Then Max3Float = rB Else Max3Float = rG
    End If
End Function

'Return the maximum of three integer values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Max3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR > rG) Then
        If (rR > rB) Then Max3Int = rR Else Max3Int = rB
    Else
        If (rB > rG) Then Max3Int = rB Else Max3Int = rG
    End If
End Function

Public Function Min2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 < f2) Then Min2Float_Single = f1 Else Min2Float_Single = f2
End Function

Public Function Min2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 < l2) Then Min2Int = l1 Else Min2Int = l2
End Function

'Return the minimum of three floating point values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Min3Float(ByVal rR As Double, ByVal rG As Double, ByVal rB As Double) As Double
    If (rR < rG) Then
        If (rR < rB) Then Min3Float = rR Else Min3Float = rB
    Else
        If (rB < rG) Then Min3Float = rB Else Min3Float = rG
    End If
End Function

'Return the minimum of three integer values.  (PD commonly uses this for colors, hence the RGB notation.)
Public Function Min3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR < rG) Then
        If (rR < rB) Then Min3Int = rR Else Min3Int = rB
    Else
        If (rB < rG) Then Min3Int = rB Else Min3Int = rG
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
Public Function Modulo(ByVal quotient As Double, ByVal divisor As Double) As Double
    Modulo = quotient - Fix(quotient / divisor) * divisor
    If (Modulo < 0#) Then Modulo = Modulo + divisor
End Function

Public Function NearestInt(ByVal srcFloat As Single) As Long
    NearestInt = Int(srcFloat + 0.5!)
End Function

'Simplify an arbitrary polyline of arbitrary length using some arbitrary epsilon value (which defines
' minimum required distance between a point and the line defined by its neighbors).
'
'Pass your list of points and the number of points in the array.  (Upper array bound doesn't matter;
' it's ignored.)  This function will return the number of points removed; if it returns 0, no points
' were removed.  Also, the numOfPoints value - passed BYREF - will be updated to the current number
' of points in the final, simplified polyline array.  (Note that points beyond the final index of the
' simplified polyline *are not guaranteed to be zeroed-out*; their value is technically "undefined".)
'
'The strategy used is of my own invention.  I doubt I'm the first person to think of this approach,
' but I wanted something faster than the traditional Ramer–Douglas–Peucker algorithm (which is
' awkward to implement in VB6 since recursion is a non-starter, and stack conversions are cumbersome).
' My algorithm is O(n) and it requires no new allocations; the points are returned as-is in the
' source array, shifted as necessary to remove unimportant points.  It's very fast, with excellent
' accuracy, even on very gradual curves where traditional perpendicular-distance algorithms can fail -
' this is achieved by accumulating errors when removing points, and adding the accumulated error to
' the current point distance.  (The error tracker is reset when a point is *not* removed.)
'
'You can control the amount of errorFade with the same-named parameter; set the value to 0 to disable
' error diffusion entirely.
Public Function SimplifyLine(ByRef listOfPoints() As PointFloat, ByRef numOfPoints As Long, Optional ByVal epsilon As Single = 0.1!, Optional ByVal errorFade As Single = 0.25!) As Long
    
    'If we want to (possibly) remove points, we need at least three points to start!
    If (numOfPoints < 3) Then Exit Function
    Dim numPointsRemoved As Long
    
    'Square epsilon; this allows us to use a non-branching multiply instead of Abs() for comparisons
    epsilon = epsilon * epsilon
    
    'Start with the first line segment, comparing point (1) to the segment between (0) and (2)
    Dim leftIndex As Long, rightIndex As Long
    leftIndex = 0
    rightIndex = 2
    
    'Perpendicular distance to a given line-segment is used to determine removal, plus some temp variables
    ' to improve performance vs array accesses
    Dim curDistance As Single, x1 As Single, y1 As Single, x2 As Single, y2 As Single
    
    'Error diffusion is used to correct for gentle slopes in a uniform direction; we detect these
    ' via error accumulation, which automates the process of handling them.
    Dim curError As Single, origDistance As Single
    
    'Iterate all points except the endpoints (which are essential and non-removable)
    Dim i As Long
    For i = 1 To numOfPoints - 2
        
        'Compare the current point to the line running from startIndex to endIndex.
        ' (For improved performance, we manually in-line the perpendicular distance calculation.
        ' Note that we also do *not* apply absolute value until after the running error is updated.)
        y1 = (listOfPoints(rightIndex).y - listOfPoints(leftIndex).y)
        x1 = (listOfPoints(rightIndex).x - listOfPoints(leftIndex).x)
        curDistance = Sqr(y1 * y1 + x1 * x1)
        If (curDistance <> 0!) Then
            x1 = listOfPoints(leftIndex).x
            y1 = listOfPoints(leftIndex).y
            x2 = listOfPoints(rightIndex).x
            y2 = listOfPoints(rightIndex).y
            curDistance = ((y2 - y1) * listOfPoints(i).x - (x2 - x1) * listOfPoints(i).y + (x2 * y1) - (y2 * x1)) / curDistance
        End If
        
        'Make a note of the *unmodified* distance, then add the running error to the current distance
        origDistance = curDistance
        curDistance = curDistance + curError
        
        'Square distance, than compare to epsilon (fp multiply is faster than Abs() in VB)
        If (curDistance * curDistance < epsilon) Then
        
            'This point can be removed.  Increment the point removal counter, but otherwise do nothing;
            ' this point will be automatically "removed" by the left-shift code in the other branch.
            numPointsRemoved = numPointsRemoved + 1
            
            'Increment our running error (which is just the current perpendicular distance, multipled
            ' by a user-supplied fade value)
            curError = curError + (origDistance * errorFade)
            
        'This point cannot be removed.  Increment the *left* point index only, and shift the current
        ' point left-ward so that it's now located at the end of our running list of "good" points.
        ' (The shift step can obviously be skipped if no points have been removed yet.)  We also
        ' need to reset our running error whenever the current point is kept.
        Else
            leftIndex = leftIndex + 1
            If (numPointsRemoved > 0) Then listOfPoints(i - numPointsRemoved) = listOfPoints(i)
            curError = 0!
        End If
        
        'Right index is *always* incremented regardless of this point's removal status
        rightIndex = rightIndex + 1
        
    Next i
    
    'Shift the final polyline endpoint leftward by the number of removed points
    If (numPointsRemoved > 0) Then listOfPoints(numOfPoints - 1 - numPointsRemoved) = listOfPoints(numOfPoints - 1)
    
    'Return the number of points removed, and modify the current point count to reflect removals
    numOfPoints = numOfPoints - numPointsRemoved
    SimplifyLine = numPointsRemoved

End Function

'Simplify an arbitrary polyline of arbitrary length in preparation for UI display.  Points that are
' some amount closer together (user-specified, defaults to 1/10th of a pixel) will be merged to
' improve performance and display quality.
'
'Pass your list of points and the start index (probably 0) and end index (must be > 0) of the target line.
' (Actual upper array bound doesn't matter; it's ignored.)  This function will return a *new* end index
' via the ByRef endIndex value, and a Long indicating the number of points removed.  If 0 is returned,
' no points were removed.
'
'Note also points beyond the new final index of the simplified polyline *are not guaranteed to be zeroed-out*;
' their value is technically "undefined" and may retain the previous values you passed in.
Public Function SimplifyLineForScreen(ByRef listOfPoints() As PointFloat, ByVal startIndex As Long, ByRef endIndex As Long, Optional ByVal minDistance As Single = 0.1!) As Long
    
    Dim numOfPoints As Long
    numOfPoints = (endIndex - startIndex) + 1
    
    'If we want to (possibly) remove points, we need at least three points to start!
    If (numOfPoints < 3) Then Exit Function
    Dim numPointsRemoved As Long
    
    'Start with the first line segment, comparing point (1) to the segment between (0) and (2)
    Dim leftIndex As Long
    leftIndex = startIndex
    
    'Direct distance between two points is used to determine removal, plus some temp variables
    ' to improve performance vs array accesses.
    Dim curDistance As Single, x1 As Single, y1 As Single
    
    'Iterate all points except the endpoints (which are essential and non-removable)
    Dim i As Long
    For i = startIndex + 1 To endIndex - 1
        
        'Calculate distance between the current point and the previous point.
        ' (For improved performance, we manually in-line the distance calculation.)
        x1 = (listOfPoints(i).x - listOfPoints(leftIndex).x)
        y1 = (listOfPoints(i).y - listOfPoints(leftIndex).y)
        curDistance = Sqr(x1 * x1 + y1 * y1)
        
        'Perform removal check
        If (curDistance < minDistance) Then
        
            'This point can be removed.  Increment the point removal counter, but otherwise do nothing;
            ' this point will be automatically "removed" by the left-shift code in the other branch.
            numPointsRemoved = numPointsRemoved + 1
            
        'This point cannot be removed.  Increment the *left* point index only, and shift the current
        ' point left-ward so that it's now located at the end of our running list of "good" points.
        ' (The shift step can obviously be skipped if no points have been removed yet.)
        Else
            leftIndex = leftIndex + 1
            If (numPointsRemoved > 0) Then listOfPoints(i - numPointsRemoved) = listOfPoints(i)
        End If
        
    Next i
    
    'Shift the final polyline endpoint leftward by the number of removed points
    If (numPointsRemoved > 0) Then
        listOfPoints(endIndex - numPointsRemoved) = listOfPoints(endIndex)
        endIndex = endIndex - numPointsRemoved
    End If
    
    'Return the number of points removed
    SimplifyLineForScreen = numPointsRemoved
    
End Function

'Simplify an arbitrary polyline of arbitrary length that was produced by PD's marching-squares algorithms.
' This is particularly relevant for the selection tool, which is 100% mask-based, so we need to manually
' generate vector boundaries for the mask under a variety of circumstances.
'
'Pass your list of points and the start index (probably 0) and end index (must be > 0) of the target line.
' (Actual upper array bound doesn't matter; it's ignored.)  This function will return a *new* end index
' via the ByRef endIndex value, and a Long indicating the number of points removed.  If 0 is returned,
' no points were removed.
'
'Note also points beyond the new final index of the simplified polyline *are not guaranteed to be zeroed-out*;
' their value is technically "undefined" and may retain the previous values you passed in.
'
'A potentially minor but relevant detail is that this function assumes path points are in clock-wise order.
' It doesn't really matter, but it slightly affects whether antialiasing is performed 0.5px "outside" or
' "inside" the border of the shape.  If a closed path is in clockwise order, the antialiasing will occur
' 0.5px "outside the shape"; for counter-clockwise order, it will be "inside" the shape's border.
Public Function SimplifyLinesFromMarchingSquares(ByRef listOfPoints() As PointFloat, ByVal startIndex As Long, ByRef endIndex As Long) As Long
    
    Const REPORT_SIMPLIFICATION_STATS As Boolean = False
    
    Dim numOfPoints As Long
    numOfPoints = (endIndex - startIndex) + 1
    If REPORT_SIMPLIFICATION_STATS Then PDDebug.LogAction "Starting with " & numOfPoints
    
    'If we want to (possibly) remove points, we need at least three points to start!
    If (numOfPoints < 3) Then Exit Function
    Dim numPointsRemoved As Long
    
    'Because paths may contain multiple discrete polylines, we allow the caller to pass whatever
    ' starting index they want.
    Dim leftIndex As Long
    leftIndex = startIndex
    
    'This marching squares simplifier is my own design.  It will only produce ideal results on polylines
    ' that come directly from PD's marching squares implementation (but will likely work well on others,
    ' pending additional testing).  What matters is that the passed list of points consists *ONLY* of
    ' horizontal and vertical line segments.  Other line segments *WILL NOT WORK*.
    '
    'What we want to do in this first pass is identify three-point junctions that can be better be
    ' represented as a single two-point line.  This is true for polylines that move a single pixel over,
    ' then change direction and travel some arbitrary distance before repeating moving a single pixel
    ' over *in the same direction as before*.  These types of blocky lines can be converted to smooth
    ' lines that antialias much better while also having much lower net point counts.
    '
    'To do this, we want to build a list of point "types".  This makes it easy to identify which points
    ' can be removed and which points must remain.
    Dim pointTypes() As PD_PointSimplify
    ReDim pointTypes(startIndex To endIndex) As PD_PointSimplify
    
    Dim i As Long
    For i = startIndex To endIndex - 1
        
        'Assume most points are not removeable; we'll update this value after further checks
        pointTypes(i) = ps_Essential
        
        'Determine the direction of the line coming *into* this point.
        If (i > startIndex) Then
            If (listOfPoints(i).x = listOfPoints(i - 1).x) Then
                If (listOfPoints(i).y > listOfPoints(i - 1).y) Then
                    pointTypes(i) = ps_South
                ElseIf (listOfPoints(i).y < listOfPoints(i - 1).y) Then
                    pointTypes(i) = ps_North
                End If
            ElseIf (listOfPoints(i).y = listOfPoints(i - 1).y) Then
                If (listOfPoints(i).x > listOfPoints(i - 1).x) Then
                    pointTypes(i) = ps_East
                Else
                    pointTypes(i) = ps_West
                End If
            End If
        End If
        
        '(For the first point, assume a 90 turn coming in; this allows for smoothing that point.)
        If (i = startIndex + 1) Then
            If (pointTypes(i) = ps_East) Then
                pointTypes(i - 1) = ps_North
            ElseIf (pointTypes(i) = ps_South) Then
                pointTypes(i - 1) = ps_East
            ElseIf (pointTypes(i) = ps_West) Then
                pointTypes(i - 1) = ps_South
            Else
                pointTypes(i - 1) = ps_West
            End If
        End If
        
        'To determine if a point is removable, we need to compare two things:
        ' 1) Does this point [p0] and the point [p-2] have the *same* directionality?
        ' 2) If (1), are [p0] and [p-2] separated by only one pixel in the x or y direction?
        '
        'If (1) and (2) are both true, the point between them can be removed, and a straight line
        ' from [p0] to [p2] used in its place.
        If (i > startIndex + 1) Then
            
            'Compare directionality
            If (pointTypes(i) = pointTypes(i - 2)) Then
                
                'Look for a difference of *1* in the x or y direction
                If (Abs(listOfPoints(i).x - listOfPoints(i - 2).x) = 1) Then
                    pointTypes(i - 1) = ps_Remove
                ElseIf (Abs(listOfPoints(i).y - listOfPoints(i - 2).y) = 1) Then
                    pointTypes(i - 1) = ps_Remove
                End If
                
            End If
            
        End If
            
    Next i
    
    'With all potentially removeable points flagged, we can now perform removal in a single sweep
    numPointsRemoved = 0
    For i = startIndex To endIndex
        
        'Perform removal check
        If (pointTypes(i) = ps_Remove) Then
        
            'This point can be removed.  Increment the point removal counter, but otherwise do nothing;
            ' this point will be automatically "removed" by the left-shift code in the other branch.
            numPointsRemoved = numPointsRemoved + 1
            
        'This point cannot be removed.  Increment the *left* point index only, and shift the current
        ' point left-ward so that it's now located at the end of our running list of "good" points.
        ' (The shift step can obviously be skipped if no points have been removed yet.)
        Else
            leftIndex = leftIndex + 1
            If (numPointsRemoved > 0) Then listOfPoints(i - numPointsRemoved) = listOfPoints(i)
        End If
        
    Next i
    
    'Shift the final polyline endpoint leftward by the number of removed points
    If (numPointsRemoved > 0) Then
        listOfPoints(endIndex - numPointsRemoved) = listOfPoints(endIndex)
        endIndex = endIndex - numPointsRemoved
    End If
    
    If REPORT_SIMPLIFICATION_STATS Then PDDebug.LogAction "Removed " & numPointsRemoved & " in first pass"
    
    'To exit now, uncomment the following two lines:
    'SimplifyLinesFromMarchingSquares = numPointsRemoved
    'Exit Function
    
    'We now have new starting + ending indices, and all "useless" points have been removed.
    ' But that's not all we can do!  Removing points in the previous step likely left us with many points
    ' that lie on the same line as their neighbors.  Why not remove those redundant points while we're here?
    Dim numPointsRemoved2 As Long
    For i = startIndex + 1 To endIndex - 1
        
        'Because the previous step was only capable of converting an e.g. 1-px vertical line and neighboring
        ' horizontal line into a single slightly-sloped line, we know that neighboring points in this algorithm
        ' will only lie on the same line under certain circumstances.  This greatly simplifies the way we
        ' check for "is point on line".
        If (listOfPoints(i).x - listOfPoints(i - 1).x) = (listOfPoints(i + 1).x - listOfPoints(i).x) Then
            If (listOfPoints(i).y - listOfPoints(i - 1).y) = (listOfPoints(i + 1).y - listOfPoints(i).y) Then
                
                'This point can be removed.  Increment the point removal counter, but otherwise do nothing;
                ' this point will be automatically "removed" by the left-shift code in the other branches.
                numPointsRemoved2 = numPointsRemoved2 + 1
            
            'This point cannot be removed.  Increment the *left* point index only, and shift the current
            ' point left-ward so that it's now located at the end of our running list of "good" points.
            ' (The shift step can obviously be skipped if no points have been removed yet.)
            Else
                leftIndex = leftIndex + 1
                If (numPointsRemoved2 > 0) Then listOfPoints(i - numPointsRemoved2) = listOfPoints(i)
            End If
            
        'Same as branch above
        Else
            leftIndex = leftIndex + 1
            If (numPointsRemoved2 > 0) Then listOfPoints(i - numPointsRemoved2) = listOfPoints(i)
        End If
        
    Next i
    
    'Once again shift the final polyline endpoint leftward by the number of removed points
    If (numPointsRemoved2 > 0) Then
        listOfPoints(endIndex - numPointsRemoved2) = listOfPoints(endIndex)
        endIndex = endIndex - numPointsRemoved2
    End If
    
    If REPORT_SIMPLIFICATION_STATS Then
        PDDebug.LogAction "Removed " & numPointsRemoved2 & " in second pass"
        PDDebug.LogAction "Final point count is " & (endIndex - startIndex) + 1
    End If
    
    'Return the number of points removed from *both* passes
    SimplifyLinesFromMarchingSquares = numPointsRemoved + numPointsRemoved2
    
End Function

'Given a pd2DPath object, simplify the path for display on-screen.  This function automatically removes any points
' in the path closer than [minDistance].
'
'I recommend transforming your source path into screen coordinates, then handing it off to this function.
' Resulting GDI+ drawing will often be much faster as a result, and the quality will be improved from not
' needing to calculate complex miters on so many subpixel junctions.
Public Function SimplifyPathForScreen(ByRef srcPath As pd2DPath, Optional ByVal minDistance As Single = 0.1!) As Boolean
    
    SimplifyPathForScreen = False
    If (srcPath Is Nothing) Then Exit Function
    
    'Start by converting the source path to a list of discrete points and subpaths
    Dim listOfPoints() As PointFloat, numOfPoints As Long, listOfSubpaths() As Long, listOfClosedStates() As Boolean, numOfSubpaths As Long
    SimplifyPathForScreen = srcPath.GetPathPoints(listOfPoints, numOfPoints, listOfSubpaths, listOfClosedStates, numOfSubpaths)
    If (Not SimplifyPathForScreen) Then Exit Function
    
    'The source path has given us everything we need.  Reset it to null-path state.
    srcPath.ResetPath
    
    'We can now hand off each individual sub-path to the standalone "simplify for screen" function.
    Dim idxFirst As Long, idxLast As Long, numPointsRemoved As Long, numPointsRemovedTotal As Long
    Dim i As Long, j As Long
    
    'If you want to skip simplification and simply test correctness of retrieving points from a GDI+ path,
    ' then manually adding them back to a path (and getting an identical result), use this (ugh) GOTO:
    'GoTo SkipSimplification
    
    For i = 0 To numOfSubpaths - 1
        
        'Retrieve first/last indices for the current sub-path
        idxFirst = listOfSubpaths(i)
        If (i < numOfSubpaths - 1) Then
            idxLast = listOfSubpaths(i + 1) - 1
        
        'numOfPoints is updated "as-we-go" and thus will already account for any removed point counts
        Else
            idxLast = numOfPoints - 1
        End If
        
        'Simplify this line segment, and note that the simplifier returns the number of points removed (if any).
        ' This is important because we need to manually shift subsequent points forward.
        numPointsRemoved = PDMath.SimplifyLineForScreen(listOfPoints, idxFirst, idxLast, minDistance)
        
        'BEFORE we update the total number of points removed, shift the current list of points forward
        ' in the array.  (Any points we removed from this path won't matter until the *next* path.)
        If (numPointsRemovedTotal > 0) Then
            For j = idxFirst To idxLast
                listOfPoints(j - numPointsRemovedTotal) = listOfPoints(j)
            Next j
            listOfSubpaths(i) = listOfSubpaths(i) - numPointsRemovedTotal
        End If
        
        'Now we can update the total removed points count
        numPointsRemovedTotal = numPointsRemovedTotal + numPointsRemoved
        
    Next i
    
    'Update the final points count
    numOfPoints = numOfPoints - numPointsRemovedTotal
    
'When debugging, I verified the correctness of this function by completely skipping the previous block
' (the actual simplification) and simply reconstituting the path manually.  If you doubt my work, feel free
' to confirm the same way!
SkipSimplification:
    
    'All sub-paths have now been trimmed down to screen size.  Reconstruct the original path
    ' from the new sub-path list.
    For i = 0 To numOfSubpaths - 1
        
        idxFirst = listOfSubpaths(i)
        If (i < numOfSubpaths - 1) Then
            idxLast = listOfSubpaths(i + 1) - 1
        
        'numOfPoints is updated "as-we-go" and thus will already account for any removed point counts
        Else
            idxLast = numOfPoints - 1
        End If
        
        srcPath.AddPolygon (idxLast - idxFirst) + 1, VarPtr(listOfPoints(idxFirst)), listOfClosedStates(i), False
        
    Next i
    
    SimplifyPathForScreen = True
    
End Function

'Given a pd2DPath object generated by PD's marching-squares algorithm, simplify the path and prepare it for
' antialiased handling.
Public Function SimplifyPathFromMarchingSquares(ByRef srcPath As pd2DPath) As Boolean
    
    SimplifyPathFromMarchingSquares = False
    If (srcPath Is Nothing) Then Exit Function
    
    'Start by converting the source path to a list of discrete points and subpaths
    Dim listOfPoints() As PointFloat, numOfPoints As Long, listOfSubpaths() As Long, listOfClosedStates() As Boolean, numOfSubpaths As Long
    SimplifyPathFromMarchingSquares = srcPath.GetPathPoints(listOfPoints, numOfPoints, listOfSubpaths, listOfClosedStates, numOfSubpaths)
    If (Not SimplifyPathFromMarchingSquares) Then Exit Function
    
    'The source path has given us everything we need.  Reset it to null-path state...
    ' (April 2024) BUT also ensure we preserve fill rule (which is odd-even for outlines generated
    ' from edge-tracing, like PD's magic wand uses!)
    Dim origFillRule As PD_2D_FillRule
    origFillRule = srcPath.GetFillRule()
    srcPath.ResetPath
    srcPath.SetFillRule origFillRule
    
    'We can now hand off each individual sub-path to the standalone "simplify for screen" function.
    Dim idxFirst As Long, idxLast As Long, numPointsRemoved As Long, numPointsRemovedTotal As Long
    Dim i As Long, j As Long
    
    'If you want to skip simplification and simply test correctness of retrieving points from a GDI+ path,
    ' then manually adding them back to a path (and getting an identical result), use this (ugh) GOTO:
    'GoTo SkipSimplification
    
    For i = 0 To numOfSubpaths - 1
        
        'Retrieve first/last indices for the current sub-path
        idxFirst = listOfSubpaths(i)
        If (i < numOfSubpaths - 1) Then
            idxLast = listOfSubpaths(i + 1) - 1
            
        'numOfPoints is updated "as-we-go" and thus will already account for any removed point counts
        Else
            idxLast = numOfPoints - 1
        End If
        
        'Simplify this line segment, and note that the simplifier returns the number of points removed (if any).
        ' This is important because we need to manually shift subsequent points forward.
        numPointsRemoved = PDMath.SimplifyLinesFromMarchingSquares(listOfPoints, idxFirst, idxLast)
        
        'BEFORE we update the total number of points removed, shift the current list of points forward
        ' in the array.  (Any points we removed from this path won't matter until the *next* path.)
        If (numPointsRemovedTotal > 0) Then
            For j = idxFirst To idxLast
                listOfPoints(j - numPointsRemovedTotal) = listOfPoints(j)
            Next j
            listOfSubpaths(i) = listOfSubpaths(i) - numPointsRemovedTotal
        End If
        
        'Now we can update the total removed points count
        numPointsRemovedTotal = numPointsRemovedTotal + numPointsRemoved
        
    Next i
    
    'Update the final points count
    numOfPoints = numOfPoints - numPointsRemovedTotal
    
'When debugging, I verified the correctness of this function by completely skipping the previous block
' (the actual simplification) and simply reconstituting the path manually.  If you doubt my work, feel free
' to confirm the same way!
SkipSimplification:
    
    'All sub-paths have now been trimmed down to screen size.  Reconstruct the original path
    ' from the new sub-path list.
    For i = 0 To numOfSubpaths - 1
        
        idxFirst = listOfSubpaths(i)
        If (i < numOfSubpaths - 1) Then
            idxLast = listOfSubpaths(i + 1) - 1
        
        'numOfPoints is updated "as-we-go" and thus will already account for any removed point counts
        Else
            idxLast = numOfPoints - 1
        End If
        
        srcPath.AddPolygon (idxLast - idxFirst) + 1, VarPtr(listOfPoints(idxFirst)), listOfClosedStates(i), False
        
    Next i
    
    SimplifyPathFromMarchingSquares = True
    
End Function

'Use a simple moving-average formula to smooth a given input line on the Y-axis only.
' Strength is a value on the range [0, 1]; 0 is a nop, 1 replaces all points with their moving average
Public Sub SmoothLineY(ByRef listOfPoints() As PointFloat, ByRef numOfPoints As Long, Optional ByVal strength As Single = 0.5!)
    
    'If we want to (possibly) remove points, we need at least three points to start!
    If (numOfPoints < 3) Then Exit Sub
    
    'A temporary copy of the input points are required so we don't lose data.
    ' (This could be worked-around with clever caching, but PD's input lists are
    ' generally small so a full copy is easier.)
    Dim copyOfPoints() As PointFloat
    ReDim copyOfPoints(0 To numOfPoints - 1) As PointFloat
    CopyMemoryStrict VarPtr(copyOfPoints(0)), VarPtr(listOfPoints(0)), 8 * numOfPoints
    
    If (strength < 0!) Then strength = 0!
    If (strength > 1!) Then strength = 1!
    
    Dim invStrength As Single
    invStrength = 1! - strength
    
    Dim i As Long
    For i = 1 To numOfPoints - 2
        
        'Calculate an average y value
        Dim newY As Single
        newY = (copyOfPoints(i - 1).y + copyOfPoints(i).y + copyOfPoints(i + 1).y) * 0.3333333!
        
        'Average using the "strength" parameter
        listOfPoints(i).y = (listOfPoints(i).y * invStrength) + (newY * strength)
        
    Next i

End Sub

'Given an array of points (in floating-point format), calculate the center point of the bounding rect.
Public Sub FindCenterOfFloatPoints(ByRef dstPoint As PointFloat, ByRef srcPoints() As PointFloat)

    Dim curMinX As Single, curMinY As Single, curMaxX As Single, curMaxY As Single
    curMinX = 9999999!:   curMaxX = -9999999!:   curMinY = 9999999!:   curMaxY = -9999999!
    
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
    x1 = -srcWidth / 2#
    x2 = srcWidth / 2#
    x3 = srcWidth / 2#
    x4 = -srcWidth / 2#
    y1 = srcHeight / 2#
    y2 = srcHeight / 2#
    y3 = -srcHeight / 2#
    y4 = -srcHeight / 2#

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
Public Sub FindCornersOfRotatedRect(ByVal srcWidth As Double, ByVal srcHeight As Double, ByVal rotateAngle As Double, ByRef dstPoints() As PointFloat, Optional ByVal arrayAlreadyDimmed As Boolean = False)

    'Convert the rotation angle to radians
    rotateAngle = rotateAngle * PI_DIV_180
    
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
    halfWidth = srcWidth / 2#
    halfHeight = srcHeight / 2#
    
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
    If (Not arrayAlreadyDimmed) Then ReDim dstPoints(0 To 3) As PointFloat
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
Public Sub RotatePointAroundPoint(ByVal rotateX As Single, ByVal rotateY As Single, ByVal centerX As Single, ByVal centerY As Single, ByVal angleInRadians As Single, ByRef newX As Single, ByRef newY As Single)

    'For performance reasons, it's easier to cache the cos and sin of the angle in question
    Dim sinAngle As Double, cosAngle As Double
    sinAngle = Sin(angleInRadians)
    cosAngle = Cos(angleInRadians)
    
    'Rotation works the same way as it does around (0, 0), except we transform the center position before and
    ' after rotation.
    newX = cosAngle * (rotateX - centerX) - sinAngle * (rotateY - centerY) + centerX
    newY = cosAngle * (rotateY - centerY) + sinAngle * (rotateX - centerX) + centerY
    
End Sub

'Simple helper if you don't want to manually initialize a pdRandomize instance; *do not use* in perf-sensitive code
Public Function GetCompletelyRandomInt() As Long
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_AutomaticAndRandom
    GetCompletelyRandomInt = cRandom.GetRandomInt_WH()
End Function

'Given a RectF object, enlarge the boundaries to produce an integer-only RectF that is guaranteed
' to encompass the entire original rect.  (Said another way, the modified rect will *never* be smaller
' than the original rect.)
Public Sub GetIntClampedRectF(ByRef srcRectF As RectF)
    Dim xOffset As Single, yOffset As Single
    xOffset = srcRectF.Left - Int(srcRectF.Left)
    yOffset = srcRectF.Top - Int(srcRectF.Top)
    srcRectF.Left = Int(srcRectF.Left)
    srcRectF.Top = Int(srcRectF.Top)
    srcRectF.Width = Int(srcRectF.Width + xOffset + 0.9999!)
    srcRectF.Height = Int(srcRectF.Height + yOffset + 0.9999!)
End Sub

'Note that GDI rects (and possibly others) have strict requirements about the way right/bottom coords are defined,
' so these convenience functions may need additional tweaking by the caller if forwarding the rect to an external library.
Public Sub GetRectFRB_FromRectF(ByRef srcRectF As RectF, ByRef dstRectF_RB As RectF_RB)
    
    With dstRectF_RB
        .Left = srcRectF.Left
        .Top = srcRectF.Top
        .Right = srcRectF.Left + srcRectF.Width
        .Bottom = srcRectF.Top + srcRectF.Height
    End With
    
End Sub

Public Sub GetRectF_FromRectFRB(ByRef srcRectF_RB As RectF_RB, ByRef dstRectF As RectF)
    
    With dstRectF
        .Left = srcRectF_RB.Left
        .Top = srcRectF_RB.Top
        .Width = srcRectF_RB.Right - srcRectF_RB.Left
        .Height = srcRectF_RB.Bottom - srcRectF_RB.Top
    End With
    
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

'Cheap and easy way to find the nearest power of two.  Note that you could also do this with logarithms
' (or even bitshifts maybe?) but I haven't thought about it hard enough lol
Public Function NearestPowerOfTwo(ByVal srcNumber As Long) As Long
    
    Dim curPower As Long
    curPower = 1
    
    Do While (curPower < srcNumber)
        curPower = curPower * 2
    Loop
    
    NearestPowerOfTwo = curPower
    
End Function

'Rational erf approximation.  Adapted from the public domain "Handbook of Mathematical Functions", algorithm 7.1.26
' (link good as of September '18: http://people.math.sfu.ca/~cbm/aands/frameindex.htm).  Technically the accuracy is only
' appropriate for Singles (e(x) <= 1.5e-7), but input and output are Double due to its prevalence in PD calculations.
Public Function ERF(ByVal x As Double) As Double
    
    'Cache the sign in advance
    Dim initSgn As Double
    initSgn = Sgn(x)
    
    x = Abs(x)
    
    Dim t As Double
    t = 1# / (1# + 0.3275911 * x)
    
    Dim y As Double
    y = 1# - (((((1.061405429 * t + -1.453152027) * t) + 1.421413741) * t + -0.284496736) * t + 0.254829592) * t * Exp(-(x * x))
    
    ERF = initSgn * y
    
End Function

Public Function ERFC(ByVal x As Double) As Double
    ERFC = ERF(1# - x)
End Function

'Inverse erf() function, as estimated by Sergei Winitzki via https://en.wikipedia.org/wiki/Error_function.
' (Specific source at the time of this writing was http://sites.google.com/site/winitzki/sergei-winitzkis-files/erf-approx.pdf)
Public Function ERF_Inv(ByVal x As Double) As Double
    
    Dim initSgn As Double
    initSgn = Sgn(x)
    
    Dim a As Double
    a = (8# / (3 * PI)) * ((PI - 3#) / (4# - PI))
    
    Dim t As Double
    t = 2# / (PI * a) + Log(1# - x * x) * 0.5
    t = t * t - Log(1# - x * x) / a
    t = Sqr(t) - (2# / (PI * a) + Log(1# - x * x) * 0.5)
    ERF_Inv = initSgn * Sqr(t)
    
End Function

Public Function ERFC_Inv(ByVal x As Double) As Double
    ERFC_Inv = ERF_Inv(1# - x)
End Function
