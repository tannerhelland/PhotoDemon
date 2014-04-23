Attribute VB_Name = "Math_Functions"
'***************************************************************************
'Specialized Math Routines
'Copyright ©2013-2014 by Tanner Helland
'Created: 13/June/13
'Last updated: 13/June/13
'Last update: added a function by VB6 coder LaVolpe that converts decimals to fractions.  PhotoDemon uses this function to
'             approximate image aspect ratios from width/height values.
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June 2013.  This module is now used to store all the random bits of specialized math required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'Convert a decimal to a near-identical fraction using vector math.
' This excellent function comes courtesy of VB6 coder LaVolpe.  I have modified it slightly to suit PhotoDemon's unique needs.
' You can download the original at this link (good as of 13 June 2014): http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=61596&lngWId=1
Public Sub convertToFraction(ByVal v As Double, w As Double, n As Double, d As Double, Optional ByVal maxDenomDigits As Byte, Optional ByVal Accuracy As Double = 100#)

    Const MaxTerms As Integer = 50          'Limit to prevent infinite loop
    Const MinDivisor As Double = 1E-16      'Limit to prevent divide by zero
    Const MaxError As Double = 1E-50        'How close is enough
    Dim f As Double                         'Fraction being converted
    Dim a As Double     'Current term in continued fraction
    Dim N1 As Double    'Numerator, denominator of last approx
    Dim D1 As Double
    Dim N2 As Double    'Numerator, denominator of previous approx
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
    v = CDbl(Mid$(Str(Abs(v)), Len(Str(w))))
    
    ' check for no decimal or zero
    If v = 0 Then GoTo RtnResult
    
    f = v                       'Initialize fraction being converted
    
    N1 = 1                      'Initialize fractions with 1/0, 0/1
    D1 = 0
    N2 = 0
    D2 = 1

    On Error GoTo RtnResult
    For i = 0 To MaxTerms
        a = Fix(f)              'Get next term
        f = f - a               'Get new divisor
        n = N1 * a + N2         'Calculate new fraction
        d = D1 * a + D2
        N2 = N1                 'Save last two fractions
        D2 = D1
        N1 = n
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
            n = N1
            d = D1
        Else
            d = D2
            n = N2
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
    srcAspect = srcWidth / srcHeight
    dstAspect = dstWidth / dstHeight
    
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

'Return the distance between two 3D points
Public Function distanceThreeDimensions(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    distanceThreeDimensions = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
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
