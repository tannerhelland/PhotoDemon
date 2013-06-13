Attribute VB_Name = "Math_Functions"
'***************************************************************************
'Specialized Math Routines
'Copyright ©2012-2013 by Tanner Helland
'Created: 13/June/13
'Last updated: 13/June/13
'Last update: added a function by VB6 code LaVolpe for converting decimals to fractions.  Used to display image aspect ratio.
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June 2013.  This module is now used to store all the random bits of specialized math required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit


'Convert a decimal to a near-identical fraction using vector math.
' This excellent function comes courtesy of VB6 coder LaVolpe.  I have modified it slightly to suit PhotoDemon's unique needs.
' You can download the original at this link (good as of 13 June 2013): http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=61596&lngWId=1
Public Sub ConvertToFraction(ByVal V As Double, W As Double, N As Double, D As Double, Optional ByVal maxDenomDigits As Byte, Optional ByVal Accuracy As Double = 100#)

    Const MaxTerms As Integer = 50          'Limit to prevent infinite loop
    Const MinDivisor As Double = 1E-16      'Limit to prevent divide by zero
    Const MaxError As Double = 1E-50        'How close is enough
    Dim F As Double                         'Fraction being converted
    Dim A As Double     'Current term in continued fraction
    Dim N1 As Double    'Numerator, denominator of last approx
    Dim D1 As Double
    Dim N2 As Double    'Numerator, denominator of previous approx
    Dim D2 As Double
    Dim I As Integer
    Dim T As Double
    Dim maxDenom As Double
    Dim bIsNegative As Boolean
    Dim sDec As String
    
    If maxDenomDigits = 0 Or maxDenomDigits > 17 Then maxDenomDigits = 17
    maxDenom = 10 ^ maxDenomDigits
    If Accuracy > 100 Or Accuracy < 1 Then Accuracy = 100
    Accuracy = Accuracy / 100#
    
    bIsNegative = (V < 0)
    W = Abs(Fix(V))
    'V = Abs(V) - W << subtracting doubles can change the decimal portion by adding more numeral at end
    'TANNER'S NOTE: the original version of this included +1 to the string length, which gave me consistent errors.  So I removed it.
    V = CDbl(Mid$(CStr(Abs(V)), Len(CStr(W))))
    
    ' check for no decimal or zero
    If V = 0 Then GoTo RtnResult
    
    F = V                       'Initialize fraction being converted
    
    N1 = 1                      'Initialize fractions with 1/0, 0/1
    D1 = 0
    N2 = 0
    D2 = 1

    On Error GoTo RtnResult
    For I = 0 To MaxTerms
        A = Fix(F)              'Get next term
        F = F - A               'Get new divisor
        N = N1 * A + N2         'Calculate new fraction
        D = D1 * A + D2
        N2 = N1                 'Save last two fractions
        D2 = D1
        N1 = N
        D1 = D
                                'Quit if dividing by zero
        If F < MinDivisor Then Exit For

                                'Quit if close enough
        T = N / D               ' A=zero indicates exact match or extremely close
        A = Abs(V - T)          ' Difference btwn actual V and calculated T
        If A < MaxError Then Exit For
                                'Quit if max denominator digits encountered
        If D > maxDenom Then Exit For
                                ' Quit if preferred accuracy accomplished
        If N Then
            If T > V Then T = V / T Else T = T / V
            If T >= Accuracy And Abs(T) < 1 Then Exit For
        End If
        F = 1# / F               'Take reciprocal
    Next I

RtnResult:
    If Err Or D > maxDenom Then
        ' in above case, use the previous best N & D
        If D2 = 0 Then
            N = N1
            D = D1
        Else
            D = D2
            N = N2
        End If
    End If
    
    ' correct for negative values
    If bIsNegative Then
        If W Then W = -W Else N = -N
    End If
    
    'TANNER'S NOTE: the original function included some simple code here to generate a user-friendly string of the result.
    ' PhotoDemon does this itself (so translations can be applied) so I have removed that section of code.

End Sub

'Convert a width and height pair to a new max width and height, while preserving aspect ratio
Public Sub convertAspectRatio(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long)
    
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = srcWidth / srcHeight
    dstAspect = dstWidth / dstHeight
    
    If srcAspect > dstAspect Then
        newWidth = dstWidth
        newHeight = CSng(srcHeight / srcWidth) * newWidth + 0.5
    Else
        newHeight = dstHeight
        newWidth = CSng(srcWidth / srcHeight) * newHeight + 0.5
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

