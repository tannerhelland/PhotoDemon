Attribute VB_Name = "ComplexNumbers"
'***************************************************************************
'Complex Number math and interfaces
'Copyright 2019-2026 by Tanner Helland
'Created: 24/August/21
'Last updated: 25/August/21
'Last update: add Float versions; in PD, the precision vs performance trade-off often tilts toward perf
'
'Complex numbers are weird.  If you haven't had the (mis)fortune of studying these, Wikipedia has you covered:
' https://en.wikipedia.org/wiki/Complex_number
'
'I originally added most of these functions to PhotoDemon for a recursive gaussian approximation function by
' Pascal Getreuer in "A Survey of Gaussian Convolution Algorithms", http://www.ipol.im/pub/art/2013/87/
' Important copyright and license information regarding Pascal's work:
' - Original C version is copyright (c) 2012-2013, Pascal Getreuer <getreuer@cmla.ens-cachan.fr>
' - Used here under its original BSD license <http://www.opensource.org/licenses/bsd-license.html>
' - Translated into VB6 by Tanner Helland in 2019
'
'These complex number functions are now used in other capacities (e.g. Effects > Droste) so please take care
' if modifying them, as they affect more places than you might think.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Type ComplexNumber
    c_real As Double
    c_imag As Double
End Type

Public Type ComplexNumberF
    c_real As Single
    c_imag As Single
End Type

'Helper construction function
Public Function make_complex(ByVal a As Double, ByVal b As Double) As ComplexNumber
    make_complex.c_real = a
    make_complex.c_imag = b
End Function

'Complex conjugate
Public Function c_conj(ByRef z As ComplexNumber) As ComplexNumber
    c_conj.c_real = z.c_real
    c_conj.c_imag = -z.c_imag
End Function

'Complex addition
Public Function c_add(ByRef w As ComplexNumber, ByRef z As ComplexNumber) As ComplexNumber
    c_add.c_real = w.c_real + z.c_real
    c_add.c_imag = w.c_imag + z.c_imag
End Function

'Complex negation
Public Function c_neg(ByRef z As ComplexNumber) As ComplexNumber
    c_neg.c_real = -z.c_real
    c_neg.c_imag = -z.c_imag
End Function

'Complex subtraction
Public Function c_sub(ByRef w As ComplexNumber, ByRef z As ComplexNumber) As ComplexNumber
    c_sub.c_real = w.c_real - z.c_real
    c_sub.c_imag = w.c_imag - z.c_imag
End Function

'Complex multiplication
Public Function c_mul(ByRef w As ComplexNumber, ByRef z As ComplexNumber) As ComplexNumber
    c_mul.c_real = w.c_real * z.c_real - w.c_imag * z.c_imag
    c_mul.c_imag = w.c_real * z.c_imag + w.c_imag * z.c_real
End Function

'Complex multiplicative inverse 1/z
Public Function c_inv(ByRef z As ComplexNumber) As ComplexNumber

    'There are two mathematically-equivalent formulas for the inverse. For accuracy,
    ' choose the formula with the smaller value of |ratio|.
    Dim ratio As Double, denom As Double
    If (Abs(z.c_real) >= Abs(z.c_imag)) Then
        ratio = z.c_imag / z.c_real
        denom = z.c_real + z.c_imag * ratio
        c_inv.c_real = 1# / denom
        c_inv.c_imag = -ratio / denom
    Else
        ratio = z.c_real / z.c_imag
        denom = z.c_real * ratio + z.c_imag
        c_inv.c_real = ratio / denom
        c_inv.c_imag = -1# / denom
    End If
    
End Function

'Complex division w/z
Public Function c_div(ByRef w As ComplexNumber, ByRef z As ComplexNumber) As ComplexNumber

    'For accuracy, choose the formula with the smaller value of |ratio|.
    Dim ratio As Double, denom As Double
    If (Abs(z.c_real) >= Abs(z.c_imag)) Then
        ratio = z.c_imag / z.c_real
        denom = 1# / (z.c_real + z.c_imag * ratio)
        c_div.c_real = (w.c_real + w.c_imag * ratio) * denom
        c_div.c_imag = (w.c_imag - w.c_real * ratio) * denom
    Else
        ratio = z.c_real / z.c_imag
        denom = 1# / (z.c_real * ratio + z.c_imag)
        c_div.c_real = (w.c_real * ratio + w.c_imag) * denom
        c_div.c_imag = (w.c_imag * ratio - w.c_real) * denom
    End If
    
End Function

'Complex magnitude
Public Function c_mag(ByRef z As ComplexNumber) As Double
    
    Dim tmpz As ComplexNumber
    tmpz.c_real = Abs(z.c_real)
    tmpz.c_imag = Abs(z.c_imag)
    
    'For accuracy, choose the formula with the smaller value of |ratio|.
    Dim ratio As Double
    If (tmpz.c_real >= tmpz.c_imag) Then
        If (tmpz.c_real <> 0#) Then ratio = tmpz.c_imag / tmpz.c_real Else ratio = 0#
        c_mag = tmpz.c_real * Sqr(1# + ratio * ratio)
    Else
        If (tmpz.c_imag <> 0#) Then ratio = ratio = tmpz.c_real / tmpz.c_imag Else ratio = 0#
        c_mag = tmpz.c_imag * Sqr(1# + ratio * ratio)
    End If
    
End Function

'Complex argument (angle) in [-pi,+pi]
Public Function c_arg(ByRef z As ComplexNumber) As Double
    c_arg = PDMath.Atan2(z.c_imag, z.c_real)
End Function

'Complex power w^z
Public Function c_pow(ByRef w As ComplexNumber, ByRef z As ComplexNumber) As ComplexNumber
    
    Dim mag_w As Double
    mag_w = c_mag(w)
    
    Dim arg_w As Double
    arg_w = c_arg(w)
    
    Dim mag As Double
    mag = (mag_w ^ z.c_real) * Exp(-z.c_imag * arg_w)
    
    Dim arg As Double
    arg = z.c_real * arg_w + z.c_imag * Log(mag_w)
    
    c_pow.c_real = mag * Cos(arg)
    c_pow.c_imag = mag * Sin(arg)
    
End Function

'Complex power w^x with real exponent
Public Function c_real_pow(ByRef w As ComplexNumber, ByVal x As Double) As ComplexNumber
    
    Dim mag As Double
    mag = c_mag(w) ^ x
    
    Dim arg As Double
    arg = c_arg(w) * x
    
    c_real_pow.c_real = mag * Cos(arg)
    c_real_pow.c_imag = mag * Sin(arg)
    
End Function

'Complex square root (principal branch)
Public Function c_sqrt(ByRef z As ComplexNumber) As ComplexNumber
    
    Dim r As Double
    r = c_mag(z)
    
    c_sqrt.c_real = Sqr((r + z.c_real) * 0.5)
    c_sqrt.c_imag = Sqr((r - z.c_real) * 0.5)
    
    If (z.c_imag < 0#) Then c_sqrt.c_imag = -c_sqrt.c_imag
    
End Function

'Complex exponential
Public Function c_exp(ByRef z As ComplexNumber) As ComplexNumber
    
    Dim r As Double
    r = Exp(z.c_real)
    
    c_exp.c_real = r * Cos(z.c_imag)
    c_exp.c_imag = r * Sin(z.c_imag)
    
End Function

'Complex logarithm (principal branch)
Public Function c_log(ByRef z As ComplexNumber) As ComplexNumber
    'Dim tmpResult As Double
    'tmpResult = c_mag(z)
    'If (tmpResult <> 0#) Then c_log.c_real = Log(tmpResult) Else c_log.c_real = 0#
    
    'Alternate implementation by Tanner to avoid overflow, with thanks to https://github.com/infusion/Complex.js/blob/master/complex.js
    c_log.c_real = c_loghypot(z)
    c_log.c_imag = PDMath.Atan2(z.c_imag, z.c_real)   'Same thing as c_arg(z), but faster
End Function

Private Function c_loghypot(ByRef z As ComplexNumber) As Double
    
    If (z.c_real = 0#) Then
        If (z.c_imag = 0#) Then
            c_loghypot = 0#
        Else
            c_loghypot = Math.Log(z.c_imag)
        End If
    Else
        If (z.c_imag = 0#) Then
            c_loghypot = Math.Log(z.c_real)
        Else
            'Accurate enough and much faster
            If (Math.Abs(z.c_real) < 3000#) And (Math.Abs(z.c_imag) < 3000#) Then
                c_loghypot = Math.Log(z.c_real * z.c_real + z.c_imag * z.c_imag) * 0.5
            Else
                c_loghypot = Math.Log(z.c_real / Math.Cos(PDMath.Atan2(z.c_imag, z.c_real)))
            End If
        End If
    End If
    
End Function

'--------------------------------------------------------------------------------------------------
'Float-based versions follow.  Any changes to one set of functions should be mirrored to the other.
'--------------------------------------------------------------------------------------------------

Public Function make_complexf(ByVal a As Single, ByVal b As Single) As ComplexNumberF
    make_complexf.c_real = a
    make_complexf.c_imag = b
End Function

'Complex conjugate
Public Function c_conjf(ByRef z As ComplexNumberF) As ComplexNumberF
    c_conjf.c_real = z.c_real
    c_conjf.c_imag = -z.c_imag
End Function

'Complex addition
Public Function c_addf(ByRef w As ComplexNumberF, ByRef z As ComplexNumberF) As ComplexNumberF
    c_addf.c_real = w.c_real + z.c_real
    c_addf.c_imag = w.c_imag + z.c_imag
End Function

'Complex negation
Public Function c_negf(ByRef z As ComplexNumberF) As ComplexNumberF
    c_negf.c_real = -z.c_real
    c_negf.c_imag = -z.c_imag
End Function

'Complex subtraction
Public Function c_subf(ByRef w As ComplexNumberF, ByRef z As ComplexNumberF) As ComplexNumberF
    c_subf.c_real = w.c_real - z.c_real
    c_subf.c_imag = w.c_imag - z.c_imag
End Function

'Complex multiplication
Public Function c_mulf(ByRef w As ComplexNumberF, ByRef z As ComplexNumberF) As ComplexNumberF
    c_mulf.c_real = w.c_real * z.c_real - w.c_imag * z.c_imag
    c_mulf.c_imag = w.c_real * z.c_imag + w.c_imag * z.c_real
End Function

'Complex multiplicative inverse 1/z
Public Function c_invf(ByRef z As ComplexNumberF) As ComplexNumberF

    'There are two mathematically-equivalent formulas for the inverse. For accuracy,
    ' choose the formula with the smaller value of |ratio|.
    Dim ratio As Single, denom As Single
    If (Abs(z.c_real) >= Abs(z.c_imag)) Then
        ratio = z.c_imag / z.c_real
        denom = z.c_real + z.c_imag * ratio
        c_invf.c_real = 1! / denom
        c_invf.c_imag = -ratio / denom
    Else
        ratio = z.c_real / z.c_imag
        denom = z.c_real * ratio + z.c_imag
        c_invf.c_real = ratio / denom
        c_invf.c_imag = -1! / denom
    End If
    
End Function

'Complex division w/z
Public Function c_divf(ByRef w As ComplexNumberF, ByRef z As ComplexNumberF) As ComplexNumberF

    'For accuracy, choose the formula with the smaller value of |ratio|.
    Dim ratio As Single, denom As Single
    If (Abs(z.c_real) >= Abs(z.c_imag)) Then
        ratio = z.c_imag / z.c_real
        denom = 1! / (z.c_real + z.c_imag * ratio)
        c_divf.c_real = (w.c_real + w.c_imag * ratio) * denom
        c_divf.c_imag = (w.c_imag - w.c_real * ratio) * denom
    Else
        ratio = z.c_real / z.c_imag
        denom = 1! / (z.c_real * ratio + z.c_imag)
        c_divf.c_real = (w.c_real * ratio + w.c_imag) * denom
        c_divf.c_imag = (w.c_imag * ratio - w.c_real) * denom
    End If
    
End Function

'Complex magnitude
Public Function c_magf(ByRef z As ComplexNumberF) As Single
    
    Dim tmpz As ComplexNumberF
    tmpz.c_real = Abs(z.c_real)
    tmpz.c_imag = Abs(z.c_imag)
    
    'For accuracy, choose the formula with the smaller value of |ratio|.
    Dim ratio As Single
    If (tmpz.c_real >= tmpz.c_imag) Then
        If (tmpz.c_real <> 0!) Then ratio = tmpz.c_imag / tmpz.c_real Else ratio = 0!
        c_magf = tmpz.c_real * Sqr(1! + ratio * ratio)
    Else
        If (tmpz.c_imag <> 0!) Then ratio = ratio = tmpz.c_real / tmpz.c_imag Else ratio = 0!
        c_magf = tmpz.c_imag * Sqr(1! + ratio * ratio)
    End If
    
End Function

'Complex argument (angle) in [-pi,+pi]
Public Function c_argf(ByRef z As ComplexNumberF) As Single
    c_argf = PDMath.Atan2(z.c_imag, z.c_real)
End Function

'Complex power w^z
Public Function c_powf(ByRef w As ComplexNumberF, ByRef z As ComplexNumberF) As ComplexNumberF
    
    Dim mag_w As Single
    mag_w = c_magf(w)
    
    Dim arg_w As Single
    arg_w = c_argf(w)
    
    Dim mag As Single
    mag = (mag_w ^ z.c_real) * Exp(-z.c_imag * arg_w)
    
    Dim arg As Single
    arg = z.c_real * arg_w + z.c_imag * Log(mag_w)
    
    c_powf.c_real = mag * Cos(arg)
    c_powf.c_imag = mag * Sin(arg)
    
End Function

'Complex power w^x with real exponent
Public Function c_real_powf(ByRef w As ComplexNumberF, ByVal x As Single) As ComplexNumberF
    
    Dim mag As Single
    mag = c_magf(w) ^ x
    
    Dim arg As Single
    arg = c_argf(w) * x
    
    c_real_powf.c_real = mag * Cos(arg)
    c_real_powf.c_imag = mag * Sin(arg)
    
End Function

'Complex square root (principal branch)
Public Function c_sqrtf(ByRef z As ComplexNumberF) As ComplexNumberF
    
    Dim r As Single
    r = c_magf(z)
    
    c_sqrtf.c_real = Sqr((r + z.c_real) * 0.5!)
    c_sqrtf.c_imag = Sqr((r - z.c_real) * 0.5!)
    
    If (z.c_imag < 0!) Then c_sqrtf.c_imag = -c_sqrtf.c_imag
    
End Function

'Complex exponential
Public Function c_expf(ByRef z As ComplexNumberF) As ComplexNumberF
    
    Dim r As Single
    r = Exp(z.c_real)
    
    c_expf.c_real = r * Cos(z.c_imag)
    c_expf.c_imag = r * Sin(z.c_imag)
    
End Function

'Complex logarithm (principal branch)
Public Function c_logf(ByRef z As ComplexNumberF) As ComplexNumberF
    'Dim tmpResult As Double
    'tmpResult = c_mag(z)
    'If (tmpResult <> 0#) Then c_log.c_real = Log(tmpResult) Else c_log.c_real = 0#
    
    'Alternate implementation by Tanner to avoid overflow, with thanks to https://github.com/infusion/Complex.js/blob/master/complex.js
    c_logf.c_real = c_loghypotf(z)
    c_logf.c_imag = PDMath.Atan2(z.c_imag, z.c_real)   'Same thing as c_arg(z), but faster
End Function

Private Function c_loghypotf(ByRef z As ComplexNumberF) As Single
    
    If (z.c_real = 0!) Then
        If (z.c_imag = 0!) Then
            c_loghypotf = 0!
        Else
            c_loghypotf = Math.Log(Abs(z.c_imag))
        End If
    Else
        If (z.c_imag = 0!) Then
            c_loghypotf = Math.Log(Abs(z.c_real))
        Else
            'Accurate enough and much faster
            If (Math.Abs(z.c_real) < 3000!) And (Math.Abs(z.c_imag) < 3000!) Then
                c_loghypotf = Math.Log(z.c_real * z.c_real + z.c_imag * z.c_imag) * 0.5!
            Else
                c_loghypotf = Math.Log(z.c_real / Math.Cos(PDMath.Atan2(z.c_imag, z.c_real)))
            End If
        End If
    End If
    
End Function
