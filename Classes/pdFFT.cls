VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon FFT Class
'Copyright 2014-2015 by Tanner Helland
'Created: 27/February/15
'Last updated: 27/February/15
'Last update: initial build
'
'Per its name, this class provides a simple FFT interface.  After researching a great many FFT implementations,
' I decided that it would probably be easiest to just write my own pure-VB implementation, rather than deal with
' the headache of trying to wrap something like FFTW.  There's a speed hit to doing this in pure VB (obviously),
' but it's also nicely portable, and there are enough FFT references out there that the work wasn't particularly
' demanding.
'
'This class currently includes a few different FFT approaches.  Many thanks to the following references, which were
' invaluable in getting this code right:
'
'http://paulbourke.net/miscellaneous/dft/
' (Paul Bourke's website simply states "Any source code found here may be freely used provided credits are given to
'  the author", so thanks, Paul, for the great reference!)
'
' http://cnx.org/contents/ce67266a-1851-47e4-8bfc-82eb447212b4@7/Decimation-in-time_%28DIT%29_Radix
' http://cnx.org/contents/e240a1a1-c1cc-4427-94e1-b9d978c01421@6/Efficient_FFT_Algorithm_and_Pr
' (These two articles are CC-1.0-BY licensed, so many thanks to Douglas L. Jones for not just sharing his excellent code,
'  but wrapping it with a really excellent discussion of FFT implementations.)
'
'So far, only radix-2 approaches are used.  Radix-4 would be faster, but also significantly more complicated, and I'm
' not yet comfortable enough with FFTs to tackle them.  Outside help is always welcome!
'
'For simplicity, this class requires incoming data to be padded to powers of 2.  Zero-pad data as necessary.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const PI_HALF As Double = 1.5707963267949
Private Const PI_DOUBLE As Double = 6.28318530717958

'Cos and sin values are constant for a given length n, which is especially relevant when working with images (where each
' row or column is an identical height).
Private m_N As Long
Private m_M As Long
Private m_Direction As Boolean

Public Function prepFFT(ByVal n As Long, Optional ByVal forwardTransform As Boolean = True) As Boolean

    'If n = our previous n, we don't need to rebuild our lookup tables
    If (n = m_N) And (n <> 0) And (forwardTransform = m_Direction) Then
        prepFFT = True
        
    Else
    
        'Start by making sure that n is a valid power of 2
        m_M = Log(n) / Log(2#)
        
        If (m_M <> Int(m_M)) Then
        
            Debug.Print "WARNING! n must be a power of 2!  If using it on an array, make sure to submit n as (uBound() + 1)."
            prepFFT = False
        
        Else
        
            m_N = n
            m_Direction = forwardTransform
            
            'Some methods allow you to precalculate certain FFT values.  I have not found a clean, working example of this,
            ' but will leave this note here for potential future use.
            
            'Note that initialization was successful
            prepFFT = True
            
        End If
        
    End If
    
End Function

'Apply a forward FFT on a source data set.  n should be the same as the n supplied to prepFFT, above.
' An in-place transform is used, so srcReal() and srcImag() WILL BE OVERWRITTEN.  Plan accordingly.
'
'This function uses a simple radix-2 approach, which is not as fast as a radix-2, but is significantly simpler code.
Public Function FFT_1D_Radix2(ByVal n As Long, ByRef srcReal() As Single, ByRef srcImag() As Single, Optional ByVal forwardTransform As Boolean = True) As Boolean

    'Make sure all inputs align with n
    If (n = UBound(srcReal) + 1) And (n = UBound(srcImag) + 1) And (n = m_N) Then
    
        'A lot of ints and floats are required for FFTs
        Dim i As Long, j As Long, k As Long
        Dim m As Long, n1 As Long, n2 As Long
        
        m = m_M
        
        Dim c As Single, s As Single, e As Single, a As Single
        Dim t1 As Single, t2 As Single
        
        'Start with bit-reversal
        n2 = n / 2
        For i = 1 To n - 2
        
            n1 = n2
            
            Do While (j >= n1)
                j = j - n1
                n1 = n1 / 2
            Loop
            
            j = j + n1
            
            If (i < j) Then
                
                t1 = srcReal(i)
                srcReal(i) = srcReal(j)
                srcReal(j) = t1
                t1 = srcImag(i)
                srcImag(i) = srcImag(j)
                srcImag(j) = t1
                
            End If
        
        Next i
        
        'Bits are now reversed.  Apply the FFT!
        n1 = 0
        n2 = 1
        
        For i = 0 To m - 1
        
            n1 = n2
            n2 = n2 + n2
            
            'Non-optimized approach...
            e = (-1 * PI_DOUBLE) / n2
            If Not forwardTransform Then e = -e
            
            a = 0
            
            For j = 0 To n1 - 1
                
                'Calculate cos/sin manually (there must be a way to precalculate these, but I haven't found it yet)
                c = Cos(a)
                s = Sin(a)
                a = a + e
                
                For k = j To n - 1 Step n2
                    t1 = c * srcReal(k + n1) - s * srcImag(k + n1)
                    t2 = s * srcReal(k + n1) + c * srcImag(k + n1)
                    srcReal(k + n1) = srcReal(k) - t1
                    srcImag(k + n1) = srcImag(k) - t2
                    srcReal(k) = srcReal(k) + t1
                    srcImag(k) = srcImag(k) + t2
                Next k
                
            Next j
        
        Next i
        
        'If this is the forward transform, values must be normalized against n
        If forwardTransform Then
            For i = 0 To n - 1

                Dim invN As Single
                invN = 1 / n

                srcReal(i) = srcReal(i) * invN
                srcImag(i) = srcImag(i) * invN

            Next i
        End If
        
        'Return success
        FFT_1D_Radix2 = True
        
    'Inputs do not align with n; we cannot proceed.
    Else
        Debug.Print "Array dimensions and n are mismatched.  Make sure you called prepFFT first, and declared your arrays appropriately."
        FFT_1D_Radix2 = False
    End If

End Function

'Apply a forward FFT on a source data set.  n should be the same as the n supplied to prepFFT, above.
' An in-place transform is used, so srcReal() and srcImag() WILL BE OVERWRITTEN.  Plan accordingly.
'
'This function uses a simple radix-2 approach, which is not as fast as a radix-4, but is significantly simpler code.
' As an attempt to profile VB's specific weaknesses, this function uses estimation techniques based around Sqr()
' instead of pure Cos/Sin functions.  By my testing, this is 10-15% faster (compiled) than the trigonmetric approach.
Public Function FFT_1D_Radix2_NoTrig(ByVal n As Long, ByRef srcReal() As Single, ByRef srcImag() As Single, Optional ByVal forwardTransform As Boolean = True) As Boolean

    'Make sure all inputs align with n
    If (n = UBound(srcReal) + 1) And (n = UBound(srcImag) + 1) And (n = m_N) Then
    
        'A lot of ints and floats are required for FFTs
        Dim i As Long, i1 As Long, i2 As Long, j As Long, k As Long, l As Long, l1 As Long, l2 As Long
        Dim m As Long, n1 As Long, n2 As Long
        
        m = m_M
        
        Dim c1 As Single, c2 As Single, u1 As Single, u2 As Single, z As Single
        Dim t1 As Single, t2 As Single
        
        'Start with bit-reversal
        n2 = n / 2
        For i = 1 To n - 2
        
            n1 = n2
            
            Do While (j >= n1)
                j = j - n1
                n1 = n1 / 2
            Loop
            
            j = j + n1
            
            If (i < j) Then
                
                t1 = srcReal(i)
                srcReal(i) = srcReal(j)
                srcReal(j) = t1
                t1 = srcImag(i)
                srcImag(i) = srcImag(j)
                srcImag(j) = t1
                
            End If
        
        Next i
        
        'Bits are now reversed.  Apply the FFT!
        c1 = -1#
        c2 = 0#
        l2 = 1
        
        For l = 0 To m - 1
           l1 = l2
           l2 = l2 * 2
           u1 = 1#
           u2 = 0#
           
           For j = 0 To l1 - 1
                
                For i = j To n - 1 Step l2
                    
                    i1 = i + l1
                    t1 = u1 * srcReal(i1) - u2 * srcImag(i1)
                    t2 = u1 * srcImag(i1) + u2 * srcReal(i1)
                    
                    srcReal(i1) = srcReal(i) - t1
                    srcImag(i1) = srcImag(i) - t2
                    srcReal(i) = srcReal(i) + t1
                    srcImag(i) = srcImag(i) + t2
                    
                Next i
                
                z = u1 * c1 - u2 * c2
                u2 = u1 * c2 + u2 * c1
                u1 = z
           
           Next j
           
           c2 = Sqr((1# - c1) / 2#)
           If forwardTransform Then c2 = -c2
           c1 = Sqr((1# + c1) / 2#)
           
        Next l
        
        'If this is the forward transform, values must be normalized against n
        If forwardTransform Then
            For i = 0 To n - 1

                Dim invN As Single
                invN = 1 / n

                srcReal(i) = srcReal(i) * invN
                srcImag(i) = srcImag(i) * invN

            Next i
        End If
        
        'Return success
        FFT_1D_Radix2_NoTrig = True
        
    'Inputs do not align with n; we cannot proceed.
    Else
        Debug.Print "Array dimensions and n are mismatched.  Make sure you called prepFFT first, and declared your arrays appropriately."
        FFT_1D_Radix2_NoTrig = False
    End If

End Function

'Dump a brief FFT test to the debug window.
' If you want some predictable FFT output, try the following link for test data:
' http://www.sccon.ca/sccon/fft/fft3.htm
Public Sub testFFT()

    Dim testReal() As Single, testImag() As Single
    
    Dim testSize As Long
    testSize = 4095
    
    ReDim testReal(0 To testSize) As Single
    ReDim testImag(0 To testSize) As Single
    
    Dim i As Long, j As Long
    
    'Use this loop to populate the initial data, if desired
    'For i = 0 To testSize
    '    testReal(i) = i
    '    testImag(i) = testSize - i
    'Next i
    
    'Use this loop to run a bunch of FFTs (helpful for profiling)
    Dim startTime As Single
    startTime = Timer
    
    Me.prepFFT testSize + 1, True
    
    For j = 0 To 2047
        Me.FFT_1D_Radix2_NoTrig testSize + 1, testReal, testImag, True
        Me.FFT_1D_Radix2_NoTrig testSize + 1, testReal, testImag, False
    Next j
    
    MsgBox Timer - startTime

    'The code lines below are very helpful for testing before/after FFT results, to make sure the forward/inverse transform
    ' returns identical data.
    
'    dumpDebugDataToScreen testReal
'
'    Me.prepFFT testSize + 1, True
'    Me.FFT_1D_Radix2_NoTrig testSize + 1, testReal, testImag, True
'
'    dumpDebugDataToScreen testReal
'
'    Me.prepFFT testSize + 1, False
'    Me.FFT_1D_Radix2_NoTrig testSize + 1, testReal, testImag, False
'
'    dumpDebugDataToScreen testReal

End Sub

'Internal function to dump an array to the debug window.  Don't use with large arrays (UBound > ~30).
Private Sub dumpDebugDataToScreen(ByRef srcArray() As Single)
    
    Dim tmpString As String
    tmpString = ""
    
    Dim i As Long
    For i = LBound(srcArray) To UBound(srcArray)
        tmpString = tmpString & CStr(srcArray(i))
        If i < UBound(srcArray) Then tmpString = tmpString & ", "
    Next i
    
    Debug.Print tmpString
    
End Sub
