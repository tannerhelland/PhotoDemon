VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdRandomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Random Number Generator
'Copyright 2015-2026 by Tanner Helland
'Created: 23/June/15 (but assembled from bits scattered throughout PD, many from years earlier)
'Last updated: 07/August/17
'Last update: add gaussian distribution functions, which use a Box-Muller transform to produce random
'             floats with a normal distribution.
'
'VB's internal randomize function is confusing and ill-conceived, especially when it comes to seeding it.  This class aims to
' make random number generation far more predictable (lol, ironic?) and convenient.
'
'For now, it's just a thin wrapper to VB's internal randomize functions, but in the future, I may include functions that
' provide better random number capabilities.
'
'Many thanks to the following articles, which were invaluable for improving this class:
' http://web.archive.org/web/20110113032316/http://www.15seconds.com/issue/051110.htm
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' http://stackoverflow.com/questions/22384451/vb6-how-to-get-c-like-integer-overflow/22389687#22389687
' https://en.wikipedia.org/wiki/Box%E2%80%93Muller_transform
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current seed.  Both float and int versions are stored, as different PRNGs need different inputs
Private m_Seed_Float As Double, m_Seed_Int As Long

'Some PRNGs require seeds that meet certain conditions (e.g. non-negative, etc).  As such, *these seed values may deviate from
' the seeds supplied by the user*, by design.
Private m_Seed_WH_Int As Long

'Current bounds (integer functions only; the floating point functions return values between 0 and 1, by design).
Private m_LowBound As Long, m_HighBound As Long

'Intermediary calculation values for the Wichmann-Hill algorithm; thank you to
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' for notes on various VB implementations and their quirks.
Private iX As Long, iY As Long, iZ As Long
Private m_WH_Float As Double

'Helper functions for generating a new (predictable) seed from various inputs
Friend Sub SetSeed_Int(ByVal newSeed As Long)
    m_Seed_Float = newSeed
    m_Seed_Int = newSeed
    ApplySeed
End Sub

Friend Sub SetSeed_Float(ByVal newSeed As Double)
    m_Seed_Float = newSeed
    m_Seed_Int = GetSafeIntFromDouble(m_Seed_Float)
    ApplySeed
End Sub

'Hash a given string into a pseudo-random seed, using a fast (but arbitrary) hash function
Friend Sub SetSeed_String(ByRef seedString As String)
    
    'Use MD5 to get a "numerical"-ish representation of the underlying string.
    Dim cCrypto As pdCrypto
    Set cCrypto = New pdCrypto
    If cCrypto.QuickHash(PDCA_MD5, StrPtr(seedString), LenB(seedString)) Then
        
        'MD5 requires 16 bytes; we ultimately want to use a double-type value to seed our generator, so go ahead
        ' and grab all 16 bytes but just seed the first 8.
        Dim twoDoubles() As Double
        ReDim twoDoubles(0 To 1) As Double
        If (cCrypto.RetrieveHashedDataPtr(VarPtr(twoDoubles(0)), 16&)) Then
            m_Seed_Float = twoDoubles(0)
            CopyMemoryStrict VarPtr(m_Seed_Int), VarPtr(twoDoubles(1)), 4&
            ApplySeed
        End If
        
    End If

End Sub

Friend Sub SetSeed_AutomaticAndRandom()
    m_Seed_Float = Timer * Now
    m_Seed_Int = GetSafeIntFromDouble(m_Seed_Float)
    ApplySeed
End Sub

'PD uses Doubles for maximum randomness, but some random number generation schemes use Integer inputs.
' To safely generate a random Int from a given Double value (which has a much larger range!), we need to
' use a helper function.
Private Function GetSafeIntFromDouble(ByVal srcDouble As Double) As Long
    Dim val1 As Long, val2 As Long
    CopyMemoryStrict VarPtr(val1), VarPtr(srcDouble), 4
    CopyMemoryStrict VarPtr(val2), VarPtr(srcDouble) + 4, 4
    GetSafeIntFromDouble = val1 Xor val2
End Function

'Return the current seed.  Note that this class always stores the seed as a Double, regardless of how it was originally supplied.
Friend Function GetSeed() As Double
    GetSeed = m_Seed_Float
End Function

'Use the current seed to actually seed all supported PRNG engines.
Private Sub ApplySeed()
    
    'First, seed VB's internal generator
    Rnd -1
    Randomize m_Seed_Float
    
    'Next, seed any custom number generators
    
    'Wichmann-Hill initialization is pretty easy; we just have to ensure we start with a positive, non-zero value...
    m_Seed_WH_Int = m_Seed_Int
    If (m_Seed_WH_Int < 0) Then m_Seed_WH_Int = m_Seed_WH_Int And &H7FFFFFFF
    
    '...then we generate an initial set of offsets for the algorithm.
    iX = (m_Seed_WH_Int Mod 30269)
    iY = (m_Seed_WH_Int Mod 30307)
    iZ = (m_Seed_WH_Int Mod 30323)
    If (iX = 0) Then iX = 171
    If (iY = 0) Then iY = 172
    If (iZ = 0) Then iZ = 170
    
End Sub

'Set bounds for the integer Rnd functions
Friend Sub SetRndIntegerBounds(ByVal lowBound As Long, ByVal highBound As Long)
    m_LowBound = lowBound
    m_HighBound = highBound
End Sub

Private Sub Class_Initialize()
    
    SetSeed_Int 0
    
    'Set default integer bounds.  Note that we trim the high bound a bit to avoid overflow errors.
    m_LowBound = 0
    m_HighBound = 2147483640
    
End Sub

'Return a random integer using VB's internal randomize engine.  If supplied earlier, bounds are used.
Friend Function GetRandomInt_VB() As Long
    GetRandomInt_VB = Int((m_HighBound - m_LowBound + 1) * Rnd + m_LowBound)
End Function

Friend Function GetRandomIntRange_VB(ByVal lowBound As Long, ByVal highBound As Long) As Long
    GetRandomIntRange_VB = Int((highBound - lowBound + 1) * Rnd + lowBound)
End Function

'Return a random float using VB's internal randomize engine.  Bounds are ignored.  This is kind of a stupid function, as it would
' be faster to just use Rnd yourself, but it's included here for completeness.
Friend Function GetRandomFloat_VB() As Double
    GetRandomFloat_VB = Rnd
End Function

'Return a random integer using the Wichmann-Hill PRNG.  If supplied earlier, bounds are used.
Friend Function GetRandomInt_WH() As Long
    GetRandomInt_WH = Int((m_HighBound - m_LowBound + 1) * GetRandomFloat_WH() + m_LowBound)
End Function

Friend Function GetRandomIntRange_WH(ByVal lowBound As Long, ByVal highBound As Long) As Long
    GetRandomIntRange_WH = Int((highBound - lowBound + 1) * GetRandomFloat_WH() + lowBound)
End Function

'Return a random float using the Wichmann-Hill PRNG.  Pretty fast, good distribution too.
Friend Function GetRandomFloat_WH() As Double
    
    'Generate new offsets, using the previous offsets as our inputs
    iX = (171& * iX) Mod 30269&
    iY = (172& * iY) Mod 30307&
    iZ = (170& * iZ) Mod 30323&
    
    'Generate a random float value.  (Note that we use multiplication rather than division, for performance reasons.)
    Const RND_X_DIVISOR As Double = 1# / 30269#
    Const RND_Y_DIVISOR As Double = 1# / 30307#
    Const RND_Z_DIVISOR As Double = 1# / 30323#
    m_WH_Float = CDbl(iX) * RND_X_DIVISOR + CDbl(iY) * RND_Y_DIVISOR + CDbl(iZ) * RND_Z_DIVISOR
    
    '...then return the floating-point portion
    GetRandomFloat_WH = m_WH_Float - Int(m_WH_Float)
    
End Function

'Gaussian distribution functions are also provided; obviously, they impose a somewhat unpleasant performance hit.
' IMPORTANTLY, note that Gaussian functions, by definition, return values on the theoretical range
' [-DoubleMax, +DoubleMax], with the distribution centered around [0].  This differs from the standard behavior
' of returning floats on the range [0.0, 1.0]
Friend Function GetGaussianFloat_VB() As Double
    
    'Pass two random values them through a Box-Muller transform to produce a gaussian distribution.
    ' (The math behind this is a little weird; see https://en.wikipedia.org/wiki/Box%E2%80%93Muller_transform)
    GetGaussianFloat_VB = Sqr(-2# * Log(GetRandomFloat_VB())) * Cos(PI_DOUBLE * GetRandomFloat_VB())
    
End Function

'Gaussian distribution functions are also provided; obviously, they impose a somewhat unpleasant performance hit.
' IMPORTANTLY, note that Gaussian functions, by definition, return values on the theoretical range
' [-DoubleMax, +DoubleMax], with the distribution centered around [0].  This differs from the standard behavior
' of returning floats on the range [0.0, 1.0]
Friend Function GetGaussianFloat_WH() As Double
    
    'Pass two random values them through a Box-Muller transform to produce a gaussian distribution.
    ' (The math behind this is a little weird; see https://en.wikipedia.org/wiki/Box%E2%80%93Muller_transform)
    GetGaussianFloat_WH = Sqr(-2# * Log(GetRandomFloat_WH())) * Cos(PI_DOUBLE * GetRandomFloat_WH())
    
End Function
