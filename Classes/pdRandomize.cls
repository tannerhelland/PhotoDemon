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
'Copyright 2015-2015 by Tanner Helland
'Created: 23/June/15 (but assembled from bits scattered throughout PD, many from years earlier)
'Last updated: 14/September/15
'Last update: fix potential integer overflow; thanks to GioRock for catching and reporting
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
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'These two custom types are used for hashing input strings into pseudo-random seeds
Private Type CURRENCY_CURRENCY
    Value As Currency
End Type

Private Type CURRENCY_2LONGS
    ValueLo As Long
    ValueHi As Long
End Type

'Intermediary calculation values for the Wichmann-Hill algorithm; thank you to
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' for notes on various VB implementations and their quirks.
Private iX As Long, iy As Long, iZ As Long
Private m_WH_Float As Double

'Helper functions for generating a new (predictable) seed from various inputs
Public Sub setSeed_Int(ByVal newSeed As Long)
    m_Seed_Float = newSeed
    m_Seed_Int = newSeed
    applySeed
End Sub

Public Sub setSeed_Float(ByVal newSeed As Double)
    m_Seed_Float = newSeed
    m_Seed_Int = getSafeIntFromDouble(m_Seed_Float)
    applySeed
End Sub

'Hash a given string into a pseudo-random seed, using the well-known djb2 hash function
Public Sub setSeed_String(ByRef seedString As String)

    'Basic djb2 string hash based on http://www.cse.yorku.ca/~oz/hash.html
    ' Thanks to Bob77 for a quick-and-dirty VB implementation (http://stackoverflow.com/questions/22384451/vb6-how-to-get-c-like-integer-overflow/22389687#22389687)
    Dim C2L As CURRENCY_2LONGS
    C2L.ValueLo = 5381
    
    Dim CC As CURRENCY_CURRENCY
    
    Dim i As Long
    For i = 1 To Len(seedString)
        LSet CC = C2L
        CC.Value = CC.Value * 33@ + CCur(AscW(Mid$(seedString, i, 1))) / 10000@
        LSet C2L = CC
        C2L.ValueHi = 0
    Next i

    m_Seed_Int = C2L.ValueLo
    m_Seed_Float = m_Seed_Int
    applySeed

End Sub

Public Sub setSeed_AutomaticAndRandom()
    m_Seed_Float = Timer * Now
    m_Seed_Int = getSafeIntFromDouble(m_Seed_Float)
    applySeed
End Sub

'PD uses Doubles for maximum randomness, but some random number generation schemes use Integer inputs.  To safely generate a random
' Int from a given Double value (which has a much larger range!), we use a custom Mod function (necessary because VB's built-in
' Mod only works on Longs).
Private Function getSafeIntFromDouble(ByVal srcDouble As Double) As Long
    Const LONG_MAX_AS_DOUBLE As Double = LONG_MAX
    getSafeIntFromDouble = srcDouble - (Int(srcDouble / LONG_MAX_AS_DOUBLE) * LONG_MAX_AS_DOUBLE)
End Function

'Return the current seed.  Note that this class always stores the seed as a Double, regardless of how it was originally supplied.
Public Function getSeed() As Double
    getSeed = m_Seed_Float
End Function

'Use the current seed to actually seed all supported PRNG engines.
Private Sub applySeed()
    
    'First, seed VB's internal generator
    Rnd -1
    Randomize m_Seed_Float
    
    'Next, seed any custom number generators
    
    'Wichmann-Hill initialization is pretty easy; we just have to ensure we start with a positive, non-zero value...
    m_Seed_WH_Int = m_Seed_Int
    If m_Seed_WH_Int < 0 Then m_Seed_WH_Int = m_Seed_WH_Int And &H7FFFFFFF
    
    '...then we generate an initial set of offsets for the algorithm.
    iX = (m_Seed_WH_Int Mod 30269)
    iy = (m_Seed_WH_Int Mod 30307)
    iZ = (m_Seed_WH_Int Mod 30323)
    If iX = 0 Then iX = 171
    If iy = 0 Then iy = 172
    If iZ = 0 Then iZ = 170
    
End Sub

'Set bounds for the integer Rnd functions
Public Sub setRndIntegerBounds(ByVal lowBound As Long, ByVal highBound As Long)
    m_LowBound = lowBound
    m_HighBound = highBound
End Sub

Private Sub Class_Initialize()
    
    setSeed_Int 0
    
    'Set default integer bounds.  Note that we trim the high bound a bit to avoid overflow errors.
    m_LowBound = 0
    m_HighBound = 2147483640
    
End Sub

'Return a random integer using VB's internal randomize engine.  If supplied earlier, bounds are used.
Public Function getRandomInt_VB() As Long
    getRandomInt_VB = Int((m_HighBound - m_LowBound + 1) * Rnd + m_LowBound)
End Function

'Return a random float using VB's internal randomize engine.  Bounds are ignored.  This is kind of a stupid function, as it would
' be faster to just use Rnd yourself, but it's included here for completeness.
Public Function getRandomFloat_VB() As Double
    getRandomFloat_VB = Rnd
End Function

'Return a random integer using the Wichmann-Hill PRNG.  If supplied earlier, bounds are used.
Public Function getRandomInt_WH() As Long
    getRandomInt_WH = Int((m_HighBound - m_LowBound + 1) * getRandomFloat_WH + m_LowBound)
End Function

'Return a random float using the Wichmann-Hill PRNG.  Pretty fast, good distribution too.
Public Function getRandomFloat_WH() As Double
    
    'Generate new offsets, using the previous offsets as our inputs
    iX = (171 * iX) Mod 30269
    iy = (172 * iy) Mod 30307
    iZ = (170 * iZ) Mod 30323
    
    'Generate a random float value...
    m_WH_Float = CDbl(iX) / 30269# + CDbl(iy) / 30307# + CDbl(iZ) / 30323#
    
    '...then return the floating-point portion
    getRandomFloat_WH = m_WH_Float - Int(m_WH_Float)
    
End Function
