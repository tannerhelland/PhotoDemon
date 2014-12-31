VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPerlin3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This class is Copyright 2003 Steve McMahon, vbaccelerator.com

'This class was originally written by Steve McMahon.

'You may download the original version of this code from the following link (good as of January '13):
' http://www.vbaccelerator.com/home/VB/Code/vbMedia/Algorithmic_Images/Perlin_Noise/article.asp

'To the best of my knowledge, this code is released under a CC-BY-1.0 license.  (Assumed from the footer text of vbaccelerator.com: "All contents of this web site are licensed under a Creative Commons Licence, except where otherwise noted.")
' You may access a complete copy of this license at the following link:
' http://creativecommons.org/licenses/by/1.0/

'Many thanks to Steve McMahon and vbaccelerator.com for this excellent perlin noise class.

'NOTE: in March '14, I profiled a bunch of the functions in this class and rewrote some of them to
'      improve performance.  If 3D noise is not needed, I strongly recommend caching a z variable
'      in advance and using the Noise2D function.  Depending on other variables, this can result
'      in 10%+ faster performance, which is especially helpful when calculating noise per-pixel.

Option Explicit

'
' vbAccelerator.com
' Perlin Noise
' Pretty much optimal algorithm for calculation, converted to VB from
'// JAVA REFERENCE IMPLEMENTATION OF IMPROVED NOISE - COPYRIGHT 2002 KEN PERLIN.
' http://mrl.nyu.edu/~perlin/noise/
' See also http://www.noisemachine.com/
'
' Takes approx 30ms for a 100x100 grid with all optimisations on,
' using AMD Athlon 2k.
'
Private p() As Long
Private permutation() As Long

Private m_cachedZ As Double, m_cachedIZ As Long, m_cachedW As Double

'Part of the optimized NoiseCache() function, for callers only splicing a single level of the noise cube
Public Sub cacheZValue(ByVal srcZ As Double)
    m_cachedIZ = Int(srcZ) And &HFF&
    m_cachedZ = srcZ - Int(srcZ)
    m_cachedW = Fade(m_cachedZ)
End Sub

'2D noise function, designed for use with the cacheZValue sub above.  If a functions is only using
' a single plane of the 3D noise cube, this is quite a bit faster than a full 3D calculation.
Public Function Noise2D(ByVal x As Double, ByVal y As Double) As Double

    'Note that we'll be doing a full 3D calculation, but using cached values for everything related to z
    Dim iX As Long, iY As Long
    Dim u As Double, v As Double
    Dim a As Long, aa As Long, ab As Long, b As Long, ba As Long, bb As Long

    'Find unit cube that contains point:
    iX = Int(x) And &HFF&
    iY = Int(y) And &HFF&
   
    'Find relative x,y,z of point in cube:
    x = x - Int(x)
    y = y - Int(y)
   
    'Compute Fade curves for each of x,y,z
    u = Fade(x)
    v = Fade(y)
   
    'Hash coordinates of the 8 cube corners:
    a = p(iX) + iY
    aa = p(a) + m_cachedIZ
    ab = p(a + 1) + m_cachedIZ
    b = p(iX + 1) + iY
    ba = p(b) + m_cachedIZ
    bb = p(b + 1) + m_cachedIZ
   
    'Add the blended results from 8 corners of the cube:
    'Noise2D = lerp(m_cachedW, lerp(v, lerp(u, grad(p(aa), x, y, m_cachedZ), _
                              grad(p(ba), x - 1, y, m_cachedZ)), _
                           lerp(u, grad(p(ab), x, y - 1, m_cachedZ), _
                              grad(p(bb), x - 1, y - 1, m_cachedZ))), _
                     lerp(v, lerp(u, grad(p(aa + 1), x, y, m_cachedZ - 1), _
                              grad(p(ba + 1), x - 1, y, m_cachedZ - 1)), _
                           lerp(u, grad(p(ab + 1), x, y - 1, m_cachedZ - 1), _
                              grad(p(bb + 1), x - 1, y - 1, m_cachedZ - 1))))
    
    'Subsequent testing has shown a slight (2-3%) performance improvement on modern processors
    ' (Intel Haswell, specifically) if I unroll the function above into discrete parts.  Others
    ' are welcome to test this on their own PCs...
    Dim gradCache1 As Double, gradCache2 As Double, gradCache3 As Double
    Dim gradCache4 As Double, gradCache5 As Double, gradCache6 As Double
    Dim gradCache7 As Double, gradCache8 As Double
    
    Dim lerpCache1 As Double, lerpCache2 As Double, lerpCache3 As Double, lerpCache4 As Double
    Dim lerpCache5 As Double, lerpCache6 As Double
    
    gradCache8 = grad(p(aa), x, y, m_cachedZ)
    gradCache4 = grad(p(ba), x - 1, y, m_cachedZ)
    gradCache7 = grad(p(ab), x, y - 1, m_cachedZ)
    gradCache3 = grad(p(bb), x - 1, y - 1, m_cachedZ)
    gradCache5 = grad(p(ab + 1), x, y - 1, m_cachedZ - 1)
    gradCache1 = grad(p(bb + 1), x - 1, y - 1, m_cachedZ - 1)
    gradCache6 = grad(p(aa + 1), x, y, m_cachedZ - 1)
    gradCache2 = grad(p(ba + 1), x - 1, y, m_cachedZ - 1)
    
    lerpCache1 = lerp(u, gradCache8, gradCache4)
    lerpCache2 = lerp(u, gradCache7, gradCache3)
    lerpCache5 = lerp(v, lerpCache1, lerpCache2)
    
    lerpCache4 = lerp(u, gradCache5, gradCache1)
    lerpCache3 = lerp(u, gradCache6, gradCache2)
    lerpCache6 = lerp(v, lerpCache3, lerpCache4)
    
    Noise2D = lerp(m_cachedW, lerpCache5, lerpCache6)
    
End Function


'The original Noise() function, copied directly from Ken Perlin's reference implementation
Public Function Noise(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double

    Dim iX As Long
    Dim iY As Long
    Dim iZ As Long
    Dim u As Double
    Dim v As Double
    Dim w As Double
    Dim a As Long
    Dim aa As Long
    Dim ab As Long
    Dim b As Long
    Dim ba As Long
    Dim bb As Long

   ' Find unit cube that contains point:
   iX = Int(x) And &HFF&
   iY = Int(y) And &HFF&
   iZ = Int(z) And &HFF&
   
   ' Find relative x,y,z of point in cube:
   x = x - Int(x)
   y = y - Int(y)
   z = z - Int(z)
   
   ' Compute Fade curves for each of x,y,z
   u = Fade(x)
   v = Fade(y)
   w = Fade(z)
   
   ' Hash coordinates of the 8 cube corners:
   a = p(iX) + iY
   aa = p(a) + iZ
   ab = p(a + 1) + iZ
   b = p(iX + 1) + iY
   ba = p(b) + iZ
   bb = p(b + 1) + iZ
   
   'Add the blended results from 8 corners of the cube:
   Noise = lerp(w, lerp(v, lerp(u, grad(p(aa), x, y, z), _
                              grad(p(ba), x - 1, y, z)), _
                           lerp(u, grad(p(ab), x, y - 1, z), _
                              grad(p(bb), x - 1, y - 1, z))), _
                     lerp(v, lerp(u, grad(p(aa + 1), x, y, z - 1), _
                              grad(p(ba + 1), x - 1, y, z - 1)), _
                           lerp(u, grad(p(ab + 1), x, y - 1, z - 1), _
                              grad(p(bb + 1), x - 1, y - 1, z - 1))))
    
End Function

Private Function Fade(ByRef t As Double) As Double
    Fade = t * t * t * (t * (t * 6# - 15#) + 10#)
End Function

Private Function lerp(ByRef t As Double, ByRef a As Double, ByRef b As Double) As Double
    lerp = a + t * (b - a)
End Function

Private Function grad(ByVal hash As Long, ByRef x As Double, ByRef y As Double, ByRef z As Double) As Double

    Dim h As Long
    Dim u As Double, v As Double, g As Double

    '// CONVERT LO 4 BITS OF HASH CODE
    h = hash And &HF&
   
    '// INTO 12 GRADIENT DIRECTIONS.
    If (h < 8) Or (h = 12) Or (h = 13) Then
       u = x
    Else
       u = y
    End If
   
    If (h < 4) Or (h = 12) Or (h = 13) Then
       v = y
    Else
       v = z
    End If
   
    If ((h And 1) = 0) Then
       g = u
    Else
       g = -u
    End If
   
    If ((h And 2) = 0) Then
       g = g + v
    Else
       g = g - v
    End If
   
    grad = g
   
End Function

Private Sub Class_Initialize()
   ReDim p(0 To 511) As Long
   ReDim permutation(0 To 255) As Long
   
   ' What would I do for variable and particularly
   ' array variable initialisers in VB...
   permutation(0) = 151
   permutation(1) = 160
   permutation(2) = 137
   permutation(3) = 91
   permutation(4) = 90
   permutation(5) = 15
   permutation(6) = 131
   permutation(7) = 13
   permutation(8) = 201
   permutation(9) = 95
   permutation(10) = 96
   permutation(11) = 53
   permutation(12) = 194
   permutation(13) = 233
   permutation(14) = 7
   permutation(15) = 225
   permutation(16) = 140
   permutation(17) = 36
   permutation(18) = 103
   permutation(19) = 30
   permutation(20) = 69
   permutation(21) = 142
   permutation(22) = 8
   permutation(23) = 99
   permutation(24) = 37
   permutation(25) = 240
   permutation(26) = 21
   permutation(27) = 10
   permutation(28) = 23
   permutation(29) = 190
   permutation(30) = 6
   permutation(31) = 148
   permutation(32) = 247
   permutation(33) = 120
   permutation(34) = 234
   permutation(35) = 75
   permutation(36) = 0
   permutation(37) = 26
   permutation(38) = 197
   permutation(39) = 62
   permutation(40) = 94
   permutation(41) = 252
   permutation(42) = 219
   permutation(43) = 203
   permutation(44) = 117
   permutation(45) = 35
   permutation(46) = 11
   permutation(47) = 32
   permutation(48) = 57
   permutation(49) = 177
   permutation(50) = 33
   permutation(51) = 88
   permutation(52) = 237
   permutation(53) = 149
   permutation(54) = 56
   permutation(55) = 87
   permutation(56) = 174
   permutation(57) = 20
   permutation(58) = 125
   permutation(59) = 136
   permutation(60) = 171
   permutation(61) = 168
   permutation(62) = 68
   permutation(63) = 175
   permutation(64) = 74
   permutation(65) = 165
   permutation(66) = 71
   permutation(67) = 134
   permutation(68) = 139
   permutation(69) = 48
   permutation(70) = 27
   permutation(71) = 166
   permutation(72) = 77
   permutation(73) = 146
   permutation(74) = 158
   permutation(75) = 231
   permutation(76) = 83
   permutation(77) = 111
   permutation(78) = 229
   permutation(79) = 122
   permutation(80) = 60
   permutation(81) = 211
   permutation(82) = 133
   permutation(83) = 230
   permutation(84) = 220
   permutation(85) = 105
   permutation(86) = 92
   permutation(87) = 41
   permutation(88) = 55
   permutation(89) = 46
   permutation(90) = 245
   permutation(91) = 40
   permutation(92) = 244
   permutation(93) = 102
   permutation(94) = 143
   permutation(95) = 54
   permutation(96) = 65
   permutation(97) = 25
   permutation(98) = 63
   permutation(99) = 161
   permutation(100) = 1
   permutation(101) = 216
   permutation(102) = 80
   permutation(103) = 73
   permutation(104) = 209
   permutation(105) = 76
   permutation(106) = 132
   permutation(107) = 187
   permutation(108) = 208
   permutation(109) = 89
   permutation(110) = 18
   permutation(111) = 169
   permutation(112) = 200
   permutation(113) = 196
   permutation(114) = 135
   permutation(115) = 130
   permutation(116) = 116
   permutation(117) = 188
   permutation(118) = 159
   permutation(119) = 86
   permutation(120) = 164
   permutation(121) = 100
   permutation(122) = 109
   permutation(123) = 198
   permutation(124) = 173
   permutation(125) = 186
   permutation(126) = 3
   permutation(127) = 64
   permutation(128) = 52
   permutation(129) = 217
   permutation(130) = 226
   permutation(131) = 250
   permutation(132) = 124
   permutation(133) = 123
   permutation(134) = 5
   permutation(135) = 202
   permutation(136) = 38
   permutation(137) = 147
   permutation(138) = 118
   permutation(139) = 126
   permutation(140) = 255
   permutation(141) = 82
   permutation(142) = 85
   permutation(143) = 212
   permutation(144) = 207
   permutation(145) = 206
   permutation(146) = 59
   permutation(147) = 227
   permutation(148) = 47
   permutation(149) = 16
   permutation(150) = 58
   permutation(151) = 17
   permutation(152) = 182
   permutation(153) = 189
   permutation(154) = 28
   permutation(155) = 42
   permutation(156) = 223
   permutation(157) = 183
   permutation(158) = 170
   permutation(159) = 213
   permutation(160) = 119
   permutation(161) = 248
   permutation(162) = 152
   permutation(163) = 2
   permutation(164) = 44
   permutation(165) = 154
   permutation(166) = 163
   permutation(167) = 70
   permutation(168) = 221
   permutation(169) = 153
   permutation(170) = 101
   permutation(171) = 155
   permutation(172) = 167
   permutation(173) = 43
   permutation(174) = 172
   permutation(175) = 9
   permutation(176) = 129
   permutation(177) = 22
   permutation(178) = 39
   permutation(179) = 253
   permutation(180) = 19
   permutation(181) = 98
   permutation(182) = 108
   permutation(183) = 110
   permutation(184) = 79
   permutation(185) = 113
   permutation(186) = 224
   permutation(187) = 232
   permutation(188) = 178
   permutation(189) = 185
   permutation(190) = 112
   permutation(191) = 104
   permutation(192) = 218
   permutation(193) = 246
   permutation(194) = 97
   permutation(195) = 228
   permutation(196) = 251
   permutation(197) = 34
   permutation(198) = 242
   permutation(199) = 193
   permutation(200) = 238
   permutation(201) = 210
   permutation(202) = 144
   permutation(203) = 12
   permutation(204) = 191
   permutation(205) = 179
   permutation(206) = 162
   permutation(207) = 241
   permutation(208) = 81
   permutation(209) = 51
   permutation(210) = 145
   permutation(211) = 235
   permutation(212) = 249
   permutation(213) = 14
   permutation(214) = 239
   permutation(215) = 107
   permutation(216) = 49
   permutation(217) = 192
   permutation(218) = 214
   permutation(219) = 31
   permutation(220) = 181
   permutation(221) = 199
   permutation(222) = 106
   permutation(223) = 157
   permutation(224) = 184
   permutation(225) = 84
   permutation(226) = 204
   permutation(227) = 176
   permutation(228) = 115
   permutation(229) = 121
   permutation(230) = 50
   permutation(231) = 45
   permutation(232) = 127
   permutation(233) = 4
   permutation(234) = 150
   permutation(235) = 254
   permutation(236) = 138
   permutation(237) = 236
   permutation(238) = 205
   permutation(239) = 93
   permutation(240) = 222
   permutation(241) = 114
   permutation(242) = 67
   permutation(243) = 29
   permutation(244) = 24
   permutation(245) = 72
   permutation(246) = 243
   permutation(247) = 141
   permutation(248) = 128
   permutation(249) = 195
   permutation(250) = 78
   permutation(251) = 66
   permutation(252) = 215
   permutation(253) = 61
   permutation(254) = 156
   permutation(255) = 180
      
   Dim i As Long
   For i = 0 To 255
      p(i) = permutation(i)
      p(256 + i) = p(i)
   Next i
   
End Sub

