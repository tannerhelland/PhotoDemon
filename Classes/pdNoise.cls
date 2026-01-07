VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Noise Generator
'Copyright 2017-2026 by Tanner Helland
'Created: 17/October/17
'Last updated: 17/October/17
'Last update: migrate noise-related functions from elsewhere in the project, add some new functions from
'             external experimental sub-projects
'
'Per its name, this class provides a simple noise-generation interface.  Many image processing functions rely
' on (repeatable) noise generation behavior, and after significant dissatisfaction with 3rd-party VB6 code
' in this arena, I've written my own implementations.
'
'Perlin, Simplex, and OpenSimplex noise engines are currently available.  Per PD's mission statement,
' only 2D implementations are available at present.  3D implementations could easily be added, but there's
' not much use for them at present (and 3D implementations are significantly slower).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Perlin, Simplex, and OpenSimplex engines use the same permutation table, so this class will always initialize
' a default table.  Take this into consideration if you are rapidly creating/destroying class instances.
Private m_Permutation() As Byte

'Simplex noise requires a number of specialized constants, including some messy vector arrays (which have
' no obvious corollary in VB6).
Private m_PermMod12() As Byte
Private Const f2 As Double = 0.366025404        '0.5*(Math.sqrt(3.0)-1.0)
Private Const g2 As Double = 0.211324865        '(3.0-Math.sqrt(3.0))/6.0

Private Type SmpVec3
    x As Double
    y As Double
    z As Double
End Type

Private m_Grad3() As SmpVec3

'Open-simplex noise also uses gradients, but they use indices based on octagonal distance
Private m_Grad2() As Integer
Private Const STRETCH_CONSTANT_2D As Double = -0.211324865405187    '(1 / Sqr(2 + 1) - 1) / 2
Private Const SQUISH_CONSTANT_2D  As Double = 0.366025403784439     '(Sqr(2 + 1) - 1) / 2
Private Const INV_NORM_CONSTANT_2D As Double = 1# / 47#

'In PD, we don't need full-on 3D noise.  Instead, we create "random" 2D noise by using a fixed 2D Perlin generator
' (basically, a 3D generator with "z" always forced to zero).  For "random" noise, add fixed, random offsets to the
' x/y values supplied to this function - that will "slide" the noise map to a new position on the infinite
' 2D Perlin grid.
Friend Function PerlinNoise2d(ByVal xIn As Double, ByVal yIn As Double) As Double

    'Find the unit square containing the incoming point
    Dim xHash As Long, yHash As Long
    xHash = Int(xIn) And 255
    yHash = Int(yIn) And 255
    
    'Find the relative (x, y) point in the square
    xIn = xIn - Int(xIn)
    yIn = yIn - Int(yIn)
    
    'Compute fade curves for the x and y directions
    Dim u As Double, v As Double
    
    'I've manually in-lined fast function variants for improved performance
    u = xIn * xIn * (3# - 2# * xIn)     'u = PerlinFadeFast(xIn)
    v = yIn * yIn * (3# - 2# * yIn)     'v = PerlinFadeFast(yIn)
    
    'Hash corner coordinates
    Dim a As Long, aa As Long, ab As Long, b As Long, ba As Long, bb As Long
    a = m_Permutation(xHash) + yHash
    aa = m_Permutation(a)
    ab = m_Permutation(a + 1)
    b = m_Permutation(xHash + 1) + yHash
    ba = m_Permutation(b)
    bb = m_Permutation(b + 1)
    
    'Blend the corners
    PerlinNoise2d = PerlinLerp(v, PerlinLerp(u, PerlinGrad(m_Permutation(aa), xIn, yIn, 0#), PerlinGrad(m_Permutation(ba), xIn - 1#, yIn, 0#)), PerlinLerp(u, PerlinGrad(m_Permutation(ab), xIn, yIn - 1#, 0#), PerlinGrad(m_Permutation(bb), xIn - 1#, yIn - 1#, 0#)))
    
End Function

Private Function PerlinLerp(ByVal t As Double, ByVal a As Double, ByVal b As Double) As Double
    PerlinLerp = a + t * (b - a)
End Function

Private Function PerlinGrad(ByVal hash As Long, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double

    'Basically, the point of this function is to convert the 4 lo-bits of the incoming hash code
    ' into one of 12 possible gradient directions.  (One direction for each edge of a 3D cube.)
    Dim h As Long
    h = hash And 15

    Dim u As Double, v As Double
    If (h < 8) Then u = x Else u = y
    If (h < 4) Then v = y Else If (h = 12) Or (h = 14) Then v = x Else v = z
    If (h And 1) = 0 Then
        If (h And 2) = 0 Then PerlinGrad = u + v Else PerlinGrad = u - v
    Else
        If (h And 2) = 0 Then PerlinGrad = -u + v Else PerlinGrad = -u - v
    End If

End Function

'2D simplex noise, translated from a Java original at this link (good as of Oct 2017):
' http://weber.itn.liu.se/~stegu/simplexnoise/simplexnoise.pdf
Friend Function SimplexNoise2d(ByVal xIn As Double, ByVal yIn As Double) As Double
    
    'Skew the input space to determine which simplex cell we're in
    Dim s As Double
    s = (xIn + yIn) * f2
    
    Dim i As Long, j As Long
    i = Int(xIn + s)
    j = Int(yIn + s)
    
    Dim t As Double
    t = (i + j) * g2
    
    'Unskew the cell origin back to (x, y) space, then calculate (x, y) distances from the cell origin
    Dim x0 As Double, y0 As Double
    x0 = xIn - (i - t)
    y0 = yIn - (j - t)
    
    'For the 2D case, the simplex shape is an equilateral triangle. Determine which simplex we are in.
    
    'Offsets for second (middle) corner of simplex in (i,j) coords
    Dim i1 As Long, j1 As Long
    
    'Lower triangle, XY order: (0,0)->(1,0)->(1,1)
    If (x0 > y0) Then
        i1 = 1
        j1 = 0
    
    'Upper triangle, YX order: (0,0)->(0,1)->(1,1)
    Else
        i1 = 0
        j1 = 1
    End If
    
    'A step of (1,0) in (i,j) means a step of (1-c,-c) in (x,y), and a step of (0,1) in (i,j)
    ' means a step of (-c,1-c) in (x,y), where c = (3-sqrt(3))/6
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    
    'Offsets for middle corner in (x,y) unskewed coords
    x1 = x0 - i1 + g2
    y1 = y0 - j1 + g2
    
    'Offsets for last corner in (x,y) unskewed coords
    x2 = x0 - 1# + 2# * g2
    y2 = y0 - 1# + 2# * g2

    'Work out the hashed gradient indices of the three simplex corners
    Dim ii As Long, jj As Long
    ii = i And 255
    jj = j And 255
    
    Dim gi0 As Long, gi1 As Long, gi2 As Long
    gi0 = m_PermMod12(ii + m_Permutation(jj))
    gi1 = m_PermMod12(ii + i1 + m_Permutation(jj + j1))
    gi2 = m_PermMod12(ii + 1 + m_Permutation(jj + 1))
    
    'Calculate the contribution from the three corners
    Dim t0 As Double, t1 As Double, t2 As Double
    t0 = 0.5 - x0 * x0 - y0 * y0
    
    Dim n0 As Double, n1 As Double, n2 As Double
    
    If (t0 < 0#) Then
        n0 = 0#
    Else
        t0 = t0 * t0
        
        '(x,y) of grad3 used for 2D gradient, manually inlined for performance
        'n0 = t0 * t0 * Dot2(gi0, x0, y0)
        n0 = t0 * t0 * (m_Grad3(gi0).x * x0 + m_Grad3(gi0).y * y0)
    End If
    
    t1 = 0.5 - x1 * x1 - y1 * y1
    If (t1 < 0#) Then
        n1 = 0#
    Else
        t1 = t1 * t1
        'n1 = t1 * t1 * Dot2(gi1, x1, y1)
        n1 = t1 * t1 * (m_Grad3(gi1).x * x1 + m_Grad3(gi1).y * y1)
    End If
    
    t2 = 0.5 - x2 * x2 - y2 * y2
    If (t2 < 0#) Then
        n2 = 0#
    Else
        t2 = t2 * t2
        'n2 = t2 * t2 * Dot2(gi2, x2, y2)
        n2 = t2 * t2 * (m_Grad3(gi2).x * x2 + m_Grad3(gi2).y * y2)
    End If
    
    'Add contributions from each corner to get the final noise value.
    ' (The result is scaled to return values in the interval [-1,1].)
    SimplexNoise2d = 70# * (n0 + n1 + n2)
  
End Function

'2D Open Simplex noise, translated from the Java original at this link (good as of Oct 2017):
' https://gist.github.com/KdotJPG/b1270127455a94ac5d19
Friend Function OpenSimplexNoise2d(ByVal x As Double, ByVal y As Double) As Double

    'Place input coordinates onto grid
    Dim stretchOffset As Double
    stretchOffset = (x + y) * STRETCH_CONSTANT_2D
    Dim xs As Double, ys As Double
    xs = x + stretchOffset
    ys = y + stretchOffset
    
    'Floor to get grid coordinates of rhombus (stretched square) super-cell origin
    Dim xsb As Long, ysb As Long
    xsb = Int(xs)
    ysb = Int(ys)
    
    'Skew out to get actual coordinates of rhombus origin. We'll need these later.
    Dim squishOffset As Double
    squishOffset = (xsb + ysb) * SQUISH_CONSTANT_2D
    Dim xb As Double, yb As Double
    xb = xsb + squishOffset
    yb = ysb + squishOffset
    
    'Compute grid coordinates relative to rhombus origin.
    Dim xins As Double, yins As Double
    xins = xs - xsb
    yins = ys - ysb
    
    'Sum those together to get a value that determines which region we're in.
    Dim inSum As Double
    inSum = xins + yins

    'Positions relative to origin point.
    Dim dx0 As Double, dy0 As Double
    dx0 = x - xb
    dy0 = y - yb
    
    'We'll be defining these inside the next block and using them afterwards.
    Dim dx_ext As Double, dy_ext As Double
    Dim xsv_ext As Long, ysv_ext As Long
    Dim osValue As Double
    
    'Contribution (1,0)
    Dim dx1 As Double, dy1 As Double, attn1 As Double
    dx1 = dx0 - 1 - SQUISH_CONSTANT_2D
    dy1 = dy0 - 0 - SQUISH_CONSTANT_2D
    attn1 = 2# - dx1 * dx1 - dy1 * dy1
    If (attn1 > 0#) Then
        attn1 = attn1 * attn1
        osValue = osValue + attn1 * attn1 * OpenSimplexExtrapolate(xsb + 1, ysb, dx1, dy1)
    End If

    'Contribution (0,1)
    Dim dx2 As Double, dy2 As Double, attn2 As Double
    dx2 = dx0 - 0 - SQUISH_CONSTANT_2D
    dy2 = dy0 - 1 - SQUISH_CONSTANT_2D
    attn2 = 2# - dx2 * dx2 - dy2 * dy2
    If (attn2 > 0#) Then
        attn2 = attn2 * attn2
        osValue = osValue + attn2 * attn2 * OpenSimplexExtrapolate(xsb, ysb + 1, dx2, dy2)
    End If
    
    Dim zins As Double
    
    '//We're inside the triangle (2-Simplex) at (0,0)
    If (inSum <= 1) Then
        
        zins = 1# - inSum
        
        '(0,0) is one of the closest two triangular vertices
        If (zins > xins) Or (zins > yins) Then
        
            If (xins > yins) Then
                xsv_ext = xsb + 1
                ysv_ext = ysb - 1
                dx_ext = dx0 - 1
                dy_ext = dy0 + 1
            Else
                xsv_ext = xsb - 1
                ysv_ext = ysb + 1
                dx_ext = dx0 + 1
                dy_ext = dy0 - 1
            End If
        
        '(1,0) and (0,1) are the closest two vertices.
        Else
            xsv_ext = xsb + 1
            ysv_ext = ysb + 1
            dx_ext = dx0 - 1# - 2# * SQUISH_CONSTANT_2D
            dy_ext = dy0 - 1# - 2# * SQUISH_CONSTANT_2D
        End If
    
    'We're inside the triangle (2-Simplex) at (1,1)
    Else
    
        zins = 2 - inSum
        
        '(0,0) is one of the closest two triangular vertices
        If (zins < xins) Or (zins < yins) Then
            If (xins > yins) Then
                xsv_ext = xsb + 2
                ysv_ext = ysb + 0
                dx_ext = dx0 - 2# - 2# * SQUISH_CONSTANT_2D
                dy_ext = dy0 - 2# * SQUISH_CONSTANT_2D
            Else
                xsv_ext = xsb + 0
                ysv_ext = ysb + 2
                dx_ext = dx0 - 2# * SQUISH_CONSTANT_2D
                dy_ext = dy0 - 2# - 2# * SQUISH_CONSTANT_2D
            End If
        
        '(1,0) and (0,1) are the closest two vertices.
        Else
            dx_ext = dx0
            dy_ext = dy0
            xsv_ext = xsb
            ysv_ext = ysb
        End If
        
        xsb = xsb + 1
        ysb = ysb + 1
        dx0 = dx0 - 1# - 2# * SQUISH_CONSTANT_2D
        dy0 = dy0 - 1# - 2# * SQUISH_CONSTANT_2D
    
    End If
    
    'Contribution (0,0) or (1,1)
    Dim attn0 As Double
    attn0 = 2# - dx0 * dx0 - dy0 * dy0
    If (attn0 > 0#) Then
        attn0 = attn0 * attn0
        osValue = osValue + attn0 * attn0 * OpenSimplexExtrapolate(xsb, ysb, dx0, dy0)
    End If
    
    'Extra Vertex
    Dim attn_ext As Double
    attn_ext = 2# - dx_ext * dx_ext - dy_ext * dy_ext
    If (attn_ext > 0#) Then
        attn_ext = attn_ext * attn_ext
        osValue = osValue + attn_ext * attn_ext * OpenSimplexExtrapolate(xsv_ext, ysv_ext, dx_ext, dy_ext)
    End If
    
    OpenSimplexNoise2d = osValue * INV_NORM_CONSTANT_2D
    
End Function

Private Function OpenSimplexExtrapolate(ByVal xsb As Long, ByVal ysb As Long, ByVal dx As Double, ByVal dy As Double) As Double
    Dim gIndex As Long
    gIndex = m_Permutation((m_Permutation(xsb And 255) + ysb) And 255) And &HE&
    OpenSimplexExtrapolate = m_Grad2(gIndex) * dx + m_Grad2(gIndex + 1) * dy
End Function

Private Sub Class_Initialize()
    
    'Perform the (very ugly) required permutation init
    ReDim m_Permutation(0 To 511) As Byte
    m_Permutation(0) = 151:   m_Permutation(1) = 160:   m_Permutation(2) = 137:   m_Permutation(3) = 91
    m_Permutation(4) = 90:    m_Permutation(5) = 15:    m_Permutation(6) = 131:   m_Permutation(7) = 13
    m_Permutation(8) = 201:   m_Permutation(9) = 95:    m_Permutation(10) = 96:   m_Permutation(11) = 53
    m_Permutation(12) = 194:  m_Permutation(13) = 233:  m_Permutation(14) = 7:    m_Permutation(15) = 225
    m_Permutation(16) = 140:  m_Permutation(17) = 36:   m_Permutation(18) = 103:  m_Permutation(19) = 30
    m_Permutation(20) = 69:   m_Permutation(21) = 142:  m_Permutation(22) = 8:    m_Permutation(23) = 99
    m_Permutation(24) = 37:   m_Permutation(25) = 240:  m_Permutation(26) = 21:   m_Permutation(27) = 10
    m_Permutation(28) = 23:   m_Permutation(29) = 190:  m_Permutation(30) = 6:    m_Permutation(31) = 148
    m_Permutation(32) = 247:  m_Permutation(33) = 120:  m_Permutation(34) = 234:  m_Permutation(35) = 75:
    m_Permutation(36) = 0:    m_Permutation(37) = 26:   m_Permutation(38) = 197:  m_Permutation(39) = 62:
    m_Permutation(40) = 94:   m_Permutation(41) = 252:  m_Permutation(42) = 219:  m_Permutation(43) = 203:
    m_Permutation(44) = 117:  m_Permutation(45) = 35:   m_Permutation(46) = 11:   m_Permutation(47) = 32:
    m_Permutation(48) = 57:   m_Permutation(49) = 177:  m_Permutation(50) = 33:   m_Permutation(51) = 88:
    m_Permutation(52) = 237:  m_Permutation(53) = 149:  m_Permutation(54) = 56:   m_Permutation(55) = 87:
    m_Permutation(56) = 174:  m_Permutation(57) = 20:   m_Permutation(58) = 125:  m_Permutation(59) = 136:
    m_Permutation(60) = 171:  m_Permutation(61) = 168:  m_Permutation(62) = 68:   m_Permutation(63) = 175:
    m_Permutation(64) = 74:   m_Permutation(65) = 165:  m_Permutation(66) = 71:   m_Permutation(67) = 134:
    m_Permutation(68) = 139:  m_Permutation(69) = 48:   m_Permutation(70) = 27:   m_Permutation(71) = 166:
    m_Permutation(72) = 77:   m_Permutation(73) = 146:  m_Permutation(74) = 158:  m_Permutation(75) = 231:
    m_Permutation(76) = 83:   m_Permutation(77) = 111:  m_Permutation(78) = 229:  m_Permutation(79) = 122:
    m_Permutation(80) = 60:   m_Permutation(81) = 211:  m_Permutation(82) = 133:  m_Permutation(83) = 230:
    m_Permutation(84) = 220:  m_Permutation(85) = 105:  m_Permutation(86) = 92:   m_Permutation(87) = 41:
    m_Permutation(88) = 55:   m_Permutation(89) = 46:   m_Permutation(90) = 245:  m_Permutation(91) = 40:
    m_Permutation(92) = 244:  m_Permutation(93) = 102:  m_Permutation(94) = 143:  m_Permutation(95) = 54:
    m_Permutation(96) = 65:   m_Permutation(97) = 25:   m_Permutation(98) = 63:   m_Permutation(99) = 161:
    m_Permutation(100) = 1:   m_Permutation(101) = 216: m_Permutation(102) = 80:  m_Permutation(103) = 73:
    m_Permutation(104) = 209: m_Permutation(105) = 76:  m_Permutation(106) = 132: m_Permutation(107) = 187:
    m_Permutation(108) = 208: m_Permutation(109) = 89:  m_Permutation(110) = 18:  m_Permutation(111) = 169:
    m_Permutation(112) = 200: m_Permutation(113) = 196: m_Permutation(114) = 135: m_Permutation(115) = 130:
    m_Permutation(116) = 116: m_Permutation(117) = 188: m_Permutation(118) = 159: m_Permutation(119) = 86:
    m_Permutation(120) = 164: m_Permutation(121) = 100: m_Permutation(122) = 109: m_Permutation(123) = 198:
    m_Permutation(124) = 173: m_Permutation(125) = 186: m_Permutation(126) = 3:   m_Permutation(127) = 64:
    m_Permutation(128) = 52:  m_Permutation(129) = 217: m_Permutation(130) = 226: m_Permutation(131) = 250:
    m_Permutation(132) = 124: m_Permutation(133) = 123: m_Permutation(134) = 5:   m_Permutation(135) = 202:
    m_Permutation(136) = 38:  m_Permutation(137) = 147: m_Permutation(138) = 118: m_Permutation(139) = 126:
    m_Permutation(140) = 255: m_Permutation(141) = 82:  m_Permutation(142) = 85:  m_Permutation(143) = 212:
    m_Permutation(144) = 207: m_Permutation(145) = 206: m_Permutation(146) = 59:  m_Permutation(147) = 227:
    m_Permutation(148) = 47:  m_Permutation(149) = 16:  m_Permutation(150) = 58:  m_Permutation(151) = 17:
    m_Permutation(152) = 182: m_Permutation(153) = 189: m_Permutation(154) = 28:  m_Permutation(155) = 42:
    m_Permutation(156) = 223: m_Permutation(157) = 183: m_Permutation(158) = 170: m_Permutation(159) = 213:
    m_Permutation(160) = 119: m_Permutation(161) = 248: m_Permutation(162) = 152: m_Permutation(163) = 2:
    m_Permutation(164) = 44:  m_Permutation(165) = 154: m_Permutation(166) = 163: m_Permutation(167) = 70:
    m_Permutation(168) = 221: m_Permutation(169) = 153: m_Permutation(170) = 101: m_Permutation(171) = 155:
    m_Permutation(172) = 167: m_Permutation(173) = 43:  m_Permutation(174) = 172: m_Permutation(175) = 9:
    m_Permutation(176) = 129: m_Permutation(177) = 22:  m_Permutation(178) = 39:  m_Permutation(179) = 253:
    m_Permutation(180) = 19:  m_Permutation(181) = 98:  m_Permutation(182) = 108: m_Permutation(183) = 110:
    m_Permutation(184) = 79:  m_Permutation(185) = 113: m_Permutation(186) = 224: m_Permutation(187) = 232:
    m_Permutation(188) = 178: m_Permutation(189) = 185: m_Permutation(190) = 112: m_Permutation(191) = 104:
    m_Permutation(192) = 218: m_Permutation(193) = 246: m_Permutation(194) = 97:  m_Permutation(195) = 228:
    m_Permutation(196) = 251: m_Permutation(197) = 34:  m_Permutation(198) = 242: m_Permutation(199) = 193:
    m_Permutation(200) = 238: m_Permutation(201) = 210: m_Permutation(202) = 144: m_Permutation(203) = 12:
    m_Permutation(204) = 191: m_Permutation(205) = 179: m_Permutation(206) = 162: m_Permutation(207) = 241:
    m_Permutation(208) = 81:  m_Permutation(209) = 51:  m_Permutation(210) = 145: m_Permutation(211) = 235:
    m_Permutation(212) = 249: m_Permutation(213) = 14:  m_Permutation(214) = 239: m_Permutation(215) = 107:
    m_Permutation(216) = 49:  m_Permutation(217) = 192: m_Permutation(218) = 214: m_Permutation(219) = 31:
    m_Permutation(220) = 181: m_Permutation(221) = 199: m_Permutation(222) = 106: m_Permutation(223) = 157:
    m_Permutation(224) = 184: m_Permutation(225) = 84:  m_Permutation(226) = 204: m_Permutation(227) = 176:
    m_Permutation(228) = 115: m_Permutation(229) = 121: m_Permutation(230) = 50:  m_Permutation(231) = 45:
    m_Permutation(232) = 127: m_Permutation(233) = 4:   m_Permutation(234) = 150: m_Permutation(235) = 254:
    m_Permutation(236) = 138: m_Permutation(237) = 236: m_Permutation(238) = 205: m_Permutation(239) = 93:
    m_Permutation(240) = 222: m_Permutation(241) = 114: m_Permutation(242) = 67:  m_Permutation(243) = 29:
    m_Permutation(244) = 24:  m_Permutation(245) = 72:  m_Permutation(246) = 243: m_Permutation(247) = 141:
    m_Permutation(248) = 128: m_Permutation(249) = 195: m_Permutation(250) = 78:  m_Permutation(251) = 66:
    m_Permutation(252) = 215: m_Permutation(253) = 61:  m_Permutation(254) = 156: m_Permutation(255) = 180
    
    'Whew!  That sucked.  To avoid wraparound requirements, we double the size of our actual stored table.
    Dim i As Long
    For i = 0 To 255
        m_Permutation(i + 256) = m_Permutation(i)
    Next i
    
    'Simplex noise also uses a "% 12" table to index into a predetermined gradient array.
    ' Mod is expensive, so we pre-cache that table as well.  (You could look at conditionally creating
    ' this table, only if Simplex noise is required.)
    ReDim m_PermMod12(0 To 511) As Byte
    For i = 0 To 511
        m_PermMod12(i) = m_Permutation(i) Mod 12
    Next i
    
    'Initialize 3D Simplex vector tables
    ReDim m_Grad3(0 To 11) As SmpVec3
    m_Grad3(0).x = 1:  m_Grad3(0).y = 1:   m_Grad3(0).z = 0
    m_Grad3(1).x = -1: m_Grad3(1).y = 1:   m_Grad3(1).z = 0
    m_Grad3(2).x = 1:  m_Grad3(2).y = -1:  m_Grad3(2).z = 0
    m_Grad3(3).x = -1: m_Grad3(3).y = -1:  m_Grad3(3).z = 0
    m_Grad3(4).x = 1:  m_Grad3(4).y = 0:   m_Grad3(4).z = 1
    m_Grad3(5).x = -1: m_Grad3(5).y = 0:   m_Grad3(5).z = 1
    m_Grad3(6).x = 1:  m_Grad3(6).y = 0:   m_Grad3(6).z = -1
    m_Grad3(7).x = -1: m_Grad3(7).y = 0:   m_Grad3(7).z = -1
    m_Grad3(8).x = 0:  m_Grad3(8).y = 1:   m_Grad3(8).z = 1
    m_Grad3(9).x = 0:  m_Grad3(9).y = -1:  m_Grad3(9).z = 1
    m_Grad3(10).x = 0: m_Grad3(10).y = 1:  m_Grad3(10).z = -1
    m_Grad3(11).x = 0: m_Grad3(11).y = -1: m_Grad3(11).z = -1
    
    'Initialize 2D Open-Simplex gradient table
    ReDim m_Grad2(0 To 15) As Integer
    m_Grad2(0) = 5: m_Grad2(1) = 2: m_Grad2(2) = 2: m_Grad2(3) = 5
    m_Grad2(4) = -5: m_Grad2(5) = 2: m_Grad2(6) = -2: m_Grad2(7) = 5
    m_Grad2(8) = 5: m_Grad2(9) = -2: m_Grad2(10) = 2: m_Grad2(11) = -5
    m_Grad2(12) = -5: m_Grad2(13) = -2: m_Grad2(14) = -2: m_Grad2(15) = -5
    
End Sub
