Attribute VB_Name = "Outside_mGfxWu"
'Note: this file has been modified for use within PhotoDemon.

'This module contains a modified version of Ron van Tilburg's original "GfxPrim" code.

'IF YOU WANT TO USE THIS CODE IN YOUR OWN PROJECT, PLEASE DOWNLOAD THE ORIGINAL, UNMODIFIED VERSION FROM THIS LINK:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71370&lngWId=1

'Many thanks to Ron for his native-VB implementation of Xiaolin Wu's antialising algorithm.
' If you would like to read more about Wu's original implementation, please visit:
' http://en.wikipedia.org/wiki/Xiaolin_Wu%27s_line_algorithm

'Original header comment for this module:

'mGfxWu.bas ================================================================================ ©RVT 2008
'Drawing a line using Wu's algorithm
'===========================================================================================================
'Wu invented a very simple way of anti-aliasing a line, by determining the differences between a line and
'its plotted point. He then blends pixels based on their distance from the ideal line, fairly simply
'This is fairly fast and looks good - it uses two pixels to represent a "blurred" version of the line
'the eye sees this as quite nice as it is fooled into not seeing jaggies
'===========================================================================================================
'===========================================================================================================

Option Explicit

Private Type UDT_LONG    'a packing type for a long so that we can use LSET
  v As Long
End Type

Private Type RGBA       'The RGB type for VB colours  xBGR as a long
  r As Byte
  g As Byte
  b As Byte
  a As Byte
End Type

#Const DDEBUG = 0   'Set this to 1 to test octant drawing of Circles

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Const MAX_WEIGHT As Long = 256&    'This is big enough because all colour vals can only range 0,,255

'==========================================================================================================
'== a fairly critical piece in all is this routine = the most expensive operation in the algorithms
'== ideally this would be asm for max speed or directly writing into a DIB (beyond me here (i look forward
'to someones solution for this as I expect it could be quite quick for small DIB areas, however true cost
'must include GET/SET DIB for overall gain/loss against these routines))
'SET/GET Pixel are not that fast from here and are the limiting operations (this is where VB and GDI+ really
'win) when MAX_WEIGHT is a power of 2 some code optimisation occurs by the compiler
'==========================================================================================================
'Replace existing colour with mix of old and new colours using weight w (0<w<MAX_WEIGHT) for new colour


'=============================== USING 65535 as MAX_WEIGHT ====================================================
'==============================================================================================================

Public Sub BlendPixelWu(ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal Colour As Long, _
                        ByVal Weight As Long)

  Dim pRGB As RGBA, iRGB As RGBA, c As UDT_LONG, cw As Long

  If Weight <> 0 Then                                                 'oh dear we need to do something
    cw = 65535 - Weight                                               'complement of weight
    If cw <> 0 Then                                                   'otherwise its a solid pixel of INK
      c.v = Colour: LSet iRGB = c                                    'so calculate a blended colour
      c.v = GetPixel(hDC, x, y): LSet pRGB = c
      Call SetPixelV(hDC, x, y, _
                     RGB((Weight * iRGB.r + cw * pRGB.r) \ 65536, _
                     (Weight * iRGB.g + cw * pRGB.g) \ 65536, _
                     (Weight * iRGB.b + cw * pRGB.b) \ 65536))    'and finally replace Ink with it
    Else
      Call SetPixelV(hDC, x, y, Colour)                               'And plot it
    End If
    'else                                                               Weight=0 so dont plot anything
  End If

End Sub

'=============================================================================================================
'=============================================================================================================

Public Sub DrawLineWu(ByVal hDC As Long, _
                      ByVal x1 As Long, ByVal y1 As Long, _
                      ByVal x2 As Long, ByVal y2 As Long, _
                      ByVal Colour As Long)

 'Wu Lines are normally given in real dimensions but here we assume that we are actually working in Integer
 'coords so that we know the endpoints will always drawn and be full colour.
 'We take the view that a line of width 1 in vertical or horz directions is drawn that way up slopes too
 'This makes the lines look thinner than h or v but is usually good enough
 'All lines are actually use 2 pixels unless horz or vert

  Dim xd As Long, yd As Long, t As Long, grad As Long
  Dim xi As Long, yi As Long

  xd = x2 - x1                                    'Width and Height of the line
  yd = y2 - y1

  If Abs(xd) > Abs(yd) Then                       'check line gradient => horizontal(ish) lines
    If x1 > x2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    If yd = 0 Then                                'A Horizontal Line
      For xd = x1 To x2
        Call SetPixelV(hDC, xd, y1, Colour)
      Next xd
    Else
      grad = (yd * 65536 + xd \ 2) \ xd           'rounded scaled gradient of the line 0..1 => 0..65536
      yi = y1 * 65536                             'first y coord offset so that line is centred
      For xi = x1 To x2
        Call SetPixelV(hDC, xi, yi \ 65536, Colour)
        yi = yi + grad
      Next xi
    End If

  Else                                            'vertical(ish) lines

    If y1 > y2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    If xd = 0 Then                                ' A Vertical Line
      For yd = y1 To y2
        Call SetPixelV(hDC, x1, yd, Colour)
      Next yd
    Else
      grad = (xd * 65536 + yd \ 2) \ yd           'rounded scaled gradient of the line 0..1 => 0..65536
      xi = x1 * 65536                             'first x coord offset so that line is centred
      For yi = y1 To y2
        Call SetPixelV(hDC, xi \ 65536, yi, Colour)
        xi = xi + grad
      Next yi
    End If

  End If

End Sub

'===========================================================================================================
'Wu Lines are normally given in real dimensions but here we assume that we are actually working in Integer
'coords so that we know the endpoints will always drawn and be full colour.
'We take the view that a line of width 1 in vertical or horz directions is drawn that way up slopes too
'This makes the lines look thicker than h or v but is usually good enough, (strictly we should plot with
'a lighter blend weight for sloped lines). All lines are actually use 2 pixels unless horz or vert or +-45degrees
'===========================================================================================================

Public Sub DrawLineWuAA(ByVal hDC As Long, _
                        ByVal x1 As Long, ByVal y1 As Long, _
                        ByVal x2 As Long, ByVal y2 As Long, _
                        ByVal Colour As Long)

  Dim xd As Long, yd As Long, t As Long, grad As Long
  Dim xi As Long, yi As Long, xf As Long, yf As Long, w As Long

  xd = x2 - x1                                    'Width and Height of the line
  yd = y2 - y1

  If Abs(xd) > Abs(yd) Then                       'check line gradient => horizontal(ish) lines
    If x1 > x2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    If yd = 0 Then                                'A Horizontal Line (never needs antialiasing)
      For xd = x1 To x2
        Call SetPixelV(hDC, xd, y1, Colour)
      Next xd
    Else
      grad = (yd * 65536 + xd \ 2) \ xd           'rounded scaled gradient of the line 0..1 => 0..65536
      yf = y1 * 65536                             'first y coord offset so that line is centred
      For xi = x1 To x2
        yi = yf \ 65536
        w = yf And &HFFFF&
        Call BlendPixelWu(hDC, xi, yi, Colour, 65535 - w)
        Call BlendPixelWu(hDC, xi, yi + 1, Colour, w)
        yf = yf + grad
      Next xi
    End If
  Else                                            'vertical(ish) lines

    If y1 > y2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    If xd = 0 Then                                ' A Vertical Line (never needs antialiasing)
      For yd = y1 To y2
        Call SetPixelV(hDC, x1, yd, Colour)
      Next yd
    Else
      grad = (xd * 65536 + yd \ 2) \ yd           'rounded scaled gradient of the line 0..1 => 0..65536
      xf = x1 * 65536                             'first x coord offset so that line is centred
      For yi = y1 To y2
        xi = xf \ 65536
        w = xf And &HFFFF&
        Call BlendPixelWu(hDC, xi, yi, Colour, 65535 - w)
        Call BlendPixelWu(hDC, xi + 1, yi, Colour, w)
        xf = xf + grad
      Next yi
    End If
  End If

End Sub

'BONUS:
'===========================================================================================================
'============================ WU ALGORITHM FOR WIDER LINES =================================================
'===========================================================================================================

Public Sub DrawLineAAV(ByVal hDC As Long, _
                       ByVal x1 As Long, ByVal y1 As Long, _
                       ByVal x2 As Long, ByVal y2 As Long, _
                       ByVal Colour As Long, _
                       Optional ByVal PenWidth As Long = 1)    'try 1 to 5 at most

 'All lines actually use Penwidth+1 pixels unless horz or vert or 45 deg

  Dim xd As Long, yd As Long, t As Long, grad As Long, pw As Long
  Dim xi As Long, yi As Long, xf As Long, yf As Long, w As Long

  xd = x2 - x1                                    'Width and Height of the line
  yd = y2 - y1

  If Abs(xd) > Abs(yd) Then                       'check line gradient => horizontal(ish) lines
    If x1 > x2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    grad = (yd * 65536 + xd \ 2) \ xd             'rounded scaled gradient of the line 0..1 => 0..65536
    yi = y1 - PenWidth \ 2                        'first y coord offset so that line is centred
    yf = yi * 65536

    For xi = x1 To x2
      yi = yf \ 65536
      w = yf And &HFFFF&

      Call BlendPixelWu(hDC, xi, yi, Colour, 65535 - w)
      pw = PenWidth - 1
      Do While pw > 0
        Call SetPixelV(hDC, xi, yi + pw, Colour)  'middles are solid if width>1
        pw = pw - 1
      Loop
      Call BlendPixelWu(hDC, xi, yi + PenWidth, Colour, w)

      yf = yf + grad
    Next xi

  Else                                            'vertical(ish) lines

    If y1 > y2 Then                               'if line is back to front
      t = x1: x1 = x2: x2 = t                     'then swap it round
      t = y1: y1 = y2: y2 = t
      xd = x2 - x1                                'and recalc xd & yd
      yd = y2 - y1
    End If

    grad = (xd * 65536 + yd \ 2) \ yd             'rounded scaled gradient of the line 0..1 => 0..65536
    xi = x1 - PenWidth \ 2                        'first x coord offset so that line is centred
    xf = xi * 65536

    For yi = y1 To y2
      xi = xf \ 65536
      w = xf And &HFFFF&

      Call BlendPixelWu(hDC, xi, yi, Colour, 65535 - w)
      pw = PenWidth - 1
      Do While pw > 0
        Call SetPixelV(hDC, xi + pw, yi, Colour)  'middles are solid if width>1
        pw = pw - 1
      Loop
      Call BlendPixelWu(hDC, xi + PenWidth, yi, Colour, w)

      xf = xf + grad
    Next yi

  End If

End Sub

Private Sub BlendPixel(ByVal hDC As Long, _
                      ByVal x As Long, ByVal y As Long, _
                      ByVal Colour As Long, _
                      ByVal Weight As Long)

  Dim pRGB As RGBA, iRGB As RGBA, cw As Long, c As UDT_LONG

  If Weight <> 0 Then                                                 'oh dear we need to do something
    cw = MAX_WEIGHT - Weight                                          'complement of weight
    If cw <> 0 Then                                                   'otherwise its a solid pixel of INK
      c.v = Colour: LSet iRGB = c                                     'so calculate a blended colour
      c.v = GetPixel(hDC, x, y): LSet pRGB = c
      Call SetPixelV(hDC, x, y, _
                     RGB((Weight * iRGB.r + cw * pRGB.r) \ MAX_WEIGHT, _
                         (Weight * iRGB.g + cw * pRGB.g) \ MAX_WEIGHT, _
                         (Weight * iRGB.b + cw * pRGB.b) \ MAX_WEIGHT))
                                                                      'and finally Plot it BLENDED
    Else
      Call SetPixelV(hDC, x, y, Colour)                               'And plot it SOLID
    End If
    'else                                                             'Weight=0 so dont plot anything
  End If

End Sub

Private Sub SetPixel(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Colour As Long)

  Call SetPixelV(hDC, x, y, Colour)

End Sub

Private Function AAColour(ByVal Colour As Long, _
                         ByVal dy As Long, ByVal dx As Long, _
                         ByVal BackColour As Long) As Long

  Dim f As Double, c As UDT_LONG, iRGB As RGBA, pRGB As RGBA, cw As Double

  If (dx Or dy) <> 0 Then   'floating point I'm afraid
    If dx > dy Then
      f = dy / dx
    Else
      f = dx / dy
    End If
    f = 1 / Sqr(1 + f * f)
    cw = 1 - f
    c.v = Colour
    LSet iRGB = c
    c.v = BackColour
    LSet pRGB = c
    AAColour = RGB(f * iRGB.r + cw * pRGB.r, f * iRGB.g + cw * pRGB.g, f * iRGB.b + cw * pRGB.b)
  Else
    AAColour = Colour
  End If

End Function

'===================================== ELLIPSES ============================================================
'check if parameters are large enough/OK  to have to do something

Private Function StartingEllipse(ByVal hDC As Long, _
                                ByVal x0 As Long, ByVal y0 As Long, _
                                ByVal a As Long, ByVal b As Long, _
                                ByVal Colour As Long) As Boolean

  If a <= 0 Or b <= 0 Then
    Call SetPixelV(hDC, x0, y0, Colour)       'A centrepoint only
  Else  'a>0 and b>0
    Call SetPixelV(hDC, x0 + a, y0, Colour) 'X+   'four known complete points
    Call SetPixelV(hDC, x0, y0 - b, Colour) 'Y+
    Call SetPixelV(hDC, x0 - a, y0, Colour) 'X-
    Call SetPixelV(hDC, x0, y0 + b, Colour) 'Y-
    StartingEllipse = (a > 1 Or b > 1)        'There is More to Draw
  End If

End Function

#If DDEBUG = 0 Then
'Elliptical 4 symmetry

Private Sub SetPixel4(ByVal hDC As Long, _
                     ByVal x0 As Long, ByVal y0 As Long, _
                     ByVal dx As Long, ByVal dy As Long, _
                     ByVal Colour As Long)

  Call SetPixelV(hDC, x0 + dx, y0 - dy, Colour)   'Q0
  Call SetPixelV(hDC, x0 - dx, y0 - dy, Colour)   'Q1
  Call SetPixelV(hDC, x0 - dx, y0 + dy, Colour)   'Q2
  Call SetPixelV(hDC, x0 + dx, y0 + dy, Colour)   'Q3

End Sub

'4 symmetry Ellipse points with consideration of background based on weight

Private Sub BlendPixel4(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal dx As Long, ByVal dy As Long, _
                       ByVal Colour As Long, _
                       ByVal w As Long)

  Call BlendPixel(hDC, x0 + dx, y0 - dy, Colour, w) 'Q0
  Call BlendPixel(hDC, x0 - dx, y0 - dy, Colour, w) 'Q1
  Call BlendPixel(hDC, x0 - dx, y0 + dy, Colour, w) 'Q2
  Call BlendPixel(hDC, x0 + dx, y0 + dy, Colour, w) 'Q3

End Sub
#End If

'Anti Aliasing Routines for Ellipses

Private Sub BlendPixel4Y(ByVal hDC As Long, _
                        ByVal x0 As Long, ByVal y0 As Long, _
                        ByVal dx As Long, ByVal dy As Long, _
                        ByVal Colour As Long, _
                        ByVal w As Long)

  If w < 0 Then     'W is the weight 0,,MAX_WEIGHT of Pixel(x,y+1) and 1-wDy = weight Pixel(x,y)
    dy = dy + 1     'if W<0 then we need to complement and setp back one
    w = MAX_WEIGHT + w
  End If

  Call BlendPixel4(hDC, x0, y0, dx, dy, Colour, MAX_WEIGHT - w) 'Weight of Pixel(x,y+1)
  Call BlendPixel4(hDC, x0, y0, dx, dy - 1, Colour, w)          'Weight of Pixel(x,y  )

End Sub

Private Sub BlendPixel4X(ByVal hDC As Long, _
                        ByVal x0 As Long, ByVal y0 As Long, _
                        ByVal dx As Long, ByVal dy As Long, _
                        ByVal Colour As Long, _
                        ByVal w As Long)

  If w < 0 Then     'W is the weight 0,,MAX_WEIGHT of Pixel(x-1,y) and 1-w = weight Pixel(x,y)
    dx = dx + 1     'if W<0 then we need to complement and setp back one
    w = MAX_WEIGHT + w
  End If

  Call BlendPixel4(hDC, x0, y0, dx, dy, Colour, MAX_WEIGHT - w) 'Weight of Pixel(x  ,y)
  Call BlendPixel4(hDC, x0, y0, dx - 1, dy, Colour, w)          'Weight of Pixel(x-1,y)

End Sub

#If DDEBUG = 0 Then

'======================================= CIRCLES ==========================================================
'check if parameters are large enough/OK  to have to do something

Private Function StartingCircle(ByVal hDC As Long, _
                               ByVal x0 As Long, ByVal y0 As Long, ByVal r As Long, _
                               ByVal Colour As Long) As Boolean

  If r <= 0 Then
    Call SetPixelV(hDC, x0, y0, Colour)
  Else
    Call SetPixelV(hDC, x0 + r, y0, Colour) 'X+
    Call SetPixelV(hDC, x0, y0 - r, Colour) 'Y+ 'known solid points
    Call SetPixelV(hDC, x0 - r, y0, Colour) 'X-
    Call SetPixelV(hDC, x0, y0 + r, Colour) 'Y-
    StartingCircle = (r > 1)                    'Circle continues after these points
  End If

End Function

'Plot ALL 8 symmetry points by Octant as Viewed on screen

Private Sub SetPixel8(ByVal hDC As Long, _
                     ByVal x0 As Long, ByVal y0 As Long, _
                     ByVal dx As Long, ByVal dy As Long, _
                     ByVal Colour As Long)

  Call SetPixelV(hDC, x0 + dy, y0 - dx, Colour) 'O0
  Call SetPixelV(hDC, x0 + dx, y0 - dy, Colour) 'O1
  Call SetPixelV(hDC, x0 - dx, y0 - dy, Colour) 'O2
  Call SetPixelV(hDC, x0 - dy, y0 - dx, Colour) 'O3
  Call SetPixelV(hDC, x0 - dy, y0 + dx, Colour) 'O4
  Call SetPixelV(hDC, x0 - dx, y0 + dy, Colour) 'O5
  Call SetPixelV(hDC, x0 + dx, y0 + dy, Colour) 'O6
  Call SetPixelV(hDC, x0 + dy, y0 + dx, Colour) 'O7

End Sub

'8 symmetry Circle points with consideration of background based on weight by Octant as Viewed on screen

Private Sub BlendPixel8(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal dx As Long, ByVal dy As Long, _
                       ByVal Colour As Long, _
                       ByVal w As Long)

  'If dx > dy Then Exit Sub
  Call BlendPixel(hDC, x0 + dy, y0 - dx, Colour, w) 'O0
  Call BlendPixel(hDC, x0 + dx, y0 - dy, Colour, w) 'O1
  Call BlendPixel(hDC, x0 - dx, y0 - dy, Colour, w) 'O2
  Call BlendPixel(hDC, x0 - dy, y0 - dx, Colour, w) 'O3
  Call BlendPixel(hDC, x0 - dy, y0 + dx, Colour, w) 'O4
  Call BlendPixel(hDC, x0 - dx, y0 + dy, Colour, w) 'O5
  Call BlendPixel(hDC, x0 + dx, y0 + dy, Colour, w) 'O6
  Call BlendPixel(hDC, x0 + dy, y0 + dx, Colour, w) 'O7

End Sub
#End If

'Plot ALL 8 symmetry points by Octant as Viewed on screen (2colours)

Private Function StartingCircle3D(ByVal hDC As Long, _
                               ByVal x0 As Long, ByVal y0 As Long, ByVal r As Long, _
                               ByVal HiColour As Long, ByVal LoColour As Long) As Boolean

  If r <= 0 Then
    Call SetPixelV(hDC, x0, y0, HiColour)
  Else
    Call SetPixelV(hDC, x0 + r, y0, LoColour) 'X+
    Call SetPixelV(hDC, x0, y0 - r, HiColour) 'Y+ 'known solid points
    Call SetPixelV(hDC, x0 - r, y0, HiColour) 'X-
    Call SetPixelV(hDC, x0, y0 + r, LoColour) 'Y-
    StartingCircle3D = (r > 1)                    'Circle continues after these points
  End If

End Function

Private Sub SetPixel83D(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal dx As Long, ByVal dy As Long, _
                       ByVal HiColour As Long, ByVal LoColour As Long)

  Dim f As Boolean
  
  f = (989 * dx > 571 * dy)   'the boundary at 60deg
  
  Call SetPixelV(hDC, x0 + dy, y0 - dx, LoColour)   'O0
  If f Then
    Call SetPixelV(hDC, x0 + dx, y0 - dy, LoColour) 'O1
  Else
    Call SetPixelV(hDC, x0 + dx, y0 - dy, HiColour) 'O1
  End If
  Call SetPixelV(hDC, x0 - dx, y0 - dy, HiColour)   'O2
  Call SetPixelV(hDC, x0 - dy, y0 - dx, HiColour)   'O3
  Call SetPixelV(hDC, x0 - dy, y0 + dx, HiColour)   'O4
  If f Then
    Call SetPixelV(hDC, x0 - dx, y0 + dy, HiColour) 'O5
  Else
    Call SetPixelV(hDC, x0 - dx, y0 + dy, LoColour) 'O5
  End If
  Call SetPixelV(hDC, x0 + dx, y0 + dy, LoColour)   'O6
  Call SetPixelV(hDC, x0 + dy, y0 + dx, LoColour)   'O7

End Sub

'8 symmetry Circle points with consideration of background based on weight by Octant as Viewed on screen

Private Sub BlendPixel83D(ByVal hDC As Long, _
                         ByVal x0 As Long, ByVal y0 As Long, _
                         ByVal dx As Long, ByVal dy As Long, _
                         ByVal HiColour As Long, ByVal LoColour As Long, _
                         ByVal w As Long)

  Dim f As Boolean
  
  f = (989 * dx > 571 * dy)   'the boundary at 60deg
  
  Call BlendPixel(hDC, x0 + dy, y0 - dx, LoColour, w) 'O0
  If f Then
    Call BlendPixel(hDC, x0 + dx, y0 - dy, LoColour, w) 'O1
  Else
    Call BlendPixel(hDC, x0 + dx, y0 - dy, HiColour, w) 'O1
  End If
  Call BlendPixel(hDC, x0 - dx, y0 - dy, HiColour, w) 'O2
  Call BlendPixel(hDC, x0 - dy, y0 - dx, HiColour, w) 'O3
  Call BlendPixel(hDC, x0 - dy, y0 + dx, HiColour, w) 'O4
  If f Then
    Call BlendPixel(hDC, x0 - dx, y0 + dy, HiColour, w) 'O5
  Else
    Call BlendPixel(hDC, x0 - dx, y0 + dy, LoColour, w) 'O5
  End If
  Call BlendPixel(hDC, x0 + dx, y0 + dy, LoColour, w) 'O6
  Call BlendPixel(hDC, x0 + dy, y0 + dx, LoColour, w) 'O7

End Sub


#If DDEBUG = 1 Then
'USE THESE TO TEST THAT ELLIPSES and CIRCLES ARE DRAWN CORRECTLY

'4 symmetry Ellipse points with consideration of background based on weight

Private Sub SetPixel4(ByVal hDC As Long, _
                     ByVal x0 As Long, ByVal y0 As Long, _
                     ByVal dx As Long, ByVal dy As Long, _
                     ByVal Colour As Long)

  Call SetPixelV(hDC, x0 + dx, y0 - dy, vbBlack)  'Q0
  Call SetPixelV(hDC, x0 - dx, y0 - dy, vbRed)    'Q1
  Call SetPixelV(hDC, x0 - dx, y0 + dy, vbGreen)  'Q2
  Call SetPixelV(hDC, x0 + dx, y0 + dy, vbBlue)   'Q3

End Sub

Private Sub BlendPixel4(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal dx As Long, ByVal dy As Long, _
                       ByVal Colour As Long, _
                       ByVal w As Long)

  Call BlendPixel(hDC, x0 + dx, y0 - dy, vbBlack, w) 'Q0
  Call BlendPixel(hDC, x0 - dx, y0 - dy, vbRed, w) 'Q1
  Call BlendPixel(hDC, x0 - dx, y0 + dy, vbGreen, w) 'Q2
  Call BlendPixel(hDC, x0 + dx, y0 + dy, vbBlue, w) 'Q3

End Sub

Private Function StartingCircle(ByVal hDC As Long, _
                               ByVal x0 As Long, ByVal y0 As Long, ByVal r As Long, _
                               ByVal Colour As Long) As Boolean

  If r <= 0 Then
    Call SetPixelV(hDC, x0, y0, Colour)
  Else
    Call SetPixelV(hDC, x0 + r, y0, vbBlack) 'X+
    Call SetPixelV(hDC, x0, y0 - r, vbGreen) 'Y+ 'known solid points
    Call SetPixelV(hDC, x0 - r, y0, vbBlue) 'X-
    Call SetPixelV(hDC, x0, y0 + r, vbCyan) 'Y-
    StartingCircle = (r > 1)                    'Circle continues after these points
  End If

End Function

'Plot ALL 8 symmetry points by Octant as Viewed on screen

Private Sub SetPixel8(ByVal hDC As Long, _
                     ByVal x0 As Long, ByVal y0 As Long, _
                     ByVal dx As Long, ByVal dy As Long, _
                     ByVal Colour As Long)

  Call SetPixelV(hDC, x0 + dy, y0 - dx, vbBlack)    'O0
  Call SetPixelV(hDC, x0 + dx, y0 - dy, vbRed)      'O1
  Call SetPixelV(hDC, x0 - dx, y0 - dy, vbGreen)    'O2
  Call SetPixelV(hDC, x0 - dy, y0 - dx, vbYellow)   'O3
  Call SetPixelV(hDC, x0 - dy, y0 + dx, vbBlue)     'O4
  Call SetPixelV(hDC, x0 - dx, y0 + dy, vbMagenta)  'O5
  Call SetPixelV(hDC, x0 + dx, y0 + dy, vbCyan)     'O6
  Call SetPixelV(hDC, x0 + dy, y0 + dx, vbWhite)    'O7

End Sub

Private Sub BlendPixel8(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal dx As Long, ByVal dy As Long, _
                       ByVal Colour As Long, _
                       ByVal w As Long)

  If dx > dy Then Exit Sub
  Call BlendPixel(hDC, x0 + dy, y0 - dx, vbBlack, w) 'O0
  Call BlendPixel(hDC, x0 + dx, y0 - dy, vbRed, w) 'O1
  Call BlendPixel(hDC, x0 - dx, y0 - dy, vbGreen, w) 'O2
  Call BlendPixel(hDC, x0 - dy, y0 - dx, vbYellow, w) 'O3
  Call BlendPixel(hDC, x0 - dy, y0 + dx, vbBlue, w) 'O4
  Call BlendPixel(hDC, x0 - dx, y0 + dy, vbMagenta, w) 'O5
  Call BlendPixel(hDC, x0 + dx, y0 + dy, vbCyan, w) 'O6
  Call BlendPixel(hDC, x0 + dy, y0 + dx, vbWhite, w) 'O7

End Sub
#End If

'===========================================================================================================
'An Anti-Aliased (x^2 style) version of optimised Bresenhams Circle good for r=1 to 16384 at least
'===========================================================================================================

Public Sub DrawCircleAA(ByVal hDC As Long, _
                        ByVal x0 As Long, ByVal y0 As Long, ByVal r As Long, _
                        ByVal Colour As Long)

  Dim dx As Long, dy As Long, dd As Long
  Dim incE As Long, incSE As Long, incEx As Long, incEy As Long
  Dim Er As Long, Ez As Long, w As Long, dytmp As Long

  If StartingCircle(hDC, x0, y0, r, Colour) Then
    dx = 0
    dy = r
    dd = 1 - r
    incE = 3
    incSE = 5 - r * 2
    incEx = incE - 2    '1
    incEy = incSE - 4
    Ez = 4 - incSE      'r^2-1

    Do While dx < dy - 1
      If dd < 0 Then        '(x,y)->(x+1,Y)
        dd = dd + incE
        incSE = incSE + 2   '
      Else                  '(x,y)->(x+1,y-1)
        dd = dd + incSE
        incSE = incSE + 4
        dy = dy - 1
        'error
        Er = Er + incEy
        incEy = incEy + 2
        Ez = Ez - 2
      End If
      dx = dx + 1
      incE = incE + 2
      'error
      Er = Er + incEx
      incEx = incEx + 2

      w = (MAX_WEIGHT * Er) \ Ez    'W is the weight 0,,MAX_WEIGHT of Pixel(x,y+1) and 1-w = weight Pixel(x,y)
      dytmp = dy
      If w < 0 Then                 'if W<0 then we need to complement and step back one
        dytmp = dytmp + 1
        w = MAX_WEIGHT + w
      End If
      Call BlendPixel8(hDC, x0, y0, dx, dytmp, Colour, MAX_WEIGHT - w) 'Weight of Pixel(x,y)
      Call BlendPixel8(hDC, x0, y0, dx, dytmp - 1, Colour, w)       'Weight of Pixel(x,y+1)
    Loop
  End If
  
End Sub

