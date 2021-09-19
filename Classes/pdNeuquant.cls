VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdNeuquant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Neuquant-inspired Neural Network Color Quantization Class
'Copyright 2021-2021 by Tanner Helland
'Created: 16/September/21
'Last updated: 18/September/21
'Last update: convert some math to floating-point for improved performance; get alpha channel working
'
'This class provides a highly optimized (for VB6) Neuquant-inspired neural network
' quantization implementation.  Neuquant was originally published by Anthony Decker,
' and this copyright must be included in any derivative works:
'
'/* NeuQuant Neural-Net Quantization Algorithm
' * ------------------------------------------
' *
' * Copyright (c) 1994 Anthony Dekker
' *
' * NEUQUANT Neural-Net quantization algorithm by Anthony Dekker, 1994.
' * See "Kohonen neural networks for optimal colour quantization"
' * in "Network: Computation in Neural Systems" Vol. 5 (1994) pp 351-367.
' * for a discussion of the algorithm.
' * See also http://www.acm.org/~dekker/NEUQUANT.HTML
' *
' * Any party obtaining a copy of these files from the author, directly or
' * indirectly, is granted, free of charge, a full and unrestricted irrevocable,
' * world-wide, paid up, royalty-free, nonexclusive right and license to deal
' * in this software and documentation files (the "Software"), including without
' * limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' * and/or sell copies of the Software, and to permit persons who receive
' * copies from any such party to do so, with the only requirement being
' * that this copyright notice remain intact.
' */
'
'This class is probably not the best reference implementation of Neuquant, given the weird tricks
' I have to pull to make things like this work in VB6.  A simple online search will turn up many
' Neuquant implementations in other languages that will better serve most developers.  (I first
' learned about this algorithm from the FreeImage project, for example, which provides their
' own C version of the algorithm.)
'
'Note also that this implementation is specifically adapted to PhotoDemon's needs (and toolkit).
' For example, instead of implementing variable sampling rates here, I instead prefer to simply
' resample the source image.  (It's much faster.)  I've also added support for alpha channels,
' converted portions of the algorithm to floating-point (for improved accuracy and performance)
' and changed the prime numbers sampling strategy.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PD-specific reference to source image
Private m_srcDIB As pdDIB

'Because this implementation allows you to specify a target color count, we do not fix various
' network values against 256.  Instead, these "constants" are calculated at run-time based on
' the color count you request.
Private netsize As Long

'The original neuquant implementation used a very primitive sampling mechanism. (Sampling is
' controlled by a quality parameter that samples more sparsely as quality drops.)  I've tweaked
' the formula slightly to work better with modern image sizes, but any number of random sampling
' strategies could be used instead.  The original four primes were 499, 491, 487, and 503.
' (With the default "100 learning cycles" strategy this means the image is sampled 5x on each pass.)
' Note that for "ideal" sampling, you want no image to have a length so large that it's divisible
' by all four primes; for the four used here, the LCM is 3,368,562,317 - so we're safe.
' 241, 239, 233, 251 (faster due to locality - by 15-20%, but results are... harder to predict?  I may
' look at just doing something like a fisher-yates shuffle on the source image, then skipping this
' sampling nonsense here)
Private Const prime1 As Long = 241
Private Const prime2 As Long = 239
Private Const prime3 As Long = 233
Private Const prime4 As Long = 251
Private Const minpicturebytes As Long = 4 * 251     '/* minimum size for input image, channelsize * prime4 */

'/* Network Definitions
Private maxnetpos As Long                   '(netsize-1)
Private Const ncycles As Long = 100         '/* no. of learning cycles */

Private Const gamma As Double = 1024#
Private Const beta As Double = 1# / 1024#
Private Const betagamma As Double = beta * gamma

'/* defs for decreasing radius factor */
Private initrad As Long                     '(netsize>>3)        /* for 256 cols, radius starts at 32 */
Private Const radiusbias As Long = 2 ^ 6    '(((int) 1)<<radiusbiasshift)
Private initradius As Long                  '(initrad*radiusbias)    /* and decreases by a */
Private Const radiusdec As Long = 30        '/* factor of 1/30 each cycle */

'/* defs for decreasing alpha factor */
Private Const initalpha As Long = 2 ^ 10    '(((int) 1)<<alphabiasshift)
Private alphadec As Long                    '/* biased by 10 bits */

'/* Types and Global Variables
Private lengthcount As Long '/* lengthcount = H*W*numChannels */

Private samplefac As Long   '/* sampling factor 1..30 */
Private Type nnPixel          '/* BGRA but floats; note that any three-component+alpha color space works, e.g. LABa */
    b As Single
    g As Single
    r As Single
    a As Single
End Type
Private network() As nnPixel    '/* the network itself */, size netsize
Private bias() As Single          '/* bias and freq arrays for learning */
Private freq() As Single

'Call this function first; the number of colors determines the size of the neural network,
' which in turn controls a huge list of run-time "constants".
'
'Max color count is currently limited to 256.  Larger sizes would work fine but be slower,
' and PD has no use for them.
Friend Sub SetColorCount(ByVal numColors As Long)
    
    If (numColors < 0) Then numColors = 256
    If (numColors > 256) Then numColors = 256
    
    'For testing, it can be helpful to force 256 colors (since many parts of the algorithm rely on this value)
    'netsize = 256
    netsize = numColors
    
    'Initialize a bunch of other settings contingent on the netsize
    maxnetpos = netsize - 1
    ReDim network(0 To maxnetpos) As nnPixel
    ReDim bias(0 To maxnetpos) As Single
    ReDim freq(0 To maxnetpos) As Single
    
    initrad = netsize \ 8
    initradius = initrad * radiusbias
    
End Sub

'/* Initialise network in range (0,0,0) to (255,255,255) and set parameters
' (sample is "quality" on the range [0, 30]
Friend Sub initnet(ByRef srcDIB As pdDIB, ByVal sample As Long)
    
    Set m_srcDIB = srcDIB
    lengthcount = m_srcDIB.GetDIBStride * m_srcDIB.GetDIBHeight
    samplefac = sample
    
    Dim i As Long
    For i = 0 To netsize - 1
        With network(i)
            .b = (255 * i) / netsize
            .g = .b
            .r = .b
            .a = .b
        End With
        freq(i) = 1! / netsize
        bias(i) = 0!
    Next i

End Sub

'/* Output colour map
Friend Sub writecolourmap(ByRef dstPalette() As RGBQuad)
    
    ReDim dstPalette(0 To netsize - 1) As RGBQuad
    Dim b As Long, g As Long, r As Long, a As Long
    
    Dim i As Long
    For i = 0 To netsize - 1
        With dstPalette(i)
            b = Int(network(i).b + 0.5)
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            g = Int(network(i).g + 0.5)
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            r = Int(network(i).r + 0.5)
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            a = Int(network(i).a + 0.5)
            If (a > 255) Then a = 255
            If (a < 0) Then a = 0
            .Blue = b
            .Green = g
            .Red = r
            .Alpha = a
        End With
    Next i
    
End Sub

'/* Main Learning Loop
Friend Sub learn()
    
    Dim i As Long, j As Long, b As Single, g As Single, r As Single, a As Single, a2 As Single
    Dim radius As Long, rad As Long, Alpha As Single, step As Long, delta As Long, samplepixels As Long
    
    'Wrap a 1D array around the source pixels
    Dim imgPixels() As Byte, srcSA As SafeArray1D
    m_srcDIB.WrapArrayAroundDIB_1D imgPixels, srcSA
    
    'Instead of p as a pointer, we use it as an index into imgPixels
    'register unsigned char *p;
    'p = thepicture;
    Dim p As Long
    
    'lim becomes unnecessary, since we can use lengthcount automatically
    'unsigned char *lim;
    'lim = thepicture + lengthcount;
    
    Dim biasradius As Long
    biasradius = initrad * radiusbias
    
    alphadec = 30 + ((samplefac - 1) / 3)
    
    samplepixels = lengthcount \ (4 * samplefac)
    delta = samplepixels \ ncycles
    If (delta < 1) Then delta = 1
    Alpha = initalpha
    radius = initradius
    
    rad = radius \ 64   '(2 ^ radiusbiasshift)
    If (rad <= 1) Then rad = 0
    
    Const DIV_INIT_ALPHA As Single = 1! / initalpha
    
    PDDebug.LogAction "beginning 1D learning: initial radius=" & rad
    
    If ((lengthcount Mod prime1) <> 0) Then
        step = prime1 * 4
    Else
        If ((lengthcount Mod prime2) <> 0) Then
            step = prime2 * 4
        Else
            If ((lengthcount Mod prime3) <> 0) Then
                step = prime3 * 4
            Else
                step = prime4 * 4
            End If
        End If
    End If
    
    Dim numPixelsPreviousLoops As Long
    
    i = 0
    Do While (i + numPixelsPreviousLoops < samplepixels)
        b = imgPixels(p)
        g = imgPixels(p + 1)
        r = imgPixels(p + 2)
        a = imgPixels(p + 3)
        
        j = contest(b, g, r, a)
        
        a2 = Alpha * DIV_INIT_ALPHA '/ initalpha
        
        'altersingle a2, j, b, g, r, a '/* alter best-fit color */
        'in-line for performance:
        With network(j)
            .b = .b - (a2 * (.b - b)) '/* alter hit neuron */
            .g = .g - (a2 * (.g - g))
            .r = .r - (a2 * (.r - r))
            .a = .a - (a2 * (.a - a))
        End With
        
        If (rad <> 0) Then alterneigh a2, rad, j, b, g, r, a '/* alter neighbours */
        
        p = p + step
        If (p >= lengthcount) Then p = p - lengthcount
        
        i = i + 1
        
        'Mod is expensive; instead, just reset on each hit on delta
        If (i = delta) Then
            numPixelsPreviousLoops = numPixelsPreviousLoops + i
            i = 0
            Alpha = Alpha - Alpha / alphadec
            biasradius = biasradius - biasradius / radiusdec
            rad = biasradius \ 64
            If (rad <= 1) Then rad = 0
        End If
        
    Loop
    
    PDDebug.LogAction "finished 1D learning: final alpha=" & CStr(Alpha / initalpha)
    
    'Free unsafe array ref
    m_srcDIB.UnwrapArrayFromDIB imgPixels
    
End Sub

'/* Search for biased BGRA values
Private Function contest(ByVal b As Single, ByVal g As Single, ByVal r As Single, ByVal a As Single) As Long

    '/* finds closest neuron (min dist) and updates freq */
    '/* finds best neuron (min dist-bias) and returns position */
    '/* for frequently chosen neurons, freq[i] is high and bias[i] is negative */
    '/* bias[i] = gamma*((1/netsize)-freq[i]) */
    
    Dim dist As Single, a2 As Single, biasdist As Single
    Dim bestpos As Long, bestbiaspos As Long, bestd As Single, bestbiasd As Single
    
    bestd = SINGLE_MAX
    bestbiasd = bestd
    bestpos = -1
    bestbiaspos = bestpos
    
    Dim i As Long
    For i = 0 To netsize - 1
        
        'The original paper uses Manhattan distance instead of Euclidean, and I'm not sure why.
        ' Abs() is an expensive operation - more expensive than multiplication - but things were
        ' different in the early 90's, so maybe it made sense then.  (The paper vaguely alludes
        ' to this: https://web.archive.org/web/20030503154334/http://members.ozemail.com.au/~dekker/NeuQuant.pdf)
        ' Anyway, I've switched to Euclidean here because it's faster and the resulting palette is better.
        dist = (network(i).b - b)
        a2 = network(i).g - g
        dist = dist * dist + a2 * a2
        a2 = network(i).r - r
        dist = dist + a2 * a2
        a2 = network(i).a - a
        dist = dist + a2 * a2
        If (dist < bestd) Then
            bestd = dist
            bestpos = i
        End If
        biasdist = dist - bias(i)
        If (biasdist < bestbiasd) Then
            bestbiasd = biasdist
            bestbiaspos = i
        End If
        freq(i) = freq(i) - beta * freq(i)
        bias(i) = bias(i) + betagamma * freq(i)
    Next i
    
    freq(bestpos) = freq(bestpos) + beta
    bias(bestpos) = bias(bestpos) - betagamma
    contest = bestbiaspos

End Function

'/* Move neuron i towards biased (b, g, r) by factor alpha
Private Sub altersingle(ByVal Alpha As Single, ByVal i As Long, ByVal b As Single, ByVal g As Single, ByVal r As Single, ByVal a As Single)
    With network(i)
        .b = .b - (Alpha * (.b - b)) '/* alter hit neuron */
        .g = .g - (Alpha * (.g - g))
        .r = .r - (Alpha * (.r - r))
        .a = .a - (Alpha * (.a - a))
    End With
End Sub

'/* Move adjacent neurons
Private Sub alterneigh(ByVal Alpha As Single, ByVal rad As Long, ByVal i As Long, ByVal b As Single, ByVal g As Single, ByVal r As Single, ByVal a As Single)
    
    Dim j As Long, k As Long, lo As Long, hI As Long, a2 As Double, q As Long
    lo = i - rad
    If (lo < -1) Then lo = -1
    hI = i + rad
    If (hI > netsize) Then hI = netsize

    j = i + 1
    k = i - 1
    q = 0
    
    Dim raddiv As Double
    raddiv = 1# / (rad * rad)
    
    Do While ((j < hI) Or (k > lo))
        a2 = (Alpha * (rad * rad - q * q)) * raddiv
        q = q + 1
        
        If (j < hI) Then
            With network(j)
                .b = .b - (a2 * (.b - b))
                .g = .g - (a2 * (.g - g))
                .r = .r - (a2 * (.r - r))
                .a = .a - (a2 * (.a - a))
            End With
            j = j + 1
        End If
        
        If (k > lo) Then
            With network(k)
                .b = .b - (a2 * (.b - b))
                .g = .g - (a2 * (.g - g))
                .r = .r - (a2 * (.r - r))
                .a = .a - (a2 * (.a - a))
            End With
            k = k - 1
        End If
        
    Loop

End Sub
