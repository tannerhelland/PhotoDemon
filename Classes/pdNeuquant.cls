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
'Last updated: 16/September/21
'Last update: initial build
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
' resample the source image; I have much faster functions for that than I can implement here.
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

'The original neuquant implementation used a (crappy) pseudo-number generator for sampling.
' (Sampling is controlled by a quality parameter that samples more sparsely as quality drops.)
' I've implemented this as-is for now, but will likely revisit with a true random-number
' generator in the coming days.
'/* four primes near 500 - assume no image has a length so large */
'/* that it is divisible by all four primes */
Private Const prime1 As Long = 499
Private Const prime2 As Long = 491
Private Const prime3 As Long = 487
Private Const prime4 As Long = 503
Private Const minpicturebytes As Long = 4 * 503     '/* minimum size for input image, channelsize * prime4 */

'/* Network Definitions
Private maxnetpos As Long                   '(netsize-1)
Private Const netbiasshift As Long = 4      '/* bias for colour values */
Private Const ncycles As Long = 100         '/* no. of learning cycles */

'/* defs for freq and bias */
Private Const intbiasshift As Long = 16     '/* bias for fractions */
Private Const intbias As Long = 2 ^ 16      '(((int) 1)<<intbiasshift)
Private Const gammashift As Long = 10       '/* gamma = 1024 */
Private Const gamma As Long = 2 ^ 10        '(((int) 1)<<gammashift)
Private Const betashift As Long = 10
Private Const beta As Long = 64             '(intbias>>betashift)    /* beta = 1/1024 */
Private Const betagamma As Long = 2 ^ 16    '(intbias<<(gammashift-betashift))

'/* defs for decreasing radius factor */
Private initrad As Long                     '(netsize>>3)        /* for 256 cols, radius starts */
Private Const radiusbiasshift As Long = 6   '/* at 32.0 biased by 6 bits */
Private Const radiusbias As Long = 2 ^ 6    '(((int) 1)<<radiusbiasshift)
Private initradius As Long                  '(initrad*radiusbias)    /* and decreases by a */
Private Const radiusdec As Long = 30        '/* factor of 1/30 each cycle */

'/* defs for decreasing alpha factor */
Private Const alphabiasshift As Long = 10   '/* alpha starts at 1.0 */
Private Const initalpha As Long = 2 ^ 10    '(((int) 1)<<alphabiasshift)
Private alphadec As Long                    '/* biased by 10 bits */

'/* radbias and alpharadbias used for radpower calculation */
Private Const radbiasshift As Long = 8
Private Const radbias As Long = 2 ^ 8                   '(((int) 1)<<radbiasshift)
Private Const alpharadbshift As Long = 10 + 8           '(alphabiasshift+radbiasshift)
Private Const alpharadbias As Long = 2 ^ alpharadbshift '(((int) 1)<<alpharadbshift)

'/* Types and Global Variables
'static unsigned char *thepicture;       '/* the input image itself */
Private lengthcount As Long '/* lengthcount = H*W*numChannels */

Private samplefac As Long   '/* sampling factor 1..30 */
Private Type nnPixel          '/* BGRc */
    b As Long
    g As Long
    r As Long
    a As Long
    c As Long
End Type
Private network() As nnPixel    '/* the network itself */, size netsize
Private netindex(0 To 255) As Long  '/* for network lookup - really 256 */
Private bias() As Long          '/* bias and freq arrays for learning */
Private freq() As Long
Private radpower() As Long      '/* radpower for precomputation */

'Call this function first; the number of colors determines the size of the neural network,
' which in turn controls a huge list of run-time "constants".
'
'Max color count is currently limited to 256.  Larger sizes would work fine but be slower,
' and PD has no use for them.
Friend Sub SetColorCount(ByVal numColors As Long)
    
    If (numColors < 0) Then numColors = 256
    If (numColors > 256) Then numColors = 256
    netsize = 256
    
    'Initialize a bunch of other settings contingent on the netsize
    maxnetpos = netsize - 1
    ReDim network(0 To maxnetpos) As nnPixel
    ReDim bias(0 To maxnetpos) As Long
    ReDim freq(0 To maxnetpos) As Long
    
    initrad = netsize \ 8
    initradius = initrad * radiusbias
    ReDim radpower(0 To initrad - 1) As Long
    
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
            .b = (i * (2 ^ (netbiasshift + 8))) \ netsize
            .g = .b
            .r = .b
            .a = .b
        End With
        freq(i) = intbias \ netsize
        bias(i) = 0
    Next i

End Sub

'/* Unbias network to give byte values 0..255 and record position i to prepare for sort
Friend Sub unbiasnet()
    
    Dim rshift As Long
    rshift = 2 ^ netbiasshift
    
    Dim i As Long
    For i = 0 To netsize - 1
        Debug.Print network(i).b, network(i).g, network(i).r, rshift
        network(i).b = network(i).b \ rshift
        network(i).g = network(i).g \ rshift
        network(i).r = network(i).r \ rshift
        network(i).a = network(i).a \ rshift
        network(i).c = i    '/* record colour no */
    Next i
    
End Sub

'/* Output colour map
Friend Sub writecolourmap(ByRef dstPalette() As RGBQuad)
    
    ReDim dstPalette(0 To netsize - 1) As RGBQuad
    
    Dim i As Long
    For i = 0 To netsize - 1
        With dstPalette(i)
            .Blue = network(i).b
            .Green = network(i).g
            .Red = network(i).r
            .Alpha = network(i).a
        End With
    Next i
    
End Sub

'/* Main Learning Loop
Friend Sub learn()
    
    Dim i As Long, j As Long, b As Long, g As Long, r As Long, a As Long
    Dim radius As Long, rad As Long, Alpha As Long, step As Long, delta As Long, samplepixels As Long
    
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
    
    alphadec = 30 + ((samplefac - 1) / 4)   '4?
    
    samplepixels = lengthcount \ (4 * samplefac)
    delta = samplepixels \ ncycles
    Alpha = initalpha
    radius = initradius
    
    rad = radius \ (2 ^ radiusbiasshift)
    If (rad <= 1) Then rad = 0
    For i = 0 To rad - 1
        radpower(i) = Alpha * (((rad * rad - i * i) * radbias) / (rad * rad))
    Next i
    
    PDDebug.LogAction "beginning 1D learning: initial radius=" & rad
    
    If ((lengthcount Mod prime1) <> 0) Then
        step = 4 * prime1
    Else
        If ((lengthcount Mod prime2) <> 0) Then
            step = 4 * prime2
        Else
            If ((lengthcount Mod prime3) <> 0) Then
                step = 4 * prime3
            Else
                step = 4 * prime4
            End If
        End If
    End If
    
    i = 0
    Do While (i < samplepixels)
        b = imgPixels(p) * (2 ^ netbiasshift)
        g = imgPixels(p + 1) * (2 ^ netbiasshift)
        r = imgPixels(p + 2) * (2 ^ netbiasshift)
        a = imgPixels(p + 3) * (2 ^ netbiasshift)
        
        j = contest(b, g, r, a)
        
        altersingle Alpha, j, b, g, r, a
        If (rad <> 0) Then alterneigh rad, j, b, g, r, a '/* alter neighbours */
        
        p = p + step
        If (p >= lengthcount) Then p = p - lengthcount
        
        i = i + 1
        
        If (i Mod delta = 0) Then
            Alpha = Alpha - Alpha \ alphadec
            radius = radius - radius \ radiusdec
            rad = radius \ (2 ^ radiusbiasshift)
            If (rad <= 1) Then rad = 0
            For j = 0 To rad - 1
                radpower(j) = Alpha * (((rad * rad - j * j) * radbias) \ (rad * rad))
            Next j
        End If
        
    Loop
    
    PDDebug.LogAction "finished 1D learning: final alpha=" & CStr(Alpha / initalpha)
    
    'Free unsafe array ref
    m_srcDIB.UnwrapArrayFromDIB imgPixels
    
End Sub

'/* Search for biased BGRA values
Private Function contest(b, g, r, a) As Long

    '/* finds closest neuron (min dist) and updates freq */
    '/* finds best neuron (min dist-bias) and returns position */
    '/* for frequently chosen neurons, freq[i] is high and bias[i] is negative */
    '/* bias[i] = gamma*((1/netsize)-freq[i]) */
    
    Dim i As Long, dist As Long, a2 As Long, biasdist As Long, betafreq As Long
    Dim bestpos As Long, bestbiaspos As Long, bestd As Long, bestbiasd As Long
    
    'register int *p,*f, *n;
    bestd = LONG_MAX    'Not (2 ^ 31)    'bestd = ~(((int) 1)<<31)
    bestbiasd = bestd
    bestpos = -1
    bestbiaspos = bestpos
    'p = bias;
    'f = freq;

    For i = 0 To netsize - 1
        'n = network[i];
        dist = Abs(network(i).b - b)
        a2 = Abs(network(i).g - g)
        dist = dist + a2
        a2 = Abs(network(i).r - r)
        dist = dist + a2
        a2 = Abs(network(i).a - a)
        dist = dist + a2
        If (dist < bestd) Then
            bestd = dist
            bestpos = i
        End If
        biasdist = dist - bias(i)
        If (biasdist < bestbiasd) Then
            bestbiasd = biasdist
            bestbiaspos = i
        End If
        betafreq = freq(i) \ (2 ^ betashift)
        freq(i) = freq(i) - betafreq
        bias(i) = bias(i) + betafreq * (2 ^ gammashift)
        
    Next i
    
    freq(bestpos) = freq(bestpos) + beta
    bias(bestpos) = bias(bestpos) - betagamma
    contest = bestbiaspos

End Function

'/* Move neuron i towards biased (b,g,r) by factor alpha
Private Sub altersingle(Alpha, i, b, g, r, a)
    With network(i)
        .b = .b - (Alpha * (.b - b)) \ initalpha    '/* alter hit neuron */
        .g = .g - (Alpha * (.g - g)) \ initalpha
        .r = .r - (Alpha * (.r - r)) \ initalpha
        .a = .a - (Alpha * (.a - a)) \ initalpha
    End With
End Sub

'/* Move adjacent neurons by precomputed alpha*(1-((i-j)^2/[r]^2)) in radpower[|i-j|]
Private Sub alterneigh(rad, i, b, g, r, a)
    
    Dim j As Long, k As Long, lo As Long, hI As Long, a2 As Long, q As Long
    lo = i - rad
    If (lo < -1) Then lo = -1
    hI = i + rad
    If (hI > netsize) Then hI = netsize

    j = i + 1
    k = i - 1
    q = 0
    
    Do While ((j < hI) Or (k > lo))
        a2 = radpower(q + 1)
        q = q + 1
        
        If (j < hI) Then
            With network(j)
                .b = .b - (a2 * (.b - b)) \ alpharadbias
                .g = .g - (a2 * (.g - g)) \ alpharadbias
                .r = .r - (a2 * (.r - r)) \ alpharadbias
                .a = .a - (a2 * (.a - a)) \ alpharadbias
            End With
            j = j + 1
        End If
        
        If (k > lo) Then
            With network(k)
                .b = .b - (a2 * (.b - b)) \ alpharadbias
                .g = .g - (a2 * (.g - g)) \ alpharadbias
                .r = .r - (a2 * (.r - r)) \ alpharadbias
                .a = .a - (a2 * (.a - a)) \ alpharadbias
            End With
            k = k - 1
        End If
        
    Loop

End Sub
