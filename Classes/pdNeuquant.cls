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
'Copyright 2021-2026 by Tanner Helland
'Created: 16/September/21
'Last updated: 21/September/21
'Last update: myriad optimizations against VB6 quirks
'
'This class provides a highly optimized (for VB6) Neuquant-inspired neural network
' color quantization implementation.  Neuquant was originally published by Anthony Decker,
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
' converted all the critical portions of the algorithm to floating-point (for improved accuracy
' and performance), refactored all the code to minimize the HUGE set of magic numbers in the
' original, and changed the prime numbers sampling strategy to make it work better on small images.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PD-specific reference to source image
Private m_srcDIB As pdDIB

'One "neuron" exists for each color in the final palette.  (So to generate a 256-color palette,
' we require 256 neurons.)
Private m_neuronCount As Long

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

'Number of learning cycles.  On each cycle, the current network training radius (e.g. how many neighboring
' neurons are affected by each added color) decreases until we're simply modifying single neurons on each
' encountered color.  This produces an increasingly refined network describing the analyzed color set.
Private Const m_numTrainingCycles As Long = 100

'Sampling factor controls how much of the image is actually sampled when constructing the palette.
' [1, 30] where 1 = sample every pixel, 30 = sample 1/30th of pixels.  Values above 30 are probably
' OK but have not been tested.
Private m_samplingFactor As Long

'NOTE: in previous builds, each neuron was stored as the custom type described below.  While this
' produced very pretty, readable code, it's slower to access UDT members in VB6 (with or without a
' With statement) than it is to access a bare primitive array.  As such, the type below has been
' replaced by a bare float array, and we simply access members manually (e.g. BGRA as idx * 6 + 0
' for blue, idx * 6 + 1 for green, etc).

''Original type definition follows:
''Each neuron stores a BGRA value in floating-point; note that the actual components used do *not* matter,
'' so you can freely plug-in e.g. LABa values instead and the algorithm works the same way.  The only
'' component whose behavior must be consistent is a/alpha, because it is initialized to all 255 to better
'' match the average input to this class.
'Private Type PD_NNPixel
'    b As Single
'    g As Single
'    r As Single
'    a As Single
'
'    'Frequency and bias arrays are used by the core training function.  See FindBestNeuron() for details.
'    ' (Originally these were declared as two separate arrays, but we get way better locality by declaring
'    ' them alongside color components because the same indices in either array are always accessed together.)
'    frequency As Single
'    bias As Single
'End Type

'The neuron network itself, one neuron for every color in the output palette
Private m_network() As Single

'Call this function first; the number of colors determines the size of the neural network,
' which in turn controls a huge list of run-time "constants".
'
'Max color count is currently limited to 256.  Larger sizes would work fine but be slower,
' and PD has no use for them.
Friend Sub SetColorCount(ByVal numColors As Long)
    
    'Ensure color count is in the range [2, 256]
    If (numColors < 2) Then numColors = 2
    If (numColors > 256) Then numColors = 256
    m_neuronCount = numColors
    
    'Initialize a bunch of other settings contingent on the m_neuronCount
    
    'Create the base underlying array, with 6 floats per neuron (BGRA, frequency, bias)
    ReDim m_network(0 To m_neuronCount * 6 - 1) As Single
    
End Sub

'Initialise the neural network along the gray axis (diagonal); this is the initial set of neurons that will
' be "pushed" and "pulled" as colors are added to the network.
'INPUTS:
' - source DIB, must be 32-bpp
' - samplingQuality, which represents "quality" on the range [1, 30]; this affects sampling density of the
'   underlying image, with 1 meaning "sample every pixel" and 30 meaning "sample 1/30th of pixels"
Friend Function InitializeNeuralNetwork(ByRef srcDIB As pdDIB, Optional ByVal samplingQuality As Long = 1) As Long
    
    'Cache network parameters, including a reference to the source image
    Set m_srcDIB = srcDIB
    m_samplingFactor = samplingQuality
    
    'Initialize the network with [numPaletteColors] neurons evenly spaced along the gray axis
    Dim i As Long, idx As Long
    For i = 0 To m_neuronCount - 1
        
        '6 floats are used for each network (BGRA, frequency, and bias).  We store them inside
        ' a naked float array of size colorCount * 6 because using a custom UDT is slower
        ' in VB6 than accessing a primitive (even though we have to manually calculate offsets).
        idx = i * 6
        
        'BGRA
        m_network(idx) = (255! * i) / m_neuronCount
        m_network(idx + 1) = m_network(idx)
        m_network(idx + 2) = m_network(idx)
        m_network(idx + 3) = 255!     'Note that we initialize the network with assumed full opacity, by design
        
        'Frequency and bias
        m_network(idx + 4) = 1! / m_neuronCount
        m_network(idx + 5) = 0!
        
    Next i
    
    'Return the number of pixels we intend to sample; the caller can use this for tracking progress
    InitializeNeuralNetwork = (m_srcDIB.GetDIBWidth * m_srcDIB.GetDIBHeight) \ m_samplingFactor

End Function

'Main learning loop; basically, iterate sampled pixels and alter the network accordingly, starting with
' a large bias radius and shrinking it as we process more and more pixels
Friend Sub TrainNeuralNetwork(Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0)
    
    Dim b As Single, g As Single, r As Single, a As Single
    
    'Wrap a 1D array around the source pixels
    Dim imgPixels() As Byte, srcSA As SafeArray1D
    m_srcDIB.WrapArrayAroundDIB_1D imgPixels, srcSA
    
    'Neurons are altered in a radius that decreases on each cycle.  For a 256-color palette, an effective
    ' initial radius is 32 (this value comes from the original Neuquant paper) and we extend this ratio to
    ' any size of palette.
    Dim radius As Long
    radius = m_neuronCount \ 8
    
    'For very small color counts, there is no point in modifying neighboring neurons - instead, we can
    ' operate on single neurons for much better performance.  (Note that this same check is applied
    ' on the inner loop, which decreases the radius on each cycle.)
    If (radius <= 1) Then radius = 0
    
    'At the end of each cycle, we'll reduce the current radius by a fixed factor.  (The original paper
    ' suggests 1/30.)  Obviously this requires scaling on each pass or just using a floating-point
    ' tracker.  We use a floating-point tracker as it's far more intuitive, and it avoids the need
    ' for a bunch of magic-number constants.
    Const RADIUS_DECREASE As Single = 29! / 30!
    Dim fRadius As Single
    fRadius = radius
    
    'The total number of pixels that we intend to sample depends on the sampling factor supplied
    ' by the user (on the scale [1, 30], where 1 means "sample every pixel").
    Dim totalPixels As Long
    totalPixels = m_srcDIB.GetDIBWidth * m_srcDIB.GetDIBHeight
    
    'Calculate the net number of pixels we want to sample from the image.  (When sampling factor = 1,
    ' this is "all the pixels in the image", e.g. totalPixels.)
    Dim numPixelsToSample As Long
    numPixelsToSample = totalPixels \ m_samplingFactor
    
    'Calculate the number of pixels we want to sample on each cycle.  When we hit this number,
    ' we move to the next cycle (which triggers changes to the algorithm like reducing radius
    ' of how many neighboring neurons are affected on each weight, and reducing how much we move
    ' each matched neuron).
    Dim numPixelsPerCycle As Long
    numPixelsPerCycle = numPixelsToSample \ m_numTrainingCycles
    If (numPixelsPerCycle < 1) Then numPixelsPerCycle = 1
    
    'Weight (called alpha in the original paper, but changed here to avoid confusion with opacity)
    ' controls how much we "push" the best-match neuron toward the latest color encountered.
    ' Alpha starts at 1 and decreases on each cycle as the model converges.
    Dim weight As Single
    weight = 1!
    
    'This value (effectively a constant, if VB supported run-time consts) is the ratio describing
    ' how much we decrease neuron weighting on each cycle.  This value is applied to weight on each
    ' cycle increment.  (Note that this value is also affected by the sampling factor; as the
    ' sampling factor increases - meaning we sample less of the image - our weighting also decreases.
    Dim weightDecreasePerCycle As Single
    weightDecreasePerCycle = 1! / (30 + ((m_samplingFactor - 1) / 3))
    
    PDDebug.LogAction "Starting neural network construction; pixels to sample: " & numPixelsToSample
    
    Dim step As Long
    If ((totalPixels Mod prime1) <> 0) Then
        step = prime1 * 4
    Else
        If ((totalPixels Mod prime2) <> 0) Then
            step = prime2 * 4
        Else
            If ((totalPixels Mod prime3) <> 0) Then
                step = prime3 * 4
            Else
                step = prime4 * 4
            End If
        End If
    End If
    
    'We track processed pixels and current cycle in order to report progress back to the caller
    Dim currentCycle As Long, numPixelsPreviousLoops As Long
    
    'Index into the network array, and the array wrapped around the source image
    Dim idx As Long, idxPixel As Long
    
    Dim i As Long, j As Long
    i = 0
    Do While (i + numPixelsPreviousLoops < numPixelsToSample)
        
        'Pull original pixel values and store them as floating-point
        b = imgPixels(idxPixel)
        g = imgPixels(idxPixel + 1)
        r = imgPixels(idxPixel + 2)
        a = imgPixels(idxPixel + 3)
        
        'Find the best neuron match for this color.  Note that we *cannot* accelerate this using
        ' something like RLE, because neurons are biased as we select them more frequently (which prevents
        ' distortion for a few single best-match values).  The original paper describes the importance
        ' of biasing hits in more detail.
        j = FindBestNeuron(b, g, r, a)
        
        'Originally this was a function call (AlterSingleNeuron), but I've manually in-lined it
        ' for better performance:
        idx = j * 6
        m_network(idx) = m_network(idx) - weight * (m_network(idx) - b)
        m_network(idx + 1) = m_network(idx + 1) - weight * (m_network(idx + 1) - g)
        m_network(idx + 2) = m_network(idx + 2) - weight * (m_network(idx + 2) - r)
        m_network(idx + 3) = m_network(idx + 3) - weight * (m_network(idx + 3) - a)
        
        'If this loop has a non-zero radius, alter neighboring neurons next
        If (radius > 0) Then AlterNeighbors weight, radius, j, b, g, r, a
        
        'Move the pixel pointer to a new location (remember - step is a prime number).
        ' Obviously, we need to wrap if we move beyond the end of the image.
        idxPixel = idxPixel + step
        If (idxPixel >= totalPixels * 4) Then idxPixel = idxPixel - (totalPixels * 4)
        
        'Increment the pixel processing counter.  After [numPixelsPerCycle] iterations, we need to shrink
        ' the radius of training impacts, refining the network by smaller and smaller amounts
        ' on each pass.
        i = i + 1
        If (i = numPixelsPerCycle) Then
            
            'Reset pixel counter (but track how many pixels have been processed - we need this
            ' for progress reporting)
            numPixelsPreviousLoops = numPixelsPreviousLoops + i
            i = 0
            
            'The weight factor
            weight = weight - weight * weightDecreasePerCycle
            
            'Reduce the radius for the next cycle
            fRadius = fRadius * RADIUS_DECREASE
            radius = Int(fRadius)
            If (radius <= 1) Then radius = 0
            
            'Also report any progress updates here; note that we only do this every 8 cycles to improve performance
            currentCycle = currentCycle + 1
            If (Not suppressMessages) And ((currentCycle And &H7&) = 0) Then
                
                'If the caller specified their own progress bar values, scale progress by the image's height
                ' (which is PD's standard progress bar interval).  Otherwise, report progress as the number of
                ' pixels sampled so far.
                If (modifyProgBarMax > 0) Then
                    ProgressBars.SetProgBarVal modifyProgBarOffset + (numPixelsPreviousLoops / numPixelsToSample) * m_srcDIB.GetDIBHeight
                Else
                    ProgressBars.SetProgBarVal modifyProgBarOffset + numPixelsPreviousLoops
                End If
                
            End If
            
        End If
        
    Loop
    
    PDDebug.LogAction "Neural network training complete!"
    
    'Free unsafe array references
    m_srcDIB.UnwrapArrayFromDIB imgPixels
    
End Sub

'Search the network for the best-match BGRA value, while also accounting for bias (which prevents
' us from selecting the same neuron *too* frequently, distorting the map in unwanted ways)
Private Function FindBestNeuron(ByVal b As Single, ByVal g As Single, ByVal r As Single, ByVal a As Single) As Long
    
    'This function will...
    ' 1) Find the most similar neuron to the target color, and update its FREQUENCY
    ' 2) Find the BEST neuron for the target color (most-similar minus bias) and return *it* as the target neuron
    ' 3) As the original code notes, for frequently chosen neurons, frequency will be HIGH while
    '    bias will be NEGATIVE; this discourages the algorithm from selecting the same few neurons over-and-over,
    '    which is essential for distributing the original set of gray-axis neurons into a meaningful new
    '    assortment of colors.
    Dim colorDistance As Single
    Dim tmpR As Single, tmpG As Single, tmpB As Single, tmpA As Single
    
    'Ensure best distances start at maximum values
    Dim bestDistance As Single, bestBiasDistance As Single
    bestDistance = SINGLE_MAX: bestBiasDistance = SINGLE_MAX
    
    Dim idx As Long, idxBestDistance As Long
    
    'beta and gamma are used to increment/decrement the frequency and bias calculations for
    ' each compared neuron
    Const NEURON_BETA As Double = 1# / 1024#
    
    Dim i As Long
    For i = 0 To m_neuronCount - 1
        
        'The original paper uses Manhattan distance instead of Euclidean, and I'm not sure why.
        ' Abs() is an expensive operation - more expensive than multiplication - but things were
        ' different in the early 90's, so maybe it made sense then.  (The paper vaguely alludes
        ' to this: https://web.archive.org/web/20030503154334/http://members.ozemail.com.au/~dekker/NeuQuant.pdf)
        ' Anyway, I've switched to Euclidean here because it's faster.
        
        '(Note also that the use of temporary variables here provides a little speedup,
        ' presumably because modern CPUs can reliably pipeline them.)
        idx = i * 6
        tmpB = m_network(idx) - b
        tmpG = m_network(idx + 1) - g
        tmpR = m_network(idx + 2) - r
        tmpA = m_network(idx + 3) - a
        colorDistance = tmpB * tmpB + tmpG * tmpG + tmpR * tmpR + tmpA * tmpA
        
        'Is this the closest neuron?  If so, mark it.
        If (colorDistance < bestDistance) Then
            bestDistance = colorDistance
            idxBestDistance = i
        End If
        
        'Add bias into our calculation and see if this neuron is *now* the best match
        colorDistance = colorDistance - m_network(idx + 5)
        If (colorDistance < bestBiasDistance) Then
            bestBiasDistance = colorDistance
            FindBestNeuron = i
        End If
    
        'Update frequency and bias trackers for this neuron
        idx = idx + 4
        m_network(idx) = m_network(idx) - m_network(idx) * NEURON_BETA
        m_network(idx + 1) = m_network(idx + 1) + m_network(idx)
        
    Next i
    
    'Before exiting, perform a final frequency and bias update.  (Note that the return was already
    ' calculated above - VB6 uses a different mechanism than standard "return x" syntax.)
    idx = idxBestDistance * 6 + 4
    m_network(idx) = m_network(idx) + NEURON_BETA
    m_network(idx + 1) = m_network(idx + 1) - 1
    
End Function

'Bias adjacent neurons toward the passed BGRA quad using biasFactor as weight,
' and scaling according to the current radius
Private Sub AlterNeighbors(ByVal biasFactor As Single, ByVal radius As Long, ByVal i As Long, ByVal b As Single, ByVal g As Single, ByVal r As Single, ByVal a As Single)
    
    'Start with the neighboring neurons (above and below) the current pixel, then repeat until
    ' we've processed [radius] neurons in each direction.
    Dim idxLow As Long
    idxLow = i - radius
    If (idxLow < -1) Then idxLow = -1
    
    Dim idxHigh As Long
    idxHigh = i + radius
    If (idxHigh > m_neuronCount) Then idxHigh = m_neuronCount
    
    'Setup initial indices (on either side of the target neuron)
    Dim j As Long, k As Long, idx As Long
    j = i + 1
    k = i - 1
    
    'There's a lot of indices being used here, but basically q increments as we move outward
    ' from the target pixel, and we fade the strength of the adjustment as we move further and
    ' further from the center neuron.
    Dim q As Long
    q = 0
    
    'The amount (alpha) that we move each neuron fades as we move outward.  We calculate this
    ' inside the loop, below
    Dim weight As Single
    
    'To improve performance on the inner loop, we square radius (which is the only way we need
    ' to calculate it beyond this point) and we precalculate an inverse 1 / r^2 value
    radius = radius * radius
    
    Dim radiusDivisor As Single
    radiusDivisor = 1! / radius
    
    'Repeat as long as there are neighboring neurons to process
    Do While ((j < idxHigh) Or (k > idxLow))
    
        'Calculate a weight for these neurons
        weight = biasFactor * (radius - q * q) * radiusDivisor
        
        'Increment q (which increases fade on the next pixel)
        q = q + 1
        
        'Apply to the next neuron below (if it's within bounds)
        If (k > idxLow) Then
            idx = k * 6
            m_network(idx) = m_network(idx) - weight * (m_network(idx) - b)
            m_network(idx + 1) = m_network(idx + 1) - weight * (m_network(idx + 1) - g)
            m_network(idx + 2) = m_network(idx + 2) - weight * (m_network(idx + 2) - r)
            m_network(idx + 3) = m_network(idx + 3) - weight * (m_network(idx + 3) - a)
            k = k - 1
        End If
        
        'Apply to the next neuron above (if it's within bounds)
        If (j < idxHigh) Then
            idx = j * 6
            m_network(idx) = m_network(idx) - weight * (m_network(idx) - b)
            m_network(idx + 1) = m_network(idx + 1) - weight * (m_network(idx + 1) - g)
            m_network(idx + 2) = m_network(idx + 2) - weight * (m_network(idx + 2) - r)
            m_network(idx + 3) = m_network(idx + 3) - weight * (m_network(idx + 3) - a)
            j = j + 1
        End If
        
    Loop

End Sub

'When processing is finished, call this to output the finished palette
Friend Sub GetFinalPalette(ByRef dstPalette() As RGBQuad)
    
    ReDim dstPalette(0 To m_neuronCount - 1) As RGBQuad
    Dim b As Long, g As Long, r As Long, a As Long, idx As Long
    
    Dim i As Long
    For i = 0 To m_neuronCount - 1
        
        idx = i * 6
        
        With dstPalette(i)
            
            'Extract each color (and alpha, optionally) and round it to the nearest int
            b = Int(m_network(idx) + 0.5)
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            
            g = Int(m_network(idx + 1) + 0.5)
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            
            r = Int(m_network(idx + 2) + 0.5)
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            
            a = Int(m_network(idx + 3) + 0.5)
            If (a > 255) Then a = 255
            If (a < 0) Then a = 0
            
            .Blue = b
            .Green = g
            .Red = r
            .Alpha = a
            
        End With
        
    Next i
    
End Sub
