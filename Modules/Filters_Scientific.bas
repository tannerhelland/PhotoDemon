Attribute VB_Name = "Filters_Scientific"
'***************************************************************************
'Scientific Filter Collection
'Copyright 2019-2026 by Tanner Helland
'Created: 25/November/19
'Last updated: 25/November/19
'Last update: initial build
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Given a source image, return a map of image angles and magnitudes.  Magnitudes and angles
' are both normalized to [0, 255] scales, and are calculated using Scharr's optimization of a
' Sobel-Feldman operator.
'
'For a magnitude-only version, see the next function below.
Public Sub GetImageGradAndMag(ByRef srcDIB As pdDIB, ByRef dstAngles() As Byte, ByRef dstMagnitudes() As Byte)
    
    'Gradients are calculated in the luminance domain only
    Dim tmpGrayMap() As Byte
    DIBs.GetDIBGrayscaleMap srcDIB, tmpGrayMap, False
    
    'To avoid the need for specialized edge-handling, we are now going to pad the byte array's
    ' edges by 1-px each.
    Dim padGrayMap() As Byte
    Filters_ByteArray.PadByteArray tmpGrayMap, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, padGrayMap, 1, 1
    
    'We no longer need the original map; free it since we need to declare a few more image-sized arrays
    Erase tmpGrayMap
    
    'Padded width/height dimensions (normally we would subtract one because they're used as loop boundaries,
    ' but because we're performing edge-detection, we actually want to subtract *2* - which is the exact
    ' size by which we padded the array, above, so no change is needed!)
    Dim padWidth As Long, padHeight As Long
    padWidth = srcDIB.GetDIBWidth
    padHeight = srcDIB.GetDIBHeight
    
    'We are now going to perform a basic sobel convolution.  We need to do a horizontal and vertical pass;
    ' data from the two passes is then combined to produce usable gradient and magnitude values.
    
    'Start by creating the necessary target arrays.  (Shorts are used because we need to store
    ' positive/negative numbers.)
    Dim hGrayMap() As Integer, vGrayMap() As Integer
    ReDim hGrayMap(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Integer
    
    'Perform a horizontal Sobel on the the horizontal data
    
    'Individual values for the 9 required pixels.  We cache these and reuse them between
    ' pixels to minimize array accesses (and improve cache behavior)
    Dim gLU As Long, gU As Long, gRU As Long
    Dim gL As Long, g As Long, gR As Long
    Dim gLD As Long, gD As Long, gRD As Long
    
    Dim x As Long, y As Long, gSum As Long
    For y = 1 To padHeight
        
        'Precalculate the first set of terms for this line
        gLU = padGrayMap(0, y - 1)
        gU = gLU
        gRU = padGrayMap(1, y - 1)
        gL = padGrayMap(0, y)
        g = gL
        gR = padGrayMap(1, y)
        gLD = padGrayMap(0, y + 1)
        gD = gLD
        gRD = padGrayMap(1, y + 1)
        
    For x = 1 To padWidth
        
        'Reuse previous operators (faster than array hits)
        gLU = gU
        gU = gRU
        gL = g
        g = gR
        gLD = gD
        gD = gRD
        
        'Pull the next three pixel values into the right-side trackers
        gRU = padGrayMap(x + 1, y - 1)
        gR = padGrayMap(x + 1, y)
        gRD = padGrayMap(x + 1, y + 1)
        
        'Use Scharr's optimization for better symmetry (https://en.wikipedia.org/wiki/Sobel_operator#Alternative_operators)
        gSum = 3 * gLU + 10 * gL + 3 * gLD - 3 * gRU - 10 * gR - 3 * gRD
        hGrayMap(x - 1, y - 1) = gSum
        
    Next x
    Next y
    
    'Repeat the above steps on the vertical data
    ReDim vGrayMap(0 To srcDIB.GetDIBWidth + 1, 0 To srcDIB.GetDIBHeight + 1) As Integer
    
    For y = 1 To padHeight
    
        gLU = padGrayMap(0, y - 1)
        gU = gLU
        gRU = padGrayMap(1, y - 1)
        gL = padGrayMap(0, y)
        g = gL
        gR = padGrayMap(1, y)
        gLD = padGrayMap(0, y + 1)
        gD = gLD
        gRD = padGrayMap(1, y + 1)
        
    For x = 1 To padWidth
        
        gLU = gU
        gU = gRU
        gL = g
        g = gR
        gLD = gD
        gD = gRD
        
        gRU = padGrayMap(x + 1, y - 1)
        gR = padGrayMap(x + 1, y)
        gRD = padGrayMap(x + 1, y + 1)
        
        gSum = 3 * gLU + 10 * gU + 3 * gRU - 3 * gLD - 10 * gD - 3 * gRD
        vGrayMap(x - 1, y - 1) = gSum
        
    Next x
    Next y
    
    'With horizontal and vertical gradients calculated, we no longer need our source padded array;
    ' release it to free up memory (as we'll be immediately allocating that memory for the destination
    ' angle and magnitude arrays)
    Erase padGrayMap
    
    'Prep target arrays
    padWidth = srcDIB.GetDIBWidth - 1
    padHeight = srcDIB.GetDIBHeight - 1
    ReDim dstMagnitudes(0 To padWidth, 0 To padHeight) As Byte
    ReDim dstAngles(0 To padWidth, 0 To padHeight) As Byte
    
    'Solve for magnitude and direction
    Const ANG_NORMALIZE As Double = 255# / 6.28318530717959
    Dim hTmp As Long, vTmp As Long, magTmp As Long
    For y = 0 To padHeight
    For x = 0 To padWidth
        
        'Retrieve horizontal and vertical gradients
        hTmp = hGrayMap(x, y)
        vTmp = vGrayMap(x, y)
        
        'Calculate absolute magnitude
        magTmp = Sqr(hTmp * hTmp + vTmp * vTmp)
        
        'Hypothetically, the largest possible magnitude is sqrt(4080 * 4080 * 2), but this doesn't
        ' occurs in practice (because it would require an absolute gradient in both the horizontal
        ' *and* vertical directions simultaneously).  We want to recognize an absolute gradient in
        ' *either* direction as a "maximum", with a little wiggle room; as such, we normalize
        ' against the smaller sqrt(4080 * 4080) / 2 instead, with a failsafe check for overflow.
        magTmp = magTmp \ 8
        If (magTmp > 255) Then magTmp = 255
        dstMagnitudes(x, y) = magTmp
        
        'Calculate a normalized [0, 255] angle (from the [-pi, pi] atan2 return)
        dstAngles(x, y) = Int((PDMath.Atan2_Faster(vTmp, hTmp) + PI) * ANG_NORMALIZE + 0.5)
        
    Next x
    Next y
    
End Sub

'Simplified form of GetImageGradAndMag(), above
Public Sub GetImageGrad_MagOnly(ByRef srcDIB As pdDIB, ByRef dstMagnitudes() As Byte)
    
    'Gradients are calculated in the luminance domain only
    Dim tmpGrayMap() As Byte
    DIBs.GetDIBGrayscaleMap srcDIB, tmpGrayMap, False
    
    'To avoid the need for specialized edge-handling, we are now going to pad the byte array's
    ' edges by 1-px each.
    Dim padGrayMap() As Byte
    Filters_ByteArray.PadByteArray tmpGrayMap, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, padGrayMap, 1, 1
    
    'We no longer need the original map; free it since we need to declare a few more image-sized arrays
    Erase tmpGrayMap
    
    'Padded width/height dimensions (normally we would subtract one because they're used as loop boundaries,
    ' but because we're performing edge-detection, we actually want to subtract *2* - which is the exact
    ' size by which we padded the array, above, so no change is needed!)
    Dim padWidth As Long, padHeight As Long
    padWidth = srcDIB.GetDIBWidth
    padHeight = srcDIB.GetDIBHeight
    
    'We are now going to perform a basic sobel convolution.  We need to do a horizontal and vertical pass;
    ' data from the two passes is then combined to produce usable gradient and magnitude values.
    
    'Start by creating the necessary target arrays.  (Shorts are used because we need to store
    ' positive/negative numbers.)
    Dim hGrayMap() As Integer, vGrayMap() As Integer
    ReDim hGrayMap(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Integer
    
    'Perform a horizontal Sobel on the the horizontal data
    
    'Individual values for the 9 required pixels.  We cache these and reuse them between
    ' pixels to minimize array accesses (and improve cache behavior)
    Dim gLU As Long, gU As Long, gRU As Long
    Dim gL As Long, g As Long, gR As Long
    Dim gLD As Long, gD As Long, gRD As Long
    
    Dim x As Long, y As Long, gSum As Long
    For y = 1 To padHeight
        
        'Precalculate the first set of terms for this line
        gLU = padGrayMap(0, y - 1)
        gU = gLU
        gRU = padGrayMap(1, y - 1)
        gL = padGrayMap(0, y)
        g = gL
        gR = padGrayMap(1, y)
        gLD = padGrayMap(0, y + 1)
        gD = gLD
        gRD = padGrayMap(1, y + 1)
        
    For x = 1 To padWidth
        
        'Reuse previous operators (faster than array hits)
        gLU = gU
        gU = gRU
        gL = g
        g = gR
        gLD = gD
        gD = gRD
        
        'Pull the next three pixel values into the right-side trackers
        gRU = padGrayMap(x + 1, y - 1)
        gR = padGrayMap(x + 1, y)
        gRD = padGrayMap(x + 1, y + 1)
        
        'Use Scharr's optimization for better symmetry (https://en.wikipedia.org/wiki/Sobel_operator#Alternative_operators)
        gSum = 3 * gLU + 10 * gL + 3 * gLD - 3 * gRU - 10 * gR - 3 * gRD
        hGrayMap(x - 1, y - 1) = gSum
        
    Next x
    Next y
    
    'Repeat the above steps on the vertical data
    ReDim vGrayMap(0 To srcDIB.GetDIBWidth + 1, 0 To srcDIB.GetDIBHeight + 1) As Integer
    
    For y = 1 To padHeight
    
        gLU = padGrayMap(0, y - 1)
        gU = gLU
        gRU = padGrayMap(1, y - 1)
        gL = padGrayMap(0, y)
        g = gL
        gR = padGrayMap(1, y)
        gLD = padGrayMap(0, y + 1)
        gD = gLD
        gRD = padGrayMap(1, y + 1)
        
    For x = 1 To padWidth
        
        gLU = gU
        gU = gRU
        gL = g
        g = gR
        gLD = gD
        gD = gRD
        
        gRU = padGrayMap(x + 1, y - 1)
        gR = padGrayMap(x + 1, y)
        gRD = padGrayMap(x + 1, y + 1)
        
        gSum = 3 * gLU + 10 * gU + 3 * gRU - 3 * gLD - 10 * gD - 3 * gRD
        vGrayMap(x - 1, y - 1) = gSum
        
    Next x
    Next y
    
    'With horizontal and vertical gradients calculated, we no longer need our source padded array;
    ' release it to free up memory (as we'll be immediately allocating that memory for the destination
    ' angle and magnitude arrays)
    Erase padGrayMap
    
    'Prep target arrays
    padWidth = srcDIB.GetDIBWidth - 1
    padHeight = srcDIB.GetDIBHeight - 1
    ReDim dstMagnitudes(0 To padWidth, 0 To padHeight) As Byte
    
    'Solve for magnitude only
    Dim hTmp As Long, vTmp As Long, magTmp As Long
    For y = 0 To padHeight
    For x = 0 To padWidth
        
        'Retrieve horizontal and vertical gradients
        hTmp = hGrayMap(x, y)
        vTmp = vGrayMap(x, y)
        
        'Calculate absolute magnitude
        magTmp = Sqr(hTmp * hTmp + vTmp * vTmp)
        
        'Hypothetically, the largest possible magnitude is sqrt(4080 * 4080 * 2), but this doesn't
        ' occur in practice (because it would require an absolute gradient in both the horizontal
        ' *and* vertical directions simultaneously).  We want to recognize an absolute gradient in
        ' *either* direction as a "maximum", with a little wiggle room; as such, we normalize
        ' against the smaller sqrt(4080 * 4080) / 2 instead, with a failsafe check for overflow.
        magTmp = magTmp \ 8
        If (magTmp > 255) Then magTmp = 255
        dstMagnitudes(x, y) = magTmp
        
    Next x
    Next y
    
End Sub

'This sub is for TESTING PURPOSES ONLY!!
Public Sub InternalFFTTest()

    PDDebug.LogAction "Launching FFT test..."
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
        
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim dstRedR() As Single, dstRedI() As Single, dstGreenR() As Single, dstGreenI() As Single, dstBlueR() As Single, dstBlueI() As Single
    GetFFTFromDIB workingDIB, dstRedR, dstRedI, dstGreenR, dstGreenI, dstBlueR, dstBlueI
    
    EffectPrep.FinalizeImageData
    
    Dim timingString As String
    timingString = g_Language.TranslateMessage("Time taken")
    timingString = timingString & ": " & VBHacks.GetTimeDiffNowAsString(startTime)
    Message timingString
    
End Sub

'Given a source DIB, return six single-type FFT arrays: two each (real/imaginary) for R/G/B channels.
' Padding to the nearest power of 2 is handled automatically.
Public Function GetFFTFromDIB(ByRef srcDIB As pdDIB, ByRef dstRedR() As Single, ByRef dstRedI() As Single, ByRef dstGreenR() As Single, ByRef dstGreenI() As Single, ByRef dstBlueR() As Single, ByRef dstBlueI() As Single) As Boolean
    
    'Start by padding the DIB to the nearest power of two in each direction.
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = PDMath.NearestPowerOfTwo(srcDIB.GetDIBWidth)
    targetHeight = PDMath.NearestPowerOfTwo(srcDIB.GetDIBHeight)
    
    PDDebug.LogAction "Prepping input image (" & CStr(targetWidth) & "x" & CStr(targetHeight) & ")..."
    ProgressBars.SetProgBarMax 10
    
    Dim paddedDIB As pdDIB, xOffset As Long, yOffset As Long
    Filters_Layers.PadDIBClampedPixelsEx targetWidth, targetHeight, srcDIB, paddedDIB, xOffset, yOffset
    
    PDDebug.LogAction "Converting to floats..."
    ProgressBars.SetProgBarVal 1
    
    'FFTs take a lot of memory - sorry!
    ReDim dstRedR(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    ReDim dstRedI(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    ReDim dstGreenR(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    ReDim dstGreenI(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    ReDim dstBlueR(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    ReDim dstBlueI(0 To targetWidth - 1, 0 To targetHeight - 1) As Single
    
    'Wrap a generic byte array around the source DIB's pixel bytes
    Dim tmpSA1D As SafeArray1D, imgBytes() As Byte
    paddedDIB.WrapArrayAroundScanline imgBytes, tmpSA1D
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = tmpSA1D.pvData
    dibStride = tmpSA1D.cElements
    
    Dim xStride As Long
    
    Dim r As Long, g As Long, b As Long
    Const ONE_DIV_255 As Single = 1! / 255!
    
    'Copy the image bytes into single-type arrays, and transform to floating-point while we're at it
    Dim x As Long, y As Long
    For y = 0 To targetHeight - 1
        tmpSA1D.pvData = dibPtr + y * dibStride
    For x = 0 To targetWidth - 1
    
        xStride = x * 4
        b = imgBytes(xStride)
        g = imgBytes(xStride + 1)
        r = imgBytes(xStride + 2)
        
        dstRedR(x, y) = CSng(r) * ONE_DIV_255
        dstGreenR(x, y) = CSng(g) * ONE_DIV_255
        dstBlueR(x, y) = CSng(b) * ONE_DIV_255
        
    Next x
    Next y
    
    'Use the FFT class to perform the actual transform
    Dim cFFT As pdFFT
    Set cFFT = New pdFFT
    PDDebug.LogAction "FFT on red channel..."
    ProgressBars.SetProgBarVal 2
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstRedR, dstRedI, True
    PDDebug.LogAction "FFT on green channel..."
    ProgressBars.SetProgBarVal 3
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstGreenR, dstGreenI, True
    PDDebug.LogAction "FFT on blue channel..."
    ProgressBars.SetProgBarVal 4
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstBlueR, dstBlueI, True
    
    'TESTING ONLY!  To confirm that no data is lost in the transfer, perform an immediate reverse transform
    PDDebug.LogAction "Reverse FFT on red channel..."
    ProgressBars.SetProgBarVal 5
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstRedR, dstRedI, False
    PDDebug.LogAction "Reverse FFT on green channel..."
    ProgressBars.SetProgBarVal 6
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstGreenR, dstGreenI, False
    PDDebug.LogAction "Reverse FFT on blue channel..."
    ProgressBars.SetProgBarVal 7
    cFFT.FFT_2D_Radix2_NoTrig targetWidth, targetHeight, dstBlueR, dstBlueI, False
    
    PDDebug.LogAction "Converting back to ints..."
    ProgressBars.SetProgBarVal 8
    
    'Translate the FFT data back into RGB data
    Dim rF As Single, gF As Single, bF As Single
    For y = 0 To targetHeight - 1
        tmpSA1D.pvData = dibPtr + y * dibStride
    For x = 0 To targetWidth - 1
        xStride = x * 4
        
        rF = dstRedR(x, y) * 255!
        gF = dstGreenR(x, y) * 255!
        bF = dstBlueR(x, y) * 255!
        If (rF > 255!) Then rF = 255! Else If (rF < 0!) Then rF = 0!
        If (gF > 255!) Then gF = 255! Else If (gF < 0!) Then gF = 0!
        If (bF > 255!) Then bF = 255! Else If (bF < 0!) Then bF = 0!
        
        imgBytes(xStride) = Int(bF)
        imgBytes(xStride + 1) = Int(gF)
        imgBytes(xStride + 2) = Int(rF)
        
    Next x
    Next y
    
    'Extract the relevant portion of the padded DIB, and place it back in the source DIB
    PDDebug.LogAction "Extracting original (un-padded) portion..."
    ProgressBars.SetProgBarVal 9
    GDI.BitBltWrapper srcDIB.GetDIBDC, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, paddedDIB.GetDIBDC, xOffset, yOffset, vbSrcCopy
    
    PDDebug.LogAction "Done!"
    
    paddedDIB.UnwrapArrayFromDIB imgBytes
    
End Function
