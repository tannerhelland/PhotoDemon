Attribute VB_Name = "Filters_Sci"
Option Explicit

'This sub is for TESTING PURPOSES ONLY!!
Public Sub InternalFFTTest()

    Message "Testing FFT implementation..."
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
        
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Dim dstRedR() As Single, dstRedI() As Single, dstGreenR() As Single, dstGreenI() As Single, dstBlueR() As Single, dstBlueI() As Single
    GetFFTFromDIB workingDIB, dstRedR, dstRedI, dstGreenR, dstGreenI, dstBlueR, dstBlueI
    
    EffectPrep.FinalizeImageData
    
    Message "Time taken: %1", VBHacks.GetTimeDiffNowAsString(startTime)
    
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
