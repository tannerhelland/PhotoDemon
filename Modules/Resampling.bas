Attribute VB_Name = "Resampling"
'***************************************************************************
'Image Resampling engine
'Copyright 2021-2026 by Tanner Helland
'Created: 16/August/21
'Last updated: 19/January/23
'Last update: new floating-point variations that operate on the new pdSurfaceF class (for HDR resampling)
'
'For many years, PhotoDemon relied on external libraries (GDI+, FreeImage) for its resampling algorithms.
' As of v9.0, however, PD now ships with two native resampling engines (one floating-point-based, one integer-based).
' These native resampling engines support many more resampling filters, and their quality is excellent while
' maintaining impressive performance (especially for VB6!).
'
'The general design of PD's floating-point resampling engine is adopted from a resampling project originally
' written by Libor Tinka. Libor shared his original C# resampling code under a Code Project Open License (CPOL):
'  - https://www.codeproject.com/info/cpol10.aspx
' His original, unmodified resampling source code is available here (link good as of Aug 2021):
'  - https://www.codeproject.com/Articles/11143/Image-Resizing-outperform-GDI
' Thank you to Libor for sharing his informative C# image resampling project.  (Note that a number of
' critical bug-fixes are addressed in PD's version of the code; details are in the comments.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Timing reports are helpful during debugging.  Do not enable in production.
Private Const REPORT_RESAMPLE_PERF As Boolean = False, REPORT_DETAILED_PERF As Boolean = False

'Float- and integer-based methods are tracked separately
Private m_NetTimeF As Double, m_IterationsF As Long
Private m_NetTimeI As Double, m_IterationsI As Long

'Currently available resamplers
Public Enum PD_ResamplingFilter
    
    'Dummy entry that auto-maps to other filters depending on resize dimensions
    rf_Automatic = 0
    
    'r = 0.5
    rf_Box
    
    'r = 1.0
    rf_BilinearTriangle
    rf_Cosine
    rf_Hermite
    
    'r = 1.5
    rf_Bell
    rf_Quadratic
    rf_QuadraticBSpline
    
    'r = 2
    rf_CubicBSpline
    rf_CatmullRom
    rf_Mitchell
    
    'r = 3
    rf_CubicConvolution
    
    'r is variable; default is 3
    rf_Lanczos
    
    'Dummy entry to allow filter iteration
    rf_Max
    
End Enum

#If False Then
    Private Const rf_Automatic = 0, rf_Box = 0, rf_BilinearTriangle = 0, rf_Hermite = 0, rf_Bell = 0, rf_CubicBSpline = 0, rf_Lanczos = 0, rf_Mitchell = 0
    Private Const rf_Cosine = 0, rf_CatmullRom = 0, rf_Quadratic = 0, rf_QuadraticBSpline = 0, rf_CubicConvolution = 0
    Private Const rf_Max = 0
#End If

'Weight calculation
Private Type Contributor
    pixel As Long
    weight As Single    'Because this is declared as an array (see below), single provides a perf advantage over double
End Type
    
Private Type ContributorEntry
    nCount As Long
    weightSum As Single
    p() As Contributor
End Type

'Different structs are used by the integer-only version of the transform
Private Type ContributorI
    pixel As Long
    weight As Long
End Type
    
Private Type ContributorEntryI
    nCount As Long
    weightSum As Long
    p() As ContributorI
End Type

'Lanczos supports variable lobes currently locked on the range 1-10; note that performance scales O(2n) against
' this value (it's separable), so larger radii require longer processing times.
Private Const LANCZOS_DEFAULT As Long = 3, LANCZOS_MIN As Long = 2, LANCZOS_MAX As Long = 10
Private m_LanczosRadius As Long

'Our current resampling approach uses an intermediate copy of the image; this allows us to handle x and y
' resampling independently (which improves performance and greatly simplifies the code, at some trade-off to
' memory consumption).  This intermediate array will be reused on subsequent calls, and can also be manually
' freed when bulk resizing completes.
Private m_tmpPixels() As Byte, m_tmpPixelSize As Long
Private m_tmpPixelsF() As Single, m_tmpPixelSizeF As Long
Private m_tmpPixelsL() As Long, m_tmpPixelSizeL As Long

'Freeing the temporary resize buffers also resets perf trackers
Public Sub FreeBuffers()
    m_tmpPixelSize = 0
    Erase m_tmpPixels
    m_tmpPixelSizeF = 0
    Erase m_tmpPixelsF
    m_tmpPixelSizeL = 0
    Erase m_tmpPixelsL
    m_NetTimeF = 0#
    m_NetTimeI = 0#
    m_IterationsF = 0
    m_IterationsI = 0
End Sub

'The radius of most resample filters is fixed.  Lanczos is an exception since it's a windowed approximation of a "perfect"
' Sinc function (and technically this changes the lobe count, not the radius FYI).  Larger radius doesn't necessarily
' correlate with "better visual results"; it's more about theoretical accuracy which may or may not translate to
' something a human prefers - so don't crank this value up unnecessarily.
Public Function GetLanczosRadius() As Long
    GetLanczosRadius = m_LanczosRadius
End Function

Public Sub SetLanczosRadius(ByVal newRadius As Long)
    If (newRadius < LANCZOS_MIN) Or (newRadius > LANCZOS_MAX) Then
        m_LanczosRadius = LANCZOS_DEFAULT
    Else
        m_LanczosRadius = newRadius
    End If
End Sub

Public Function GetResamplerName(ByVal rsID As PD_ResamplingFilter) As String

    Select Case rsID
        Case rf_Automatic
            GetResamplerName = "auto"
        Case rf_Box
            GetResamplerName = "nearest"
        Case rf_BilinearTriangle
            GetResamplerName = "bilinear"
        Case rf_Cosine
            GetResamplerName = "cosine"
        Case rf_Hermite
            GetResamplerName = "hermite"
        Case rf_Bell
            GetResamplerName = "bell"
        Case rf_Quadratic
            GetResamplerName = "quadratic"
        Case rf_QuadraticBSpline
            GetResamplerName = "quadratic-spline"
        Case rf_CubicBSpline
            GetResamplerName = "bicubic"
        Case rf_CatmullRom
            GetResamplerName = "catmull"
        Case rf_Mitchell
            GetResamplerName = "mitchell"
        Case rf_CubicConvolution
            GetResamplerName = "cubic-convolve"
        Case rf_Lanczos
            GetResamplerName = "lanczos"
    End Select

End Function

Public Function GetResamplerNameUI(ByVal rsID As PD_ResamplingFilter) As String

    Select Case rsID
        Case rf_Automatic
            GetResamplerNameUI = g_Language.TranslateMessage("automatic")
        Case rf_Box
            GetResamplerNameUI = g_Language.TranslateMessage("nearest-neighbor")
        Case rf_BilinearTriangle
            GetResamplerNameUI = g_Language.TranslateMessage("bilinear")
        Case rf_Cosine
            GetResamplerNameUI = g_Language.TranslateMessage("cosine")
        Case rf_Hermite
            GetResamplerNameUI = "Hermite"
        Case rf_Bell
            GetResamplerNameUI = g_Language.TranslateMessage("bell")
        Case rf_Quadratic
            GetResamplerNameUI = g_Language.TranslateMessage("quadratic")
        Case rf_QuadraticBSpline
            GetResamplerNameUI = g_Language.TranslateMessage("quadratic b-spline")
        Case rf_CubicBSpline
            GetResamplerNameUI = g_Language.TranslateMessage("bicubic")
        Case rf_CatmullRom
            GetResamplerNameUI = "Catmull-Rom"
        Case rf_Mitchell
            GetResamplerNameUI = "Mitchell-Netravali"
        Case rf_CubicConvolution
            GetResamplerNameUI = g_Language.TranslateMessage("cubic convolution")
        Case rf_Lanczos
            GetResamplerNameUI = "Lanczos"
    End Select

End Function

Public Function GetResamplerID(ByRef rsName As String) As PD_ResamplingFilter

    Select Case LCase$(rsName)
        Case "auto", "automatic"
            GetResamplerID = rf_Automatic
        Case "nearest"
            GetResamplerID = rf_Box
        Case "bilinear"
            GetResamplerID = rf_BilinearTriangle
        Case "cosine"
            GetResamplerID = rf_Cosine
        Case "hermite"
            GetResamplerID = rf_Hermite
        Case "bell"
            GetResamplerID = rf_Bell
        Case "quadratic"
            GetResamplerID = rf_Quadratic
        Case "quadratic-spline"
            GetResamplerID = rf_QuadraticBSpline
        Case "bicubic"
            GetResamplerID = rf_CubicBSpline
        Case "catmull"
            GetResamplerID = rf_CatmullRom
        Case "mitchell"
            GetResamplerID = rf_Mitchell
        Case "cubic-convolve"
            GetResamplerID = rf_CubicConvolution
        Case "lanczos"
            GetResamplerID = rf_Lanczos
        Case Else
            GetResamplerID = rf_Automatic
    End Select

End Function

'Resample an image using a variety of resampling filters.  A few notes...
' 1) This implementation uses separable resampling, which means resampling occurs in two passes (one in each direction).
'     This provides performance O(2n) vs O(n^2).
' 2) Resampling requires an intermediary copy to store a copy of the data resampled by the first pass.  The code
'     attempts to reuse the same buffer between calls; to free this shared buffer, call FreeBuffers().  The size of
'     this buffer is (newWidth x oldHeight x numOfChannels).
' 3) When two-dimensional resampling is required, the x-dimension will be resampled first.  This is by design as
'     VB array allocation favors locality in the x-direction, so it's generally faster to do the largest computational
'     pass in the X-direction.
' 4) 32-bpp 4-channel inputs are required.  All channels are resampled using identical code and weights.
' 5) Some resampling kernels can produce values outside the [0, 255] range.  PD catches and clamps these cases automatically,
'    but for improved performance, you could write yet *another* resample function that only enables these checks on
'    kernels that require them.
' 6) Consider alpha premultiplication state when using this function.  As with any area function, premultiplication is
'    generally advised, as you don't want the (likely arbitrary) color of transparent pixels bleeding into neighboring
'    opaque pixels.  However, caution is warranted here as clamping can produce unpredictable results that may violate
'    premultiplication state, particularly when alpha is 0.  (A post-resampling toggle of alpha premultiplion off/on
'    can rectify this if necessary.)
Public Function ResampleImage(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal rsFilter As PD_ResamplingFilter, Optional ByVal displayProgress As Boolean = False) As Boolean
    
    ResampleImage = False
    Const FUNC_NAME As String = "ResampleImage"
    If REPORT_RESAMPLE_PERF Then PDDebug.LogAction "Float resampler started."
    
    'Validate all inputs
    If (srcDIB Is Nothing) Then
        InternalError FUNC_NAME, "null source"
        Exit Function
    End If
    
    If (dstWidth <= 0) Or (dstHeight <= 0) Then
        InternalError FUNC_NAME, "bad width/height: " & dstWidth & ", " & dstHeight
        Exit Function
    End If
    
    'Initialize destination as necessary
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> dstWidth) Or (dstDIB.GetDIBHeight <> dstHeight) Then dstDIB.CreateBlank dstWidth, dstHeight, srcDIB.GetDIBColorDepth, 0, 0
    
    'Performance reporting (via debug logs) is controlled by constants at the top of this class
    Dim startTime As Currency, firstTime As Currency
    VBHacks.GetHighResTime startTime
    firstTime = startTime
    
    'Validate all internal values as well
    If (m_LanczosRadius < LANCZOS_MIN) Or (m_LanczosRadius > LANCZOS_MAX) Then m_LanczosRadius = LANCZOS_DEFAULT
    
    'Inputs look good.  Prepare intermediary data structs.  Custom types are used to improve memory locality.
    Dim srcWidth As Long: srcWidth = srcDIB.GetDIBWidth
    Dim srcHeight As Long: srcHeight = srcDIB.GetDIBHeight
    
    'Allocate the intermediary "working" copy of the image width dimensions [dstWidth, srcHeight].
    ' Note that a 32-bpp BGRA structure is *always* assumed.  (It would be trivial to extend this to
    ' floating-point or higher bit-depths, but at present, 32-bpp works well!)
    If (dstWidth * srcHeight * 4 > m_tmpPixelSize) Then
        m_tmpPixelSize = dstWidth * srcHeight * 4
        ReDim m_tmpPixels(0 To m_tmpPixelSize - 1) As Byte
    Else
        VBHacks.ZeroMemory VarPtr(m_tmpPixels(0)), m_tmpPixelSize
    End If
    
    'Calculate x/y scales; this provides a simple mechanism for checking up- vs downsampling
    ' in either direction.
    Dim xScale As Double, yScale As Double
    xScale = dstWidth / srcWidth
    yScale = dstHeight / srcHeight
    
    'If progress bar reports are wanted, calculate max values now
    Dim progX As Long, progY As Long, progBarCheck As Long
    If displayProgress Then
        If (xScale <> 1#) Then progX = srcHeight
        If (yScale <> 1#) Then progY = dstWidth
        ProgressBars.SetProgBarMax progX + progY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep array of contributors (which contain per-pixel weights)
    Dim contrib() As ContributorEntry
    ReDim contrib(0 To dstWidth - 1) As ContributorEntry
    
    Dim radius As Double, center As Double, weight As Double
    Dim intensityR As Double, intensityG As Double, intensityB As Double, intensityA As Double
    Dim pxLeft As Long, pxRight As Long, i As Long, j As Long, k As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim xOffset As Long
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Now, calculate all input weights
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Horizontal downsampling
    If (xScale < 1#) Then
        
        'The source width is larger than the destination width
        radius = (GetDefaultRadius(rsFilter) / xScale)
        
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column (and ensure valid boundaries)
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                'Calculate weight for this pixel, according to the current filter
                weight = GetValue(rsFilter, (center - j - 0.5) * xScale)
                
                'For a "perfect" implementation, you would want to include the weight of all contributing
                ' pixels, regardless of how minute they are.  Because we're working with 32-bpp data, however,
                ' and clamping between the horizontal and vertical passes, there is no compelling reason to
                ' include values that contribute less than 1 integer value of a pixel's potential color (1/255)
                ' to the final result.  Ignoring extreme tails of the distribution provides a nice performance
                ' improvement with no meaningful change to the final result (across all resampling algorithms).
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * 4   'Offset by 32-bits-per-pixel in advance
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).weightSum = contrib(i).weightSum + weight
                    contrib(i).nCount = contrib(i).nCount + 1
                End If
            
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    'Horizontal upsampling
    ElseIf (xScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source width is smaller than the destination width
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * 4   'Offset by 32-bits-per-pixel in advance
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).weightSum = contrib(i).weightSum + weight
                    contrib(i).nCount = contrib(i).nCount + 1
                End If
            
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    '/End up- vs downsampling.  Note that the special case of xScale = 1.0 (e.g. horizontal width
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage horizontal 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter horizontally from the input image
    ' to the temporary "working" copy.
    Dim imgPixels() As Byte, srcSA As SafeArray1D
    Dim idxPixel As Long, wSum As Double
    
    'If the image is changing size, perform resampling now
    If (xScale <> 1#) Then
        
        'Each row (source image height)...
        For k = 0 To srcHeight - 1
            
            'Wrap a VB array around the image at this line
            srcDIB.WrapArrayAroundScanline imgPixels, srcSA, k
            
            'Each column (destination image width)...
            For i = 0 To dstWidth - 1
                
                intensityB = 0#
                intensityG = 0#
                intensityR = 0#
                intensityA = 0#
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel
                    intensityB = intensityB + (imgPixels(idxPixel) * weight)
                    intensityG = intensityG + (imgPixels(idxPixel + 1) * weight)
                    intensityR = intensityR + (imgPixels(idxPixel + 2) * weight)
                    intensityA = intensityA + (imgPixels(idxPixel + 3) * weight)
                Next j
                
                'Weight and clamp final RGBA values.  (Note that normally you'd *divide* by the
                ' weighted sum here, but we already normalized that value in a previous step.)
                wSum = contrib(i).weightSum
                
                b = Int(intensityB * wSum + 0.5)
                g = Int(intensityG * wSum + 0.5)
                r = Int(intensityR * wSum + 0.5)
                a = Int(intensityA * wSum + 0.5)
                
                If (b > 255) Then b = 255
                If (g > 255) Then g = 255
                If (r > 255) Then r = 255
                If (a > 255) Then a = 255
                
                If (b < 0) Then b = 0
                If (g < 0) Then g = 0
                If (r < 0) Then r = 0
                If (a < 0) Then a = 0
                
                'Assign new RGBA values to the working data array
                idxPixel = k * dstWidth * 4 + i * 4
                m_tmpPixels(idxPixel) = b
                m_tmpPixels(idxPixel + 1) = g
                m_tmpPixels(idxPixel + 2) = r
                m_tmpPixels(idxPixel + 3) = a
                
            'Next pixel in row...
            Next i
            
            'Report progress
            If displayProgress And ((k And progBarCheck) = 0) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k
            End If
            
        'Next row in image...
        Next k
    
        'Free any unsafe references
        srcDIB.UnwrapArrayFromDIB imgPixels
    
    'If the image's horizontal size *isn't* changing, just mirror the data into the temporary array.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination in the next step - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict VarPtr(m_tmpPixels(0)), srcDIB.GetDIBPointer, srcHeight * dstWidth * 4
    End If
    
    'Horizontal sampling is now complete.
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage horizontal 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Next, we need to perform nearly identical sampling from the "working" copy to the destination image,
    ' while resampling in the y-direction.
    
    'Reset contributor weight table (one entry per row for vertical resampling)
    ReDim contrib(0 To dstHeight - 1) As ContributorEntry
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage vertical prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Vertical downsampling
    If (yScale < 1#) Then
        
        'The source height is larger than the destination height
        radius = GetDefaultRadius(rsFilter) / yScale
        
        'Iterate through each row in the image
        For i = 0 To dstHeight - 1
          
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
          
            center = (i + 0.5) / yScale
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            'Precalculate all weights for this column (technically these are not left/right values
            ' but up/down ones, remember)
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, (center - j - 0.5) * yScale)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).weightSum = contrib(i).weightSum + weight
                    contrib(i).nCount = contrib(i).nCount + 1
                End If
            
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    'Vertical upsampling
    ElseIf (yScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source height is smaller than the destination height
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            center = ((i + 0.5) / yScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.9999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).weightSum = contrib(i).weightSum + weight
                    contrib(i).nCount = contrib(i).nCount + 1
                End If
            
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    '/End up- vs downsampling.  Note that the special case of yScale = 1.0 (e.g. vertical height
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage vertical 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter vertically from the "working copy"
    ' to the destination image.
    If (yScale <> 1#) Then
        
        'Because we need to access pixels across rows, it's easiest to just wrap a single array around
        'the entire image.
        dstDIB.WrapArrayAroundDIB_1D imgPixels, srcSA
        
        'Each column (new image width)...
        For k = 0 To dstWidth - 1
            
            'Pre-calculate a fixed x-offset for this column
            xOffset = k * 4
            
            'Each row (destination image height)...
            For i = 0 To dstHeight - 1
                
                intensityB = 0#
                intensityG = 0#
                intensityR = 0#
                intensityA = 0#
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel + xOffset
                    intensityB = intensityB + (m_tmpPixels(idxPixel) * weight)
                    intensityG = intensityG + (m_tmpPixels(idxPixel + 1) * weight)
                    intensityR = intensityR + (m_tmpPixels(idxPixel + 2) * weight)
                    intensityA = intensityA + (m_tmpPixels(idxPixel + 3) * weight)
                Next j
                
                'Weight and clamp final RGBA values
                wSum = contrib(i).weightSum
                
                b = Int(intensityB * wSum + 0.5)
                g = Int(intensityG * wSum + 0.5)
                r = Int(intensityR * wSum + 0.5)
                a = Int(intensityA * wSum + 0.5)
                
                If (b > 255) Then b = 255
                If (g > 255) Then g = 255
                If (r > 255) Then r = 255
                If (a > 255) Then a = 255
                
                If (b < 0) Then b = 0
                If (g < 0) Then g = 0
                If (r < 0) Then r = 0
                If (a < 0) Then a = 0
                
                'Assign new RGBA values to the working data array
                idxPixel = (k * 4) + (i * dstWidth * 4)
                imgPixels(idxPixel) = b
                imgPixels(idxPixel + 1) = g
                imgPixels(idxPixel + 2) = r
                imgPixels(idxPixel + 3) = a
                
            'Next row...
            Next i
            
            'Report progress
            If (displayProgress And (((k + progX) And progBarCheck) = 0)) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k + progX
            End If
            
        'Next column...
        Next k
        
        'Release all unsafe references
        dstDIB.UnwrapArrayFromDIB imgPixels
        
    'If the image's vertical size *isn't* changing, just mirror the intermediate data into dstImage.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination earlier - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict dstDIB.GetDIBPointer, VarPtr(m_tmpPixels(0)), dstHeight * dstWidth * 4
    End If
    
    If REPORT_RESAMPLE_PERF Then
        If REPORT_DETAILED_PERF Then PDDebug.LogAction "Resampling.ResampleImage vertical 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        m_NetTimeF = m_NetTimeF + VBHacks.GetTimerDifferenceNow(firstTime)
        m_IterationsF = m_IterationsF + 1
        PDDebug.LogAction "Average resampling time for this session (F): " & VBHacks.GetTotalTimeAsString((m_NetTimeF / m_IterationsF) * 1000)
    End If
    
    'Resampling complete!
    ResampleImage = True
    
End Function

'Resample an image using integer-based transforms.  This function has 2x or better performance than the "pure"
' floating-point version, above, but at a slight hit to quality.  Whether or not the differences are noticeable
' is ultimately up to the viewer, but PSNR suggests the differences are not meaningful (99.8-99.9% similarity to
' the floating-point version).
'
'Finally, note that all the same caveats as the original function apply, especially regarding input/output formats.
Public Function ResampleImageI(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal rsFilter As PD_ResamplingFilter, Optional ByVal displayProgress As Boolean = False) As Boolean
    
    ResampleImageI = False
    Const FUNC_NAME As String = "ResampleImageI"
    If REPORT_RESAMPLE_PERF Then PDDebug.LogAction "Integer resampler started."
    
    'Validate all inputs
    If (srcDIB Is Nothing) Then
        InternalError FUNC_NAME, "null source"
        Exit Function
    End If
    
    If (dstWidth <= 0) Or (dstHeight <= 0) Then
        InternalError FUNC_NAME, "bad width/height: " & dstWidth & ", " & dstHeight
        Exit Function
    End If
    
    'Initialize destination as necessary
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> dstWidth) Or (dstDIB.GetDIBHeight <> dstHeight) Then dstDIB.CreateBlank dstWidth, dstHeight, srcDIB.GetDIBColorDepth, 0, 0
    
    'Performance reporting (via debug logs) is controlled by constants at the top of this class
    Dim startTime As Currency, firstTime As Currency
    VBHacks.GetHighResTime startTime
    firstTime = startTime
    
    'Validate all internal values as well
    If (m_LanczosRadius < LANCZOS_MIN) Or (m_LanczosRadius > LANCZOS_MAX) Then m_LanczosRadius = LANCZOS_DEFAULT
    
    'Inputs look good.  Prepare intermediary data structs.  Custom types are used to improve memory locality.
    Dim srcWidth As Long: srcWidth = srcDIB.GetDIBWidth
    Dim srcHeight As Long: srcHeight = srcDIB.GetDIBHeight
    
    'Allocate the intermediary "working" copy of the image width dimensions [dstWidth, srcHeight].
    ' Note that a 32-bpp BGRA structure is *always* assumed.  (It would be trivial to extend this to
    ' floating-point or higher bit-depths, but at present, 32-bpp works well!)
    If (dstWidth * srcHeight * 4 > m_tmpPixelSize) Then
        m_tmpPixelSize = dstWidth * srcHeight * 4
        ReDim m_tmpPixels(0 To m_tmpPixelSize - 1) As Byte
    Else
        VBHacks.ZeroMemory VarPtr(m_tmpPixels(0)), m_tmpPixelSize
    End If
    
    'Calculate x/y scales; this provides a simple mechanism for checking up- vs downsampling
    ' in either direction.
    Dim xScale As Double, yScale As Double
    xScale = dstWidth / srcWidth
    yScale = dstHeight / srcHeight
    
    'If progress bar reports are wanted, calculate max values now
    Dim progX As Long, progY As Long, progBarCheck As Long
    If displayProgress Then
        If (xScale <> 1#) Then progX = srcHeight
        If (yScale <> 1#) Then progY = dstWidth
        ProgressBars.SetProgBarMax progX + progY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep array of contributors (which contain per-pixel weights)
    Dim contrib() As ContributorEntry, contribI() As ContributorEntryI
    ReDim contrib(0 To dstWidth - 1) As ContributorEntry
    ReDim contribI(0 To dstWidth - 1) As ContributorEntryI
    
    Dim radius As Double, center As Double, weight As Double, weightI As Long
    Dim iR As Long, iG As Long, iB As Long, iA As Long
    Dim pxLeft As Long, pxRight As Long, i As Long, j As Long, k As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim xOffset As Long, weightMax As Double, intFactor As Long
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Now, calculate all input weights.  (Unique to this integer version is that we need to track the
    ' maximum weight value in the table; we'll use this to normalize all weights against LONG_MAX.)
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'VB6's ancient compiler is smart enough to swap-in bit-shifts for fixed powers-of-two.  Because this is
    ' a COMPILE-time optimization, we can't calculate an ideal power-of-two at run-time - instead, we need
    ' to use a fixed value that's big enough to minimize quality loss, but small enough to ensure we do
    ' not accidentally overflow when resizing drastically different amounts (e.g. 1x1 <-> 5000x5000).
    ' Fortunately, PD's support for image sizes is predictable, so we can guarantee a good power-of-two
    ' at compile-time, which lets the compiler generate fast bit-shifts instead of integer divides on
    ' the inner loop!
    Dim oldSum As Long
    Const LARGE_POWER_OF_TWO As Long = 2097152   '2 ^ 21 is a good compromise between expected range and high safety margin
    
    'Horizontal downsampling
    If (xScale < 1#) Then
        
        'The source width is larger than the destination width
        radius = (GetDefaultRadius(rsFilter) / xScale)
        
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            weightMax = 0#
            
            'Calculate center/left/right for this column (and ensure valid boundaries)
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                'Calculate weight for this pixel, according to the current filter
                weight = GetValue(rsFilter, (center - j - 0.5) * xScale)
                
                'For a "perfect" implementation, you would want to include the weight of all contributing
                ' pixels, regardless of how minute they are.  Because we're working with 32-bpp data, however,
                ' and clamping between the horizontal and vertical passes, there is no compelling reason to
                ' include values that contribute less than 1 integer value of a pixel's potential color (1/255)
                ' to the final result.  Ignoring extreme tails of the distribution provides a nice performance
                ' improvement with no meaningful change to the final result (across all resampling algorithms).
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * 4   'Offset by 32-bits-per-pixel in advance
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).nCount = contrib(i).nCount + 1
                    If (weight > weightMax) Then weightMax = weight
                End If
            
            Next j
            
            'Now that we know the maximum weight of the filter, we can determine an integer scale-factor
            ' that guarantees overflow safety for this table while maximizing integer estimation "accuracy".
            intFactor = LONG_MAX \ Int(weightMax * contrib(i).nCount * 256 + 0.999999999999999)
            
            'Rebuild the weight table using the newly determined integer factor.
            ReDim contribI(i).p(0 To contrib(i).nCount - 1) As ContributorI
            contribI(i).nCount = contrib(i).nCount
            
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
                contribI(i).weightSum = contribI(i).weightSum + contribI(i).p(j).weight
            Next j
            
            'The existing formula works as-is, but for even better performance we can do one additional change.
            ' Instead of using a custom weight for this table, let's use a fixed power-of-two which allows the
            ' compiler to preferentially choose bit-shifts instead of integer divides when resolving table weights.
            ' This provides a large relative speed-up at a minor hit to calculation quality.
            oldSum = contribI(i).weightSum
            contribI(i).weightSum = LARGE_POWER_OF_TWO
            
            'Because changing the weighted sum has also changed the scale of the entire calculation,
            ' we need to take one last pass through the data to reassign it to this new numerical range.
            ' (Again, this imposes a minor quality hit, but greatly improves relative performance.)
            intFactor = Int(CDbl(intFactor) * (CDbl(contribI(i).weightSum) / CDbl(oldSum)))
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
            Next j
            
        Next i
    
    'Horizontal upsampling
    ElseIf (xScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source width is smaller than the destination width
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            weightMax = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * 4   'Offset by 32-bits-per-pixel in advance
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).nCount = contrib(i).nCount + 1
                    If (weight > weightMax) Then weightMax = weight
                End If
            
            Next j
            
            'Scale against LONG_MAX and rebuild as an integer weight table.
            intFactor = LONG_MAX \ Int(weightMax * contrib(i).nCount * 256 + 0.999999999999999)
            
            ReDim contribI(i).p(0 To contrib(i).nCount - 1) As ContributorI
            contribI(i).nCount = contrib(i).nCount
            
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
                contribI(i).weightSum = contribI(i).weightSum + contribI(i).p(j).weight
            Next j
            
            'Finally, scale the final integer results against a fixed power-of-two for even better performance.
            ' (See notes, above, for further details.)
            oldSum = contribI(i).weightSum
            contribI(i).weightSum = LARGE_POWER_OF_TWO
            
            intFactor = Int(CDbl(intFactor) * (CDbl(contribI(i).weightSum) / CDbl(oldSum)))
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
            Next j
            
        Next i
    
    '/End up- vs downsampling.  Note that the special case of xScale = 1.0 (e.g. horizontal width
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    'Note that we no longer need the floating-point weight table, but we're going to leave it
    ' allocated because we need to construct a similar table for the vertical resize, below.
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage horizontal 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter horizontally from the input image
    ' to the temporary "working" copy.
    Dim imgPixels() As Byte, srcSA As SafeArray1D
    Dim idxPixel As Long
    
    'If the image is changing size, perform resampling now
    If (xScale <> 1#) Then
        
        'Each row (source image height)...
        For k = 0 To srcHeight - 1
            
            'Wrap a VB array around the image at this line
            srcDIB.WrapArrayAroundScanline imgPixels, srcSA, k
            xOffset = k * dstWidth * 4
            
            'Each column (destination image width)...
            For i = 0 To dstWidth - 1
                
                iB = 0
                iG = 0
                iR = 0
                iA = 0
                
                'Generate weighted result for each color component
                For j = 0 To contribI(i).nCount - 1
                    weightI = contribI(i).p(j).weight
                    idxPixel = contribI(i).p(j).pixel
                    iB = iB + weightI * imgPixels(idxPixel)
                    iG = iG + weightI * imgPixels(idxPixel + 1)
                    iR = iR + weightI * imgPixels(idxPixel + 2)
                    iA = iA + weightI * imgPixels(idxPixel + 3)
                Next j
                
                'Weight and clamp final RGBA values.
                b = iB \ LARGE_POWER_OF_TWO
                g = iG \ LARGE_POWER_OF_TWO
                r = iR \ LARGE_POWER_OF_TWO
                a = iA \ LARGE_POWER_OF_TWO
                
                If (b > 255) Then b = 255
                If (g > 255) Then g = 255
                If (r > 255) Then r = 255
                If (a > 255) Then a = 255
                
                If (b < 0) Then b = 0
                If (g < 0) Then g = 0
                If (r < 0) Then r = 0
                If (a < 0) Then a = 0
                
                'Assign new RGBA values to the working data array
                idxPixel = xOffset + i * 4
                m_tmpPixels(idxPixel) = b
                m_tmpPixels(idxPixel + 1) = g
                m_tmpPixels(idxPixel + 2) = r
                m_tmpPixels(idxPixel + 3) = a
                
            'Next pixel in row...
            Next i
            
            'Report progress
            If displayProgress And ((k And progBarCheck) = 0) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k
            End If
            
        'Next row in image...
        Next k
    
        'Free any unsafe references
        srcDIB.UnwrapArrayFromDIB imgPixels
    
    'If the image's horizontal size *isn't* changing, just mirror the data into the temporary array.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination in the next step - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict VarPtr(m_tmpPixels(0)), srcDIB.GetDIBPointer, srcHeight * dstWidth * 4
    End If
    
    'Horizontal sampling is now complete.
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage horizontal 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Next, we need to perform nearly identical sampling from the "working" copy to the destination image,
    ' while resampling in the y-direction.
    
    'Reset contributor weight table (one entry per row for vertical resampling)
    ReDim contrib(0 To dstHeight - 1) As ContributorEntry
    ReDim contribI(0 To dstHeight - 1) As ContributorEntryI
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage vertical prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Vertical downsampling
    If (yScale < 1#) Then
        
        'The source height is larger than the destination height
        radius = GetDefaultRadius(rsFilter) / yScale
        
        'Iterate through each row in the image
        For i = 0 To dstHeight - 1
          
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            weightMax = 0#
          
            center = (i + 0.5) / yScale
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            'Precalculate all weights for this column (technically these are not left/right values
            ' but up/down ones, remember)
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, (center - j - 0.5) * yScale)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).nCount = contrib(i).nCount + 1
                    If (weight > weightMax) Then weightMax = weight
                End If
            
            Next j
            
            'Scale against LONG_MAX and rebuild as an integer weight table.
            intFactor = LONG_MAX \ Int(weightMax * contrib(i).nCount * 256 + 0.999999999999999)
            
            ReDim contribI(i).p(0 To contrib(i).nCount - 1) As ContributorI
            contribI(i).nCount = contrib(i).nCount
            
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
                contribI(i).weightSum = contribI(i).weightSum + contribI(i).p(j).weight
            Next j
            
            'Finally, scale the final integer results against a fixed power-of-two for even better performance.
            ' (See notes, above, for further details.)
            oldSum = contribI(i).weightSum
            contribI(i).weightSum = LARGE_POWER_OF_TWO
            
            intFactor = Int(CDbl(intFactor) * (CDbl(contribI(i).weightSum) / CDbl(oldSum)))
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
            Next j
            
        'Next row...
        Next i
    
    'Vertical upsampling
    ElseIf (yScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source height is smaller than the destination height
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            weightMax = 0#
            
            center = ((i + 0.5) / yScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.9999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                If (Abs(weight) > ONE_DIV_255) Then
                    contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                    contrib(i).p(contrib(i).nCount).weight = weight
                    contrib(i).nCount = contrib(i).nCount + 1
                    If (weight > weightMax) Then weightMax = weight
                End If
            
            Next j
            
            'Scale against LONG_MAX and rebuild as an integer weight table.
            intFactor = LONG_MAX \ Int(weightMax * contrib(i).nCount * 256 + 0.999999999999999)
            
            ReDim contribI(i).p(0 To contrib(i).nCount - 1) As ContributorI
            contribI(i).nCount = contrib(i).nCount
            
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
                contribI(i).weightSum = contribI(i).weightSum + contribI(i).p(j).weight
            Next j
            
            'Finally, scale the final integer results against a fixed power-of-two for even better performance.
            ' (See notes, above, for further details.)
            oldSum = contribI(i).weightSum
            contribI(i).weightSum = LARGE_POWER_OF_TWO
            
            intFactor = Int(CDbl(intFactor) * (CDbl(contribI(i).weightSum) / CDbl(oldSum)))
            For j = 0 To contribI(i).nCount - 1
                contribI(i).p(j).pixel = contrib(i).p(j).pixel
                contribI(i).p(j).weight = Int(contrib(i).p(j).weight * intFactor)
            Next j
            
        'Next row...
        Next i
    
    '/End up- vs downsampling.  Note that the special case of yScale = 1.0 (e.g. vertical height
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImage vertical 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Floating-point weight table is no longer required
    Erase contrib
    
    'With weights successfully calculated, we can now filter vertically from the "working copy"
    ' to the destination image.
    If (yScale <> 1#) Then
        
        'Because we need to access pixels across rows, it's easiest to just wrap a single array around
        'the entire image.
        dstDIB.WrapArrayAroundDIB_1D imgPixels, srcSA
        
        'Each column (new image width)...
        For k = 0 To dstWidth - 1
            
            'Pre-calculate a fixed x-offset for this column
            xOffset = k * 4
            
            'Each row (destination image height)...
            For i = 0 To dstHeight - 1
                
                iB = 0
                iG = 0
                iR = 0
                iA = 0
                
                'Generate weighted result for each color component
                For j = 0 To contribI(i).nCount - 1
                    weightI = contribI(i).p(j).weight
                    idxPixel = contribI(i).p(j).pixel + xOffset
                    iB = iB + weightI * m_tmpPixels(idxPixel)
                    iG = iG + weightI * m_tmpPixels(idxPixel + 1)
                    iR = iR + weightI * m_tmpPixels(idxPixel + 2)
                    iA = iA + weightI * m_tmpPixels(idxPixel + 3)
                Next j
                
                'Weight and clamp final RGBA values.
                b = iB \ LARGE_POWER_OF_TWO
                g = iG \ LARGE_POWER_OF_TWO
                r = iR \ LARGE_POWER_OF_TWO
                a = iA \ LARGE_POWER_OF_TWO
                
                If (b > 255) Then b = 255
                If (g > 255) Then g = 255
                If (r > 255) Then r = 255
                If (a > 255) Then a = 255
                
                If (b < 0) Then b = 0
                If (g < 0) Then g = 0
                If (r < 0) Then r = 0
                If (a < 0) Then a = 0
                
                'Assign new RGBA values to the working data array
                idxPixel = (k * 4) + (i * dstWidth * 4)
                imgPixels(idxPixel) = b
                imgPixels(idxPixel + 1) = g
                imgPixels(idxPixel + 2) = r
                imgPixels(idxPixel + 3) = a
                
            'Next row...
            Next i
            
            'Report progress
            If (displayProgress And (((k + progX) And progBarCheck) = 0)) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k + progX
            End If
            
        'Next column...
        Next k
        
        'Release all unsafe references
        dstDIB.UnwrapArrayFromDIB imgPixels
        
    'If the image's vertical size *isn't* changing, just mirror the intermediate data into dstImage.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination earlier - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict dstDIB.GetDIBPointer, VarPtr(m_tmpPixels(0)), dstHeight * dstWidth * 4
    End If
    
    If REPORT_RESAMPLE_PERF Then
        If REPORT_DETAILED_PERF Then PDDebug.LogAction "Resampling.ResampleImage vertical 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        m_NetTimeI = m_NetTimeI + VBHacks.GetTimerDifferenceNow(firstTime)
        m_IterationsI = m_IterationsI + 1
        PDDebug.LogAction "Average resampling time for this session (I): " & VBHacks.GetTotalTimeAsString((m_NetTimeI / m_IterationsI) * 1000)
    End If
    
    'Resampling complete!
    ResampleImageI = True
    
End Function

'Resample from one floating-point surface to another.  Optimizations used in the integer-based resampling functions
' are not reused here, by design; this means this function provides full HDR color coverage.
Public Function ResampleImageF(ByRef dstSurface As pdSurfaceF, ByRef srcSurface As pdSurfaceF, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal rsFilter As PD_ResamplingFilter, Optional ByVal displayProgress As Boolean = False) As Boolean
    
    ResampleImageF = False
    Const FUNC_NAME As String = "ResampleImageF"
    If REPORT_RESAMPLE_PERF Then PDDebug.LogAction "Float resampler started."
    
    'Validate all inputs
    If (srcSurface Is Nothing) Then
        InternalError FUNC_NAME, "null source"
        Exit Function
    End If
    
    If (dstWidth <= 0) Or (dstHeight <= 0) Then
        InternalError FUNC_NAME, "bad width/height: " & dstWidth & ", " & dstHeight
        Exit Function
    End If
    
    'Initialize destination as necessary
    If (dstSurface Is Nothing) Then Set dstSurface = New pdSurfaceF
    If (dstSurface.GetWidth <> dstWidth) Or (dstSurface.GetHeight <> dstHeight) Then dstSurface.CreateBlank dstWidth, dstHeight, srcSurface.GetChannelCount
    
    'Reflect alpha channel behavior in the destination surface
    dstSurface.SetInitialAlphaPremultiplicationState srcSurface.GetAlphaPremultiplication
    
    'Performance reporting (via debug logs) is controlled by constants at the top of this class
    Dim startTime As Currency, firstTime As Currency
    VBHacks.GetHighResTime startTime
    firstTime = startTime
    
    'Validate all internal values as well
    If (m_LanczosRadius < LANCZOS_MIN) Or (m_LanczosRadius > LANCZOS_MAX) Then m_LanczosRadius = LANCZOS_DEFAULT
    
    'Inputs look good.  Prepare intermediary data structs.  Custom types are used to improve memory locality.
    Dim srcWidth As Long: srcWidth = srcSurface.GetWidth
    Dim srcHeight As Long: srcHeight = srcSurface.GetHeight
    Dim srcChannelCount As Long: srcChannelCount = srcSurface.GetChannelCount
    
    'Allocate the intermediary "working" copy of the image width dimensions [dstWidth, srcHeight].
    ' Note that unlike the other resampling functions, variable-sized channel counts are supported.
    If (dstWidth * srcHeight * srcChannelCount > m_tmpPixelSizeF) Then
        m_tmpPixelSizeF = dstWidth * srcHeight * srcChannelCount
        ReDim m_tmpPixelsF(0 To m_tmpPixelSizeF - 1) As Single
    Else
        VBHacks.ZeroMemory VarPtr(m_tmpPixelsF(0)), m_tmpPixelSizeF * 4
    End If
    
    'Calculate x/y scales; this provides a simple mechanism for checking up- vs downsampling
    ' in either direction.
    Dim xScale As Double, yScale As Double
    xScale = dstWidth / srcWidth
    yScale = dstHeight / srcHeight
    
    'If progress bar reports are wanted, calculate max values now
    Dim progX As Long, progY As Long, progBarCheck As Long
    If displayProgress Then
        If (xScale <> 1#) Then progX = srcHeight
        If (yScale <> 1#) Then progY = dstWidth
        ProgressBars.SetProgBarMax progX + progY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep array of contributors (which contain per-pixel weights)
    Dim contrib() As ContributorEntry
    ReDim contrib(0 To dstWidth - 1) As ContributorEntry
    
    Dim radius As Double, center As Double, weight As Double
    Dim intensityR As Single, intensityG As Single, intensityB As Single, intensityA As Single
    Dim pxLeft As Long, pxRight As Long, i As Long, j As Long, k As Long
    Dim r As Single, g As Single, b As Single, a As Single
    Dim xOffset As Long
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImageF prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Horizontal downsampling
    If (xScale < 1#) Then
        
        'The source width is larger than the destination width
        radius = (GetDefaultRadius(rsFilter) / xScale)
        
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table to the max possible size
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column (and ensure valid boundaries)
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                'Calculate weight for this pixel, according to the current filter
                weight = GetValue(rsFilter, (center - j - 0.5) * xScale)
                
                contrib(i).p(contrib(i).nCount).pixel = j * srcChannelCount   'Offset by channel count in advance
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    'Horizontal upsampling
    ElseIf (xScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source width is smaller than the destination width
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                contrib(i).p(contrib(i).nCount).pixel = j * srcChannelCount   'Offset by channel count in advance
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    '/End up- vs downsampling.  Note that the special case of xScale = 1.0 (e.g. horizontal width
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImageF horizontal 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter horizontally from the input image
    ' to the temporary "working" copy.
    Dim imgPixels() As Single, srcSA As SafeArray1D
    Dim idxPixel As Long, wSum As Double
    
    'If the image is changing size, perform resampling now
    If (xScale <> 1#) Then
        
        'Each row (source image height)...
        For k = 0 To srcHeight - 1
            
            'Wrap a VB array around the image at this line
            srcSurface.WrapArrayAroundScanline imgPixels, srcSA, k
            
            'Each column (destination image width)...
            For i = 0 To dstWidth - 1
                
                intensityB = 0!
                intensityG = 0!
                intensityR = 0!
                intensityA = 0!
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel
                    intensityB = intensityB + (imgPixels(idxPixel) * weight)
                    intensityG = intensityG + (imgPixels(idxPixel + 1) * weight)
                    intensityR = intensityR + (imgPixels(idxPixel + 2) * weight)
                    intensityA = intensityA + (imgPixels(idxPixel + 3) * weight)
                Next j
                
                'Weight and clamp final RGBA values.  (Note that normally you'd *divide* by the
                ' weighted sum here, but we already normalized that value in a previous step.)
                wSum = contrib(i).weightSum
                
                b = intensityB * wSum
                g = intensityG * wSum
                r = intensityR * wSum
                a = intensityA * wSum
                
                'Clamping isn't technically required on floating-point values, but note that out-of-gamut values
                ' *can* occur after resampling.
                '
                'The exception to this rule is alpha data, which cannot exist outside [0, 1]
                If (a < 0!) Then a = 0!
                If (a > 1!) Then a = 1!
                
                'Assign new RGBA values to the working data array
                idxPixel = k * dstWidth * 4 + i * 4
                m_tmpPixelsF(idxPixel) = b
                m_tmpPixelsF(idxPixel + 1) = g
                m_tmpPixelsF(idxPixel + 2) = r
                m_tmpPixelsF(idxPixel + 3) = a
                
            'Next pixel in row...
            Next i
            
            'Report progress
            If displayProgress And ((k And progBarCheck) = 0) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k
            End If
            
        'Next row in image...
        Next k
    
        'Free any unsafe references
        srcSurface.UnwrapArrayFromSurface imgPixels
    
    'If the image's horizontal size *isn't* changing, just mirror the data into the temporary array.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination in the next step - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict VarPtr(m_tmpPixelsF(0)), srcSurface.GetPixelPtr, dstWidth * srcHeight * srcSurface.GetChannelCount * 4
    End If
    
    'Horizontal sampling is now complete.
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImageF horizontal 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Next, we need to perform nearly identical sampling from the "working" copy to the destination image,
    ' while resampling in the y-direction.
    
    'Reset contributor weight table (one entry per row for vertical resampling)
    ReDim contrib(0 To dstHeight - 1) As ContributorEntry
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImageF vertical prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Vertical downsampling
    If (yScale < 1#) Then
        
        'The source height is larger than the destination height
        radius = GetDefaultRadius(rsFilter) / yScale
        
        'Iterate through each row in the image
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
          
            center = (i + 0.5) / yScale
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            'Precalculate all weights for this column (technically these are not left/right values
            ' but up/down ones, remember)
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, (center - j - 0.5) * yScale)
                
                contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    'Vertical upsampling
    ElseIf (yScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source height is smaller than the destination height
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            center = ((i + 0.5) / yScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.9999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                contrib(i).p(contrib(i).nCount).pixel = j * dstWidth * 4    'Precalculate a row offset
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    '/End up- vs downsampling.  Note that the special case of yScale = 1.0 (e.g. vertical height
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleImageF vertical 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter vertically from the "working copy"
    ' to the destination image.
    If (yScale <> 1#) Then
        
        'Because we need to access pixels across rows, it's easiest to just wrap a single array around
        'the entire image.
        dstSurface.WrapArrayAroundSurface_1D imgPixels, srcSA
        
        'Each column (new image width)...
        For k = 0 To dstWidth - 1
            
            'Pre-calculate a fixed x-offset for this column
            xOffset = k * 4
            
            'Each row (destination image height)...
            For i = 0 To dstHeight - 1
                
                intensityB = 0!
                intensityG = 0!
                intensityR = 0!
                intensityA = 0!
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel + xOffset
                    intensityB = intensityB + (m_tmpPixelsF(idxPixel) * weight)
                    intensityG = intensityG + (m_tmpPixelsF(idxPixel + 1) * weight)
                    intensityR = intensityR + (m_tmpPixelsF(idxPixel + 2) * weight)
                    intensityA = intensityA + (m_tmpPixelsF(idxPixel + 3) * weight)
                Next j
                
                'Weight and clamp final RGBA values
                wSum = contrib(i).weightSum
                
                b = intensityB * wSum
                g = intensityG * wSum
                r = intensityR * wSum
                a = intensityA * wSum
                
                'Clamping isn't technically required on floating-point values, but note that out-of-gamut values
                ' *can* occur after resampling.
                
                'Assign new RGBA values to the working data array
                idxPixel = (k * 4) + (i * dstWidth * 4)
                imgPixels(idxPixel) = b
                imgPixels(idxPixel + 1) = g
                imgPixels(idxPixel + 2) = r
                imgPixels(idxPixel + 3) = a
                
            'Next row...
            Next i
            
            'Report progress
            If (displayProgress And (((k + progX) And progBarCheck) = 0)) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k + progX
            End If
            
        'Next column...
        Next k
        
        'Release all unsafe references
        dstSurface.UnwrapArrayFromSurface imgPixels
        
    'If the image's vertical size *isn't* changing, just mirror the intermediate data into dstImage.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination earlier - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict dstSurface.GetPixelPtr, VarPtr(m_tmpPixelsF(0)), dstWidth * dstHeight * dstSurface.GetChannelCount * 4
    End If
    
    'Safely clamp [0, 1] values to ensure proper alpha handling.
    dstSurface.ClampGamut
    
    If REPORT_RESAMPLE_PERF Then
        If REPORT_DETAILED_PERF Then PDDebug.LogAction "Resampling.ResampleImageF vertical 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        m_NetTimeF = m_NetTimeF + VBHacks.GetTimerDifferenceNow(firstTime)
        m_IterationsF = m_IterationsF + 1
        PDDebug.LogAction "Average resampling time for this session (F): " & VBHacks.GetTotalTimeAsString((m_NetTimeF / m_IterationsF) * 1000)
    End If
    
    'Resampling complete!
    ResampleImageF = True
    
End Function

'Resample from one int (long) array to another.  Note that this function has not been aggressively tested for
' OOB issues; VB6's lack of a LongLong type makes it very difficult to handle this case efficiently.  PD currently
' uses this function for handling 16-bit data (a lack of UShort makes it easier to handle via long arrays)
' so range issues aren't a concern, but it could be for very large and/or arbitrary Long-type data.
Public Function ResampleArrayL(ByRef dstArray() As Long, ByRef srcArray() As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rsFilter As PD_ResamplingFilter, Optional ByVal displayProgress As Boolean = False) As Boolean
    
    ResampleArrayL = False
    Const FUNC_NAME As String = "ResampleArrayL"
    If REPORT_RESAMPLE_PERF Then PDDebug.LogAction "Long array resampler started."
    
    'Validate all inputs
    If Not VBHacks.IsArrayInitialized(srcArray) Or (srcWidth <= 0) Or (srcHeight <= 0) Then
        InternalError FUNC_NAME, "null source"
        Exit Function
    End If
    
    If (dstWidth <= 0) Or (dstHeight <= 0) Then
        InternalError FUNC_NAME, "bad width/height: " & dstWidth & ", " & dstHeight
        Exit Function
    End If
    
    'Initialize destination as necessary
    If (Not VBHacks.IsArrayInitialized(dstArray)) Then ReDim dstArray(0 To dstWidth - 1, 0 To dstHeight - 1) As Long
    
    'Performance reporting (via debug logs) is controlled by constants at the top of this class
    Dim startTime As Currency, firstTime As Currency
    VBHacks.GetHighResTime startTime
    firstTime = startTime
    
    'Validate all internal values as well
    If (m_LanczosRadius < LANCZOS_MIN) Or (m_LanczosRadius > LANCZOS_MAX) Then m_LanczosRadius = LANCZOS_DEFAULT
    
    'Inputs look good.  Prepare intermediary data structs.  Custom types are used to improve memory locality.
    
    'Allocate the intermediary "working" copy of the image width dimensions [dstWidth, srcHeight].
    ' Note that unlike the other resampling functions, variable-sized channel counts are supported.
    If (dstWidth * srcHeight > m_tmpPixelSizeL) Then
        m_tmpPixelSizeL = dstWidth * srcHeight
        ReDim m_tmpPixelsL(0 To m_tmpPixelSizeL - 1) As Long
    Else
        VBHacks.ZeroMemory VarPtr(m_tmpPixelsL(0)), m_tmpPixelSizeL * 4
    End If
    
    'Calculate x/y scales; this provides a simple mechanism for checking up- vs downsampling
    ' in either direction.
    Dim xScale As Double, yScale As Double
    xScale = dstWidth / srcWidth
    yScale = dstHeight / srcHeight
    
    'If progress bar reports are wanted, calculate max values now
    Dim progX As Long, progY As Long, progBarCheck As Long
    If displayProgress Then
        If (xScale <> 1#) Then progX = srcHeight
        If (yScale <> 1#) Then progY = dstWidth
        ProgressBars.SetProgBarMax progX + progY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep array of contributors (which contain per-pixel weights)
    Dim contrib() As ContributorEntry
    ReDim contrib(0 To dstWidth - 1) As ContributorEntry
    
    Dim radius As Double, center As Double, weight As Double
    Dim intensity As Long
    Dim pxLeft As Long, pxRight As Long, i As Long, j As Long, k As Long
    Dim r As Single
    Dim xOffset As Long
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleArrayL prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Horizontal downsampling
    If (xScale < 1#) Then
        
        'The source width is larger than the destination width
        radius = (GetDefaultRadius(rsFilter) / xScale)
        
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table to the max possible size
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column (and ensure valid boundaries)
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                'Calculate weight for this pixel, according to the current filter
                weight = GetValue(rsFilter, (center - j - 0.5) * xScale)
                
                contrib(i).p(contrib(i).nCount).pixel = j
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    'Horizontal upsampling
    ElseIf (xScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source width is smaller than the destination width
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.99999999999)
            If (pxRight >= srcWidth) Then pxRight = srcWidth - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                contrib(i).p(contrib(i).nCount).pixel = j
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        Next i
    
    '/End up- vs downsampling.  Note that the special case of xScale = 1.0 (e.g. horizontal width
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleArrayL horizontal 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter horizontally from the input image
    ' to the temporary "working" copy.
    Dim idxPixel As Long, wSum As Double
    
    'If the image is changing size, perform resampling now
    If (xScale <> 1#) Then
        
        'Each row (source image height)...
        For k = 0 To srcHeight - 1
            
            'Each column (destination image width)...
            For i = 0 To dstWidth - 1
                
                intensity = 0
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel
                    intensity = intensity + (srcArray(idxPixel, k) * weight)
                Next j
                
                'Weight and clamp final RGBA values.  (Note that normally you'd *divide* by the
                ' weighted sum here, but we already normalized that value in a previous step.)
                wSum = contrib(i).weightSum
                r = intensity * wSum
                
                'Clamping isn't technically required on HDR values, but note that out-of-gamut values
                ' *can* occur after resampling.
                
                'Assign new RGBA values to the working data array
                idxPixel = k * dstWidth + i
                m_tmpPixelsL(idxPixel) = r
                
            'Next pixel in row...
            Next i
            
            'Report progress
            If displayProgress And ((k And progBarCheck) = 0) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k
            End If
            
        'Next row in image...
        Next k
    
    'If the image's horizontal size *isn't* changing, just mirror the data into the temporary array.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination in the next step - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict VarPtr(m_tmpPixelsL(0)), VarPtr(srcArray(0, 0)), dstWidth * srcHeight * 4
    End If
    
    'Horizontal sampling is now complete.
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleArrayL horizontal 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Next, we need to perform nearly identical sampling from the "working" copy to the destination image,
    ' while resampling in the y-direction.
    
    'Reset contributor weight table (one entry per row for vertical resampling)
    ReDim contrib(0 To dstHeight - 1) As ContributorEntry
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleArrayL vertical prep time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Vertical downsampling
    If (yScale < 1#) Then
        
        'The source height is larger than the destination height
        radius = GetDefaultRadius(rsFilter) / yScale
        
        'Iterate through each row in the image
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
          
            center = (i + 0.5) / yScale
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            'Precalculate all weights for this column (technically these are not left/right values
            ' but up/down ones, remember)
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, (center - j - 0.5) * yScale)
                
                contrib(i).p(contrib(i).nCount).pixel = j * dstWidth
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    'Vertical upsampling
    ElseIf (yScale > 1#) Then
        
        radius = GetDefaultRadius(rsFilter)
        
        'The source height is smaller than the destination height
        For i = 0 To dstHeight - 1
            
            contrib(i).nCount = 0
            ReDim contrib(i).p(0 To Int(2 * radius + 2)) As Contributor
            contrib(i).weightSum = 0#
            
            center = ((i + 0.5) / yScale)
            pxLeft = Int(center - radius)
            If (pxLeft < 0) Then pxLeft = 0
            pxRight = Int(center + radius + 0.9999999999)
            If (pxRight >= srcHeight) Then pxRight = srcHeight - 1
            
            For j = pxLeft To pxRight
                
                weight = GetValue(rsFilter, center - j - 0.5)
                
                contrib(i).p(contrib(i).nCount).pixel = j * dstWidth
                contrib(i).p(contrib(i).nCount).weight = weight
                contrib(i).weightSum = contrib(i).weightSum + weight
                contrib(i).nCount = contrib(i).nCount + 1
                
            Next j
            
            'Normalize the weight sum before exiting
            If (contrib(i).weightSum <> 0#) Then contrib(i).weightSum = 1# / contrib(i).weightSum
            
        'Next row...
        Next i
    
    '/End up- vs downsampling.  Note that the special case of yScale = 1.0 (e.g. vertical height
    ' isn't changing) is not handled here; we'll handle it momentarily as a special case.
    End If
    
    If REPORT_DETAILED_PERF Then
        PDDebug.LogAction "Resampling.ResampleArrayL vertical 1 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With weights successfully calculated, we can now filter vertically from the "working copy"
    ' to the destination image.
    If (yScale <> 1#) Then
        
        'Each column (new image width)...
        For k = 0 To dstWidth - 1
            
            'Pre-calculate a fixed x-offset for this column
            xOffset = k
            
            'Each row (destination image height)...
            For i = 0 To dstHeight - 1
                
                intensity = 0
                
                'Generate weighted result for each color component
                For j = 0 To contrib(i).nCount - 1
                    weight = contrib(i).p(j).weight
                    idxPixel = contrib(i).p(j).pixel + xOffset
                    intensity = intensity + (m_tmpPixelsL(idxPixel) * weight)
                Next j
                
                'Weight and clamp final RGBA values
                wSum = contrib(i).weightSum
                r = intensity * wSum
                
                'Clamping isn't technically required on floating-point values, but note that out-of-gamut values
                ' *can* occur after resampling.
                
                'Assign new RGBA values to the working data array
                idxPixel = k + (i * dstWidth)
                dstArray(idxPixel) = r
                
            'Next row...
            Next i
            
            'Report progress
            If (displayProgress And (((k + progX) And progBarCheck) = 0)) Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal k + progX
            End If
            
        'Next column...
        Next k
        
    'If the image's vertical size *isn't* changing, just mirror the intermediate data into dstImage.
    ' (NOTE: with some trickery, we could also skip this step entirely and simply copy data from source to
    '  destination earlier - but VB makes this harder than it needs to be so I haven't attempted
    '  it yet...)
    Else
        CopyMemoryStrict VarPtr(dstArray(0, 0)), VarPtr(m_tmpPixelsL(0)), dstWidth * dstHeight * 4
    End If
    
    If REPORT_RESAMPLE_PERF Then
        If REPORT_DETAILED_PERF Then PDDebug.LogAction "Resampling.ResampleArrayL vertical 2 time: " & VBHacks.GetTimeDiffNowAsString(startTime)
        m_NetTimeF = m_NetTimeF + VBHacks.GetTimerDifferenceNow(firstTime)
        m_IterationsF = m_IterationsF + 1
        PDDebug.LogAction "Average resampling time for this session (F): " & VBHacks.GetTotalTimeAsString((m_NetTimeF / m_IterationsF) * 1000)
    End If
    
    'Resampling complete!
    ResampleArrayL = True
    
End Function

Private Function GetDefaultRadius(ByVal rsType As PD_ResamplingFilter) As Double

    Select Case rsType
        
        Case rf_Automatic
            'dummy entry
            
        Case rf_Box
            GetDefaultRadius = 0.5
        
        Case rf_BilinearTriangle
            GetDefaultRadius = 1#
        Case rf_Cosine
            GetDefaultRadius = 1#
        Case rf_Hermite
            GetDefaultRadius = 1#
        
        Case rf_Bell
            GetDefaultRadius = 1.5
        Case rf_Quadratic
            GetDefaultRadius = 1.5
        Case rf_QuadraticBSpline
            GetDefaultRadius = 1.5
        
        Case rf_CubicBSpline
            GetDefaultRadius = 2#
        Case rf_CatmullRom
            GetDefaultRadius = 2#
        Case rf_Mitchell
            GetDefaultRadius = 2#
        
        Case rf_CubicConvolution
            GetDefaultRadius = 3#
        
        'Lanczos now supports variable radii
        Case rf_Lanczos
            If (m_LanczosRadius < LANCZOS_MIN) Or (m_LanczosRadius > LANCZOS_MAX) Then
                GetDefaultRadius = LANCZOS_DEFAULT
            Else
                GetDefaultRadius = m_LanczosRadius
            End If

    End Select

End Function

Private Function GetValue(ByVal rsType As PD_ResamplingFilter, ByVal x As Double) As Double
    
    Dim temp As Double
    
    Select Case rsType
        
        Case rf_Automatic
            'dummy entry
        
        Case rf_Box
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then GetValue = 1# Else GetValue = 0#
      
        Case rf_BilinearTriangle
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                GetValue = (1# - x)
            Else
                GetValue = 0#
            End If
        
        Case rf_Cosine
            If ((x >= -1) And (x <= 1)) Then
                GetValue = (Cos(x * PI) + 1#) * 0.5
            Else
                GetValue = 0#
            End If
            
        Case rf_Hermite
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                GetValue = ((2# * x - 3#) * x * x + 1#)
            Else
                GetValue = 0#
            End If
        
        Case rf_Bell
            If (x < 0#) Then x = -x
            If (x < 0.5) Then
                GetValue = (0.75 - x * x)
            ElseIf (x < 1.5) Then
                temp = x - 1.5
                GetValue = (0.5 * temp * temp)
            Else
                GetValue = 0#
            End If
      
        Case rf_Quadratic
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then
                GetValue = (-2# * x * x + 1#)
            ElseIf (x <= 1.5) Then
                GetValue = (x * x - 2.5 * x + 1.5)
            Else
                GetValue = 0#
            End If
            
        Case rf_QuadraticBSpline
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then
                GetValue = (-x * x + 0.75)
            ElseIf (x <= 1.5) Then
                GetValue = 0.5 * x * x - 1.5 * x + 1.125
            Else
                GetValue = 0#
            End If
            
        Case rf_CubicBSpline
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                temp = x * x
                GetValue = (0.5 * temp * x - temp + 0.666666666666667)
            ElseIf (x < 2#) Then
                x = 2# - x
                GetValue = x * x * x * 0.166666666666667
            Else
                GetValue = 0#
            End If
            
        Case rf_CatmullRom
            If (x < 0#) Then x = -x
            temp = x * x
            If (x <= 1#) Then
                GetValue = (1.5 * temp * x - 2.5 * temp + 1#)
            ElseIf (x <= 2#) Then
                GetValue = (-0.5 * temp * x + 2.5 * temp - 4 * x + 2#)
            Else
                GetValue = 0#
            End If
            
        Case rf_Mitchell
            Const MC As Double = 0.333333333333333
            
            If (x < 0#) Then x = -x
            temp = x * x
            
            If (x < 1#) Then
                x = (((12 - 9 * MC - 6 * MC) * (x * temp)) + ((-18 + 12 * MC + 6 * MC) * temp) + (6 - 2 * MC))
                GetValue = (x / 6)
            ElseIf (x < 2#) Then
                x = (((-MC - 6 * MC) * (x * temp)) + ((6 * MC + 30 * MC) * temp) + ((-12 * MC - 48 * MC) * x) + (8 * MC + 24 * MC))
                GetValue = x * 0.166666666666667
            Else
                GetValue = 0#
            End If
            
        Case rf_CubicConvolution
            If (x < 0#) Then x = -x
            temp = x * x
            If (x <= 1#) Then
                GetValue = ((4# / 3#) * temp * x - (7# / 3#) * temp + 1#)
            ElseIf (x <= 2#) Then
                GetValue = (-(7# / 12#) * temp * x + 3# * temp - (59# / 12#) * x + 2.5)
            ElseIf (x <= 3#) Then
                GetValue = ((1# / 12#) * temp * x - (0.666666666666667) * temp + 1.75 * x - 1.5)
            Else
                GetValue = 0#
            End If
            
        Case rf_Lanczos
            If (x < 0#) Then x = -x
            If (x < m_LanczosRadius) Then
                GetValue = (SinC(x) * SinC(x / m_LanczosRadius))
            Else
                GetValue = 0#
            End If
            
    End Select

End Function

Private Function SinC(ByVal x As Double) As Double
    If (x <> 0#) Then
        x = x * PI
        SinC = Sin(x) / x
    Else
        SinC = 1#
    End If
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "Resampling." & funcName & "() reported an error: " & errDescription
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub
