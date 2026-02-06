VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPixelIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Pixel Iterator class
'Copyright 2015-2026 by Tanner Helland
'Created: 04/December/14
'Last updated: 05/April/20
'Last update: code clean-up and stability improvements
'
'By the time the 7.0 release rolled around, a ridiculous number of PD effects managed
' their own pixel iteration methods.  Most of these were variants of a "sliding window"
' implementation, where a running histogram is maintained for a given pixel region,
' and when moving to the next pixel, instead of recalculating the entire region from scratch,
' histograms from the previous pixel are simply updated against the new region, typically by
' subtracting a line of pixels on the left or top, and adding a new line of pixels on the
' right or bottom.
'
'These methods tended to use a single copy+paste chunk of code that was pretty advanced,
' supporting serpentine scanning and a number of other neat features, but maintenance quickly
' became unwieldy because any improvements to the method required copy+pasting the changes
' across dozens of functions.  Worse still, the function only supported rectangular regions,
' and many region functions look more natural when a circle is used.
'
'So as part of the 7.0 release, this class was created.  This class is designed to abstract
' away the messy duties of iterating per-pixel regions, while supporting more features than
' PD's old implementation.  Because this class uses generic histograms, many functions can
' tap into it, without needing to modify the actual iteration code.
'
'Please note that this class has several dependencies throughout PD, including pdDIB,
' and some specialized enums (located in PD's central Public_Enums_and_Types module).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current (x, y) position of the iterator.  Call InitializeIterator() to reset these to (0, 0).
Private m_X As Long, m_Y As Long

'Initial and final loop boundaries
Private m_InitX As Long, m_InitY As Long
Private m_FinalX As Long, m_FinalY As Long

'Shape of the current iterator, including boundary distance
Private m_WindowShape As PD_PixelRegionShape
Private m_XBLeft As Long, m_XBRight As Long, m_XBSize As Long
Private m_YBTop As Long, m_YBBottom As Long, m_YBSize As Long

'Other relevant DIB properties
Private m_DibPointer As Long, m_dibHeight As Long, m_dibStride As Long

'Some functions may choose to use a byte array target instead of a DIB.  That's fine!
Private m_BytesSA2D As SafeArray2D, m_Bytes() As Byte
Private m_BytesSA1D As SafeArray1D, m_ByteLine() As Byte
Private m_ByteArrayPtr As Long, m_ArrayWidth As Long, m_ArrayHeight As Long

'Pixel array (alias only; the actual pixel bits are stored by the caller, and treated as read-only by this class).
' In some situations it is advantageous to address the pixels using a 1D array, as VB is significantly slower when
' addressing 2D pixel positions.
Private m_PixelSA2D As SafeArray2D, m_Pixels() As Byte
Private m_PixelSA1D As SafeArray1D, m_PixelLine() As Byte

'If the function wants alpha tracked, this will be set to TRUE.
' (NOTE: this tracker has been manually disabled, because no PD functions use it as present.  Skipping it on inner
'        loops provides a small performance boost - so if you *do* need alpha iterated in the future, you will need
'        to uncomment various lines where this is implemented.)
Private m_AlphaSupport As Boolean

'This class currently supports two modes of operation: RGBA, and Luminance.  While we could track both simultaneously,
' there is a performance penalty to this, so it's easier to simply track one or the other.
' (NOTE: this value is set by the respective LockTargetHistograms function)
Private m_HistogramMode As PD_PixelIteratorMode
Private m_LuminanceMode As PD_LuminanceMode

'Histogram arrays.  Managing these are a bit tricky, because these are simply "trick" wrappers against arrays
' provided by the caller.
Private m_Red() As Long, m_Green() As Long, m_Blue() As Long, m_Alpha() As Long

'Note that the caller can also use the pixel iterator in luminance mode.  This relies on only a single destination
' luminance array, and dedicated movement/calculation functions are used.
Private m_Luminance() As Long

'Number of pixels in the histogram.  This is required for things like median calculations.
Private m_NumOfPixels As Long

'Non-standard shapes require an intermediate image; we use this as our guide for how to process kernel edges.
Private m_ShapeImage As pdDIB, m_ReferenceMap() As Byte

'Non-standard shapes require us to generate custom boundary indices.  These tell us where the kernel region ends
' in each of the forward/backward x/y directions.
Private m_XLeft() As Long, m_XRight() As Long, m_YTop() As Long, m_YBottom() As Long

'This function is the first one you need to call.  It will initialize a bunch of internal bits against the target DIB,
' bits that are required prior to actually iterating through individual pixels.
'
'Returns: TRUE if successful.  DO NOT PROCEED with pixel iterating if the function returns FALSE.
Friend Function InitializeIterator(ByRef targetDIB As pdDIB, ByVal xRadius As Long, ByVal yRadius As Long, Optional ByVal windowShape As PD_PixelRegionShape = PDPRS_Rectangle) As Boolean

    'Reset the iterator coordinates and pixel count
    m_X = 0
    m_Y = 0
    m_NumOfPixels = 0
    
    'Cache loop boundaries
    m_InitX = 0
    m_InitY = 0
    m_FinalX = targetDIB.GetDIBWidth - 1
    m_FinalY = targetDIB.GetDIBHeight - 1
    
    'Cache shape and assumed bounds.  (In the future, these bounds may change according to automated trim functions,
    ' but for now, they are hard-coded per the user's request.)
    m_WindowShape = windowShape
    
    If (xRadius < 1) Then xRadius = 1
    If (yRadius < 1) Then yRadius = 1
    m_XBLeft = xRadius
    m_XBRight = xRadius
    m_YBTop = yRadius
    m_YBBottom = yRadius
    
    'At certain sizes, some window shapes are unreliable.  Force these to use the failsafe rectangular engine.
    If (m_WindowShape = PDPRS_Circle) And ((xRadius <= 1) Or (yRadius <= 1)) Then m_WindowShape = PDPRS_Rectangle
    
    'Apply some failsafe dimension testing to the incoming bounds
    If (m_XBRight > (m_FinalX - m_InitX)) Then
        m_XBRight = (m_FinalX - m_InitX)
        m_XBLeft = m_XBRight
    End If
    
    If (m_YBBottom > (m_FinalY - m_InitY)) Then
        m_YBBottom = (m_FinalY - m_InitY)
        m_YBTop = m_YBBottom
    End If
    
    'Store the final kernel size
    m_XBSize = m_XBLeft + m_XBRight + 1
    m_YBSize = m_YBTop + m_YBBottom + 1
    
    'Retrieve other relevant DIB properties
    m_DibPointer = targetDIB.GetDIBPointer
    m_dibHeight = targetDIB.GetDIBHeight
    m_dibStride = targetDIB.GetDIBStride
    
    'Set alpha to a default value, based on the source image's color depth
    'm_AlphaSupport = True
    
    InitializeIterator = True
    
End Function

'Same idea as the previous function, except that a target byte array is used instead.  (This can yield significantly
' better performance, but obviously it requires you to only need one histogram of data.)
'
'Returns: TRUE if successful.  DO NOT PROCEED with pixel iterating if the function returns FALSE.
Friend Function InitializeIterator_ByteArray(ByRef srcByteArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal xRadius As Long, ByVal yRadius As Long, Optional ByVal windowShape As PD_PixelRegionShape = PDPRS_Rectangle) As Boolean

    'Reset the iterator coordinates and pixel count
    m_X = 0
    m_Y = 0
    m_NumOfPixels = 0
    
    'Cache loop boundaries
    m_InitX = 0
    m_InitY = 0
    m_FinalX = arrayWidth - 1
    m_FinalY = arrayHeight - 1
    
    'Cache shape and assumed bounds.  (In the future, these bounds may change according to automated trim functions,
    ' but for now, they are hard-coded per the user's request.)
    m_WindowShape = windowShape
    
    If (xRadius < 1) Then xRadius = 1
    If (yRadius < 1) Then yRadius = 1
    m_XBLeft = xRadius
    m_XBRight = xRadius
    m_YBTop = yRadius
    m_YBBottom = yRadius
    
    'At certain sizes, some window shapes are unreliable.  Force these to use the failsafe rectangular engine.
    If (m_WindowShape = PDPRS_Circle) And ((xRadius <= 1) Or (yRadius <= 1)) Then m_WindowShape = PDPRS_Rectangle
    
    'Apply some failsafe dimension testing to the incoming bounds
    If (m_XBRight > (m_FinalX - m_InitX)) Then
        m_XBRight = (m_FinalX - m_InitX)
        m_XBLeft = m_XBRight
    End If
    
    If (m_YBBottom > (m_FinalY - m_InitY)) Then
        m_YBBottom = (m_FinalY - m_InitY)
        m_YBTop = m_YBBottom
    End If
    
    'Store the final kernel size
    m_XBSize = m_XBLeft + m_XBRight + 1
    m_YBSize = m_YBTop + m_YBBottom + 1
    
    'Cache misc pointers and other items
    m_ByteArrayPtr = VarPtr(srcByteArray(0, 0))
    m_ArrayWidth = arrayWidth
    m_ArrayHeight = arrayHeight
    
    InitializeIterator_ByteArray = True
    
End Function

'After you've initialized the iterator, call this function to setup the initial pixel region.  The caller must supply
' their own histogram arrays; we will wrap these with some "trick" internal array references, to avoid the need for
' passing these on every pixel request.
'
'IMPORTANT NOTE: PRIOR TO CALLING THIS FUNCTION, you must redim these arrays to range [0, 255].  Do not ReDim them until
' you have completed your function and freed the histograms safely (via ReleaseTargetHistograms, below).
'
'OTHER IMPORTANT NOTE: at present, PD does *not* make use of the dstAlpha parameter anywhere.  To improve performance,
' I have manually disabled alpha channel support in the inner loops.  If alpha iterating becomes important in the future,
' you will need to manually uncomment those lines of code.
'
'This function will return the pixel count of the first window in the image.  DO NOT PROCEED if it returns zero.
Friend Function LockTargetHistograms_RGBA(ByRef dstRed() As Long, ByRef dstGreen() As Long, ByRef dstBlue() As Long, ByRef dstAlpha() As Long, Optional ByVal calcAlpha As Boolean = True) As Long
    
    'Put the iterator in RGB mode
    m_HistogramMode = PDPIM_RGBA
    
    'Alias our internal histogram arrays around the destination ones.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target arrays until the iterator has finished.
    VBHacks.AliasArbitraryArray VarPtrArray(dstRed), VarPtrArray(m_Red)
    VBHacks.AliasArbitraryArray VarPtrArray(dstGreen), VarPtrArray(m_Green)
    VBHacks.AliasArbitraryArray VarPtrArray(dstBlue), VarPtrArray(m_Blue)
    
    'NOTE: see comments elsewhere on m_AlphaSupport.  It is currently disabled; you'll need to uncomment various bits
    ' to make it work.
    calcAlpha = False
    If calcAlpha And m_AlphaSupport Then
        VBHacks.AliasArbitraryArray VarPtrArray(dstAlpha), VarPtrArray(m_Alpha)
    Else
        m_AlphaSupport = False
    End If
    
    'Point our internal 2D pixel array at the target DIB and generate the initial histogram window
    AliasPixelArray
    LockTargetHistograms_RGBA = GenerateInitialWindow()
    
End Function

'When the iterator is finished (due to any condition - success, error, etc), the caller MUST call this function to
' release our aliases to their histogram arrays and DIB.
Friend Function ReleaseTargetHistograms_RGBA(ByRef dstRed() As Long, ByRef dstGreen() As Long, ByRef dstBlue() As Long, ByRef dstAlpha() As Long) As Boolean

    VBHacks.UnaliasArbitraryArray VarPtrArray(dstRed), VarPtrArray(m_Red)
    VBHacks.UnaliasArbitraryArray VarPtrArray(dstGreen), VarPtrArray(m_Green)
    VBHacks.UnaliasArbitraryArray VarPtrArray(dstBlue), VarPtrArray(m_Blue)
    If m_AlphaSupport Then VBHacks.UnaliasArbitraryArray VarPtrArray(dstAlpha), VarPtrArray(m_Alpha)
    
    'While we're here, release our DIB reference, too
    UnaliasPixelArray
    
    ReleaseTargetHistograms_RGBA = True
    
End Function

'Luminance counterpart to the RGBA lock function above.  All the same caveats and rules apply.
'
'IMPORTANT NOTE: PRIOR TO CALLING THIS FUNCTION, you must redim the target array to range [0, 255].  Do not ReDim
' it until you have completed your function and freed the histogram safely (via ReleaseTargetHistograms, below).
'
'This function will return the pixel count of the first window in the image.  DO NOT PROCEED if it returns zero.
Friend Function LockTargetHistograms_Luminance(ByRef dstLum() As Long, Optional ByVal lumMode As PD_LuminanceMode = PDLM_Value) As Long
    
    'Put the iterator in Luminance mode
    m_HistogramMode = PDPIM_Luminance
    m_LuminanceMode = lumMode
    
    'Alias our internal histogram arrays around the destination ones.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target arrays until the iterator has finished.
    VBHacks.AliasArbitraryArray VarPtrArray(dstLum), VarPtrArray(m_Luminance)
    
    'Point our internal 2D pixel array at the target DIB and generate the initial histogram window
    AliasPixelArray
    LockTargetHistograms_Luminance = GenerateInitialWindow()
    
End Function

'When the iterator is finished (due to any condition - success, error, etc), the caller MUST call this function to
' release our aliases to their histogram arrays and DIB.
Friend Function ReleaseTargetHistograms_Luminance(ByRef dstLum() As Long) As Boolean

    VBHacks.UnaliasArbitraryArray VarPtrArray(dstLum), VarPtrArray(m_Luminance)
    
    'While we're here, release our DIB reference, too
    UnaliasPixelArray
    ReleaseTargetHistograms_Luminance = True
    
End Function

'Byte array counterpart to the lock functions above.  All the same caveats and rules apply.
'
'IMPORTANT NOTE: PRIOR TO CALLING THIS FUNCTION, you must redim the target array to range [0, 255].  Do not ReDim
' it until you have completed your function and freed the histogram safely (via ReleaseTargetHistograms, below).
'
'This function will return the pixel count of the first window in the image.  DO NOT PROCEED if it returns zero.
Friend Function LockTargetHistograms_ByteArray(ByRef dstHist() As Long) As Long
    
    'Put the iterator in byte array mode
    m_HistogramMode = PDPIM_ByteArray
    
    'Alias our internal histogram array around the destination ones.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target array until the iterator has finished.
    VBHacks.AliasArbitraryArray VarPtrArray(dstHist), VarPtrArray(m_Luminance)
    
    'Point our internal 2D pixel array at the target DIB and generate the initial histogram window
    AliasByteArray
    LockTargetHistograms_ByteArray = GenerateInitialWindow_ByteArray()
    
End Function

'When the iterator is finished (due to any condition - success, error, etc), the caller MUST call this function to
' release our aliases to their histogram arrays and DIB.
Friend Function ReleaseTargetHistograms_ByteArray(ByRef dstHist() As Long) As Boolean

    VBHacks.UnaliasArbitraryArray VarPtrArray(dstHist), VarPtrArray(m_Luminance)
    
    'While we're here, release our DIB reference, too
    UnaliasByteArray
    ReleaseTargetHistograms_ByteArray = True
    
End Function

'Alias/unalias the target pixel array, specifically.
Private Sub AliasPixelArray()
    With m_PixelSA2D
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_dibStride
        .pvData = m_DibPointer
    End With
    PutMem4 VarPtrArray(m_Pixels()), VarPtr(m_PixelSA2D)
End Sub

Private Sub UnaliasPixelArray()
    PutMem4 VarPtrArray(m_Pixels()), 0&
End Sub

Private Sub AliasByteArray()
    With m_BytesSA2D
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_ArrayHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_ArrayWidth
        .pvData = m_ByteArrayPtr
    End With
    PutMem4 VarPtrArray(m_Bytes()), VarPtr(m_BytesSA2D)
End Sub

Private Sub UnaliasByteArray()
    PutMem4 VarPtrArray(m_Bytes()), 0&
End Sub

'Generating the initial histogram window has variable complexity, depending on the shape being used.  This function
' is universal regardless of histogram type or window shape.
Private Function GenerateInitialWindow() As Long
    
    Select Case m_WindowShape
        
        Case PDPRS_Rectangle
            GenerateInitialWindow = GenerateInitialWindow_Square()
        
        Case PDPRS_Circle
            'Generate an initial shape
            CreateShape PDPRS_Circle
            
            'Generate boundary indices
            FindBoundaries
            
            'Actually generate the initial histogram window
            GenerateInitialWindow = GenerateInitialWindow_ArbitraryShape()
            
        Case PDPRS_Diamond
            CreateShape PDPRS_Diamond
            FindBoundaries
            GenerateInitialWindow = GenerateInitialWindow_ArbitraryShape()
        
    End Select
    
End Function

'Generating the initial histogram window has variable complexity, depending on the shape being used.  This function
' is universal regardless of histogram type or window shape.
Private Function GenerateInitialWindow_ByteArray() As Long
    
    Select Case m_WindowShape
        
        Case PDPRS_Rectangle
            GenerateInitialWindow_ByteArray = GenerateInitialWindowBytes_Square()
        
        Case PDPRS_Circle
            'Generate an initial shape
            CreateShape PDPRS_Circle
            
            'Generate boundary indices
            FindBoundaries
            
            'Actually generate the initial histogram window
            GenerateInitialWindow_ByteArray = GenerateInitialWindowBytes_ArbitraryShape()
        
        Case PDPRS_Diamond
            CreateShape PDPRS_Diamond
            FindBoundaries
            GenerateInitialWindow_ByteArray = GenerateInitialWindowBytes_ArbitraryShape()
        
    End Select
    
End Function

'Create the reference map for a non-rectangular kernel
Private Sub CreateShape(ByVal srcShape As PD_PixelRegionShape)
    
    If (m_ShapeImage Is Nothing) Then Set m_ShapeImage = New pdDIB
    m_ShapeImage.CreateBlank m_XBSize, m_YBSize, 32, 0&, 255
    
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_ShapeImage
    
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor RGB(255, 255, 255)
    
    Select Case srcShape
        
        'At small sizes, the GDI+ circle rendering algorithm doesn't work very well.  To bypass this, we activate
        ' antialiasing when the search radius is only 1 px wide (for a net search diameter of 3).
        Case PDPRS_Circle
            
            Dim useAA As Boolean
            useAA = (m_XBSize <= 3) Or (m_YBSize <= 3)
            cSurface.SetSurfaceAntialiasing useAA
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            PD2D.FillEllipseF_AbsoluteCoords cSurface, cBrush, 0, 0, m_XBSize, m_YBSize
            
        Case PDPRS_Diamond
            
            cSurface.SetSurfaceAntialiasing P2_AA_None
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            Dim dstPoints() As PointFloat
            ReDim dstPoints(0 To 3) As PointFloat
            dstPoints(0).x = 0
            dstPoints(0).y = m_YBSize / 2
            dstPoints(1).x = m_XBSize / 2
            dstPoints(1).y = 0
            dstPoints(2).x = m_XBSize
            dstPoints(2).y = m_YBSize / 2
            dstPoints(3).x = m_XBSize / 2
            dstPoints(3).y = m_YBSize
            PD2D.FillPolygonF_FromPtF cSurface, cBrush, 4, VarPtr(dstPoints(0)), False
        
    End Select
    
    Set cSurface = Nothing: Set cBrush = Nothing
    
    'TODO: as a failsafe, scan the DIB and trim any empty lines, because they'll grossly mess
    ' up our calculations. (Weird shapes may not extend all the way to the edge of the target DIB,
    ' and it's difficult to predict this in advance.)
    
    'Convert the reference image into a 2D byte array, which is much faster to access and process.
    DIBs.GetDIBGrayscaleMap m_ShapeImage, m_ReferenceMap
    m_ShapeImage.EraseDIB
    
End Sub

'Using the reference map created by CreateShape, generate all four custom boundary arrays.
Private Sub FindBoundaries()

    'The reference map was created from a simple white-on-black DIB.  Black pixels (0) are ignored.  White pixels
    ' (non-zero) are included.  As such, this step just involves scanning each edge of the reference map, and storing
    ' the matching boundary coordinate at each point of each edge.  This tells us where to add and subtract histogram
    ' values when moving between rows and/or columns.
    
    'Each direction is handled in turn.
    Dim x As Long, y As Long
    
    ReDim m_XLeft(0 To m_YBSize) As Long
    ReDim m_XRight(0 To m_YBSize) As Long
    
    'For the x-directions, process each row in turn
    For y = 0 To m_YBSize - 1
        
        x = 0
        Do
        
            If (m_ReferenceMap(x, y) > 0) Then Exit Do
            
            'Move to the next column
            x = x + 1
            
            'As a failsafe for blank lines, set a dummy value if the entire line is blank
            If (x > m_XBSize - 1) Then
                x = m_XBLeft
                Exit Do
            End If
            
        Loop
        
        'x now points at the position on this line where the first valid kernel position lies, when scanning from
        ' the left boundary.  Calculate a position relative to [0] (basically, remove the kernel's left offset).
        m_XLeft(y) = x - m_XBLeft
        
        'Repeat the above steps, but scanning from the right this time
        x = m_XBSize - 1
        Do
            If (m_ReferenceMap(x, y) > 0) Then Exit Do
            x = x - 1
            If (x < 0) Then
                x = m_XBLeft
                Exit Do
            End If
        Loop
        
        m_XRight(y) = x - m_XBLeft
        
    Next y
    
    'Now we're going to repeat the above steps, but for the y boundaries (instead of x)
    ReDim m_YTop(0 To m_XBSize) As Long
    ReDim m_YBottom(0 To m_XBSize) As Long
    
    For x = 0 To m_XBSize - 1
        
        y = 0
        Do
            If (m_ReferenceMap(x, y) > 0) Then Exit Do
            y = y + 1
            If (y > m_YBSize - 1) Then
                y = m_YBTop
                Exit Do
            End If
        Loop
        
        m_YTop(x) = y - m_YBTop
        
        y = m_YBSize - 1
        Do
            If (m_ReferenceMap(x, y) > 0) Then Exit Do
            y = y - 1
            If (y < 0) Then
                y = m_YBTop
                Exit Do
            End If
        Loop
        
        m_YBottom(x) = y - m_YBTop
        
    Next x

    'We have now generated a map of the edges for each side of the kernel.

End Sub

'Populate an initial window of values for a square shape.  The cope is optimized for this specific use-case, and it's not
' meant to be applied to other shapes!
Private Function GenerateInitialWindow_Square() As Long
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim xStart As Long, xFinal As Long
    xStart = m_InitX * 4
    xFinal = (m_InitX + m_XBRight) * 4
    
    'To improve performance, we're going to wrap a 1D array around the current image scanline.
    ' This doesn't provide a huge performance benefit on small kernels, but on large kernels,
    ' it can be meaningfully faster.
    With m_PixelSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_DibPointer
    End With
    
    PutMem4 VarPtrArray(m_PixelLine()), VarPtr(m_PixelSA1D)
    
    For y = m_InitY To m_InitY + m_YBBottom
        m_PixelSA1D.pvData = m_DibPointer + y * m_dibStride
    For x = xStart To xFinal Step 4
        
        b = m_PixelLine(x)
        g = m_PixelLine(x + 1)
        r = m_PixelLine(x + 2)
        
        If (m_HistogramMode = PDPIM_RGBA) Then
        
            m_Blue(b) = m_Blue(b) + 1
            m_Green(g) = m_Green(g) + 1
            m_Red(r) = m_Red(r) + 1
            
            'If m_AlphaSupport Then
            '    a = m_PixelLine(x + 3)
            '    m_Alpha(a) = m_Alpha(a) + 1
            'End If
            
        Else
            a = GetLuminanceFromRGB(r, g, b)
            m_Luminance(a) = m_Luminance(a) + 1
        End If
        
        m_NumOfPixels = m_NumOfPixels + 1
    
    Next x
    Next y
    
    PutMem4 VarPtrArray(m_PixelLine()), 0&
    
    GenerateInitialWindow_Square = m_NumOfPixels
    
End Function

'Populate an initial window of values for a square shape.  The cope is optimized for this specific use-case,
' and it's not meant to be applied to other shapes!
Private Function GenerateInitialWindowBytes_Square() As Long
    
    Dim x As Long, y As Long, v As Long
    
    Dim xStart As Long, xFinal As Long
    xStart = m_InitX
    xFinal = (m_InitX + m_XBRight)
    
    'To improve performance, we're going to wrap a 1D array around the current image scanline.
    ' This doesn't provide a huge performance benefit on small kernels, but on large kernels,
    ' it can be meaningfully faster.
    With m_BytesSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_ArrayWidth
        .pvData = m_ByteArrayPtr
    End With
    
    PutMem4 VarPtrArray(m_ByteLine()), VarPtr(m_BytesSA1D)
    
    For y = m_InitY To m_InitY + m_YBBottom
        m_BytesSA1D.pvData = m_ByteArrayPtr + y * m_ArrayWidth
    For x = xStart To xFinal
        v = m_ByteLine(x)
        m_Luminance(v) = m_Luminance(v) + 1
        m_NumOfPixels = m_NumOfPixels + 1
    Next x
    Next y
    
    PutMem4 VarPtrArray(m_ByteLine()), 0&
    
    GenerateInitialWindowBytes_Square = m_NumOfPixels
    
End Function

'Populate an initial window of values for an arbitrary shape.
Private Function GenerateInitialWindow_ArbitraryShape() As Long
    
    Dim x As Long, y As Long
    Dim xStart As Long, xFinal As Long
    xStart = m_InitX * 4
    xFinal = (m_InitX + m_XBRight) * 4
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    'To improve performance, we're going to wrap a 1D array around the current image scanline.
    ' This doesn't provide a huge performance benefit on small kernels, but on large kernels,
    ' it can be meaningfully faster.
    With m_PixelSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_DibPointer
    End With
    
    PutMem4 VarPtrArray(m_PixelLine()), VarPtr(m_PixelSA1D)
    
    For y = m_InitY To m_InitY + m_YBBottom
        m_PixelSA1D.pvData = m_DibPointer + y * m_dibStride
    For x = xStart To xFinal Step 4
        
        'Only calculate pixels that lie inside the reference map
        If (m_ReferenceMap(((x \ 4) - m_InitX) + m_XBLeft, (y - m_InitY) + m_YBTop) > 0) Then
            
            b = m_PixelLine(x)
            g = m_PixelLine(x + 1)
            r = m_PixelLine(x + 2)
            
            If (m_HistogramMode = PDPIM_RGBA) Then
            
                m_Blue(b) = m_Blue(b) + 1
                m_Green(g) = m_Green(g) + 1
                m_Red(r) = m_Red(r) + 1
                
                'If m_AlphaSupport Then
                '    a = m_PixelLine(x + 3)
                '    m_Alpha(a) = m_Alpha(a) + 1
                'End If
                            
            Else
                a = GetLuminanceFromRGB(r, g, b)
                m_Luminance(a) = m_Luminance(a) + 1
            End If
            
            m_NumOfPixels = m_NumOfPixels + 1
            
        End If
    
    Next x
    Next y
    
    PutMem4 VarPtrArray(m_PixelLine()), 0&
    
    GenerateInitialWindow_ArbitraryShape = m_NumOfPixels
    
End Function

'Populate an initial window of values for an arbitrary shape.
Private Function GenerateInitialWindowBytes_ArbitraryShape() As Long
    
    Dim x As Long, y As Long, v As Long
    Dim xStart As Long, xFinal As Long
    xStart = m_InitX
    xFinal = (m_InitX + m_XBRight)
    
    'To improve performance, we're going to wrap 1D arrays around both the current image scanline,
    ' and the reference map. This doesn't provide a huge performance benefit on small kernels,
    ' but on large kernels it can be meaningfully faster.
    With m_BytesSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_ArrayWidth
        .pvData = m_ByteArrayPtr
    End With
    
    PutMem4 VarPtrArray(m_ByteLine()), VarPtr(m_BytesSA1D)
    
    For y = m_InitY To m_InitY + m_YBBottom
        m_BytesSA1D.pvData = m_ByteArrayPtr + y * m_ArrayWidth
    For x = xStart To xFinal
        
        'Only calculate pixels that lie inside the reference map
        If (m_ReferenceMap((x - m_InitX) + m_XBLeft, (y - m_InitY) + m_YBTop) > 0) Then
            v = m_ByteLine(x)
            m_Luminance(v) = m_Luminance(v) + 1
            m_NumOfPixels = m_NumOfPixels + 1
        End If
    
    Next x
    Next y
    
    PutMem4 VarPtrArray(m_ByteLine()), 0&
    
    GenerateInitialWindowBytes_ArbitraryShape = m_NumOfPixels
    
End Function

'Given source RGB values, return a corresponding luminance value (high-quality calculation)
Private Function GetLuminanceFromRGB(ByVal srcR As Long, ByVal srcG As Long, ByVal srcB As Long) As Long
    If (m_LuminanceMode = PDLM_Value) Then
        GetLuminanceFromRGB = Colors.GetLuminance(srcR, srcG, srcB)
    Else
        GetLuminanceFromRGB = Colors.GetHQLuminance(srcR, srcG, srcB)
    End If
End Function

'After the caller has successfully processed a pixel, they can call these functions to move to
' the next pixel in the X or Y direction.  Because this class uses serpentine scanning, the caller
' is responsible for changing direction on each line increment.
Friend Function MoveXRight() As Long
    
    Dim yTop As Long, yBottom As Long, y As Long
    Dim xLeft As Long, xRight As Long, x As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate to the right
    m_X = m_X + 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
    
        'Figure out Y bounds first
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom
        If (yTop < m_InitY) Then yTop = m_InitY
        If (yBottom > m_FinalY) Then yBottom = m_FinalY
            
        'If the *left* x-bound is within bounds, remove a line of pixels from the window.
        xLeft = (m_X - m_XBLeft) - 1
        If (xLeft >= m_InitX) Then
            
            x = xLeft * 4
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_Pixels(x + 3, y)
                    '    m_Alpha(a) = m_Alpha(a) - 1
                    'End If
                                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) - 1
                End If
                
            Next y
            
            'Infer the number of pixels removed from the histogram(s)
            m_NumOfPixels = m_NumOfPixels - (yBottom - yTop + 1)
        
        End If
        
        'If the *right* x-bound is within bounds, add a new line of pixels to the window.
        xRight = m_X + m_XBRight
        If (xRight <= m_FinalX) Then
            
            x = xRight * 4
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_Pixels(x + 3, y)
                    '    m_Alpha(a) = m_Alpha(a) + 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) + 1
                End If
                
            Next y
            
            'Infer the number of pixels added to the histogram(s)
            m_NumOfPixels = m_NumOfPixels + (yBottom - yTop) + 1
        
        End If
    
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        
        'Start with trailing pixels
        For y = -m_YBTop To m_YBBottom
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpY = y + m_Y
            
            'If y-coordinate lies out of bounds, ignore it
            If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                'This y-boundary potentially lies in-bounds.  Check the matching (x) position.
                tmpX = (m_X + m_XLeft(y + m_YBTop)) - 1
                
                'If this x-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                    tmpX = tmpX * 4
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) - 1
                        m_Green(g) = m_Green(g) - 1
                        m_Red(r) = m_Red(r) - 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) - 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                
                'Repeat the above steps, but for the matching (x) position on the right.
                tmpX = m_X + m_XRight(y + m_YBTop)
                
                'If this x-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                    tmpX = tmpX * 4
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) + 1
                        m_Green(g) = m_Green(g) + 1
                        m_Red(r) = m_Red(r) + 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) + 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                
            End If
        
        Next y
    
    End If
    
    MoveXRight = m_NumOfPixels
    
End Function

'Equivalent of MoveXRight(), above, but designed against non-DIB byte arrays
Friend Function MoveXRight_Byte() As Long
    
    Dim yTop As Long, yBottom As Long, y As Long
    Dim xLeft As Long, xRight As Long
    Dim v As Long
    
    'Move our target pixel coordinate to the right
    m_X = m_X + 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
    
        'Figure out Y bounds first
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom
        If (yTop < m_InitY) Then yTop = m_InitY
        If (yBottom > m_FinalY) Then yBottom = m_FinalY
        
        'Next, figure out X bounds
        xLeft = (m_X - m_XBLeft) - 1
        xRight = m_X + m_XBRight
            
        'If the *left* x-bound is within bounds, remove a line of pixels from the window.
        If (xLeft >= m_InitX) Then
            For y = yTop To yBottom
                v = m_Bytes(xLeft, y)
                m_Luminance(v) = m_Luminance(v) - 1
                m_NumOfPixels = m_NumOfPixels - 1
            Next y
        End If
        
        'If the *right* x-bound is within bounds, add a new line of pixels to the window.
        If (xRight <= m_FinalX) Then
            For y = yTop To yBottom
                v = m_Bytes(xRight, y)
                m_Luminance(v) = m_Luminance(v) + 1
                m_NumOfPixels = m_NumOfPixels + 1
            Next y
        End If
    
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        
        'Start with trailing pixels
        For y = -m_YBTop To m_YBBottom
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpY = y + m_Y
            
            'If y-coordinate lies out of bounds, ignore it
            If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                'This y-boundary potentially lies in-bounds.  Check the matching (x) position.
                tmpX = (m_X + m_XLeft(y + m_YBTop)) - 1
                
                'If this x-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) - 1
                    m_NumOfPixels = m_NumOfPixels - 1
                End If
                
                'Repeat the above steps, but for the matching (x) position on the right.
                tmpX = m_X + m_XRight(y + m_YBTop)
                
                'If this x-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) + 1
                    m_NumOfPixels = m_NumOfPixels + 1
                End If
                
            End If
        
        Next y
    
    End If
    
    MoveXRight_Byte = m_NumOfPixels
    
End Function

Friend Function MoveXLeft() As Long
    
    Dim yTop As Long, yBottom As Long, y As Long
    Dim xLeft As Long, xRight As Long, x As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate to the left
    m_X = m_X - 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
    
        'Figure out Y bounds first
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom
        If (yTop < m_InitY) Then yTop = m_InitY
        If (yBottom > m_FinalY) Then yBottom = m_FinalY
            
        'If the *left* x-bound is within bounds, add that line of pixels to the window.
        xLeft = (m_X - m_XBLeft)
        If (xLeft >= m_InitX) Then
            
            x = xLeft * 4
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_Pixels(x + 3, y)
                    '    m_Alpha(a) = m_Alpha(a) + 1
                    'End If
                                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) + 1
                End If
                
            Next y
            
            'Infer the number of pixels added to the histogram(s)
            m_NumOfPixels = m_NumOfPixels + (yBottom - yTop) + 1
        
        End If
        
        'If the *right* x-bound is within bounds, remove that line of pixels from the window.
        xRight = m_X + m_XBRight + 1
        If (xRight <= m_FinalX) Then
            
            x = xRight * 4
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_Pixels(x + 3, y)
                    '    m_Alpha(a) = m_Alpha(a) - 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) - 1
                End If
                
            Next y
            
            'Infer the number of pixels removed from the histogram(s)
            m_NumOfPixels = m_NumOfPixels - (yBottom - yTop + 1)
        
        End If
    
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel
        ' needs to be added or removed.  The only way to know is to scan each boundary pixel
        ' in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        
        'Start with trailing pixels
        For y = -m_YBTop To m_YBBottom
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpY = y + m_Y
            
            'If y-coordinate lies out of bounds, ignore it
            If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                'This y-boundary potentially lies in-bounds.  Check the matching (x) position.
                tmpX = (m_X + m_XLeft(y + m_YBTop))
                
                'If this x-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                    tmpX = tmpX * 4
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) + 1
                        m_Green(g) = m_Green(g) + 1
                        m_Red(r) = m_Red(r) + 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) + 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                
                'Repeat the above steps, but for the matching (x) position on the right.
                tmpX = m_X + m_XRight(y + m_YBTop) + 1
                
                'If this x-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                    tmpX = tmpX * 4
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) - 1
                        m_Green(g) = m_Green(g) - 1
                        m_Red(r) = m_Red(r) - 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) - 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                
            End If
        
        Next y
    
    End If
    
    MoveXLeft = m_NumOfPixels
    
End Function

Friend Function MoveYDown() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate down
    m_Y = m_Y + 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
    
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        If (xLeft < m_InitX) Then xLeft = m_InitX
        xRight = m_X + m_XBRight
        If (xRight > m_FinalX) Then xRight = m_FinalX
        
        'To improve performance, we're going to wrap a 1D array around the current image scanline.
        ' This doesn't provide a huge performance benefit on small kernels, but on large kernels,
        ' it can be meaningfully faster. (Note that the SafeArray header was already populated
        ' correctly in the initialization step, so we don't need to initialize it here.)
        PutMem4 VarPtrArray(m_PixelLine()), VarPtr(m_PixelSA1D)
        
        'If the *top* y-bound is within bounds, remove a line of pixels from the window.
        yTop = (m_Y - m_YBTop) - 1
        If (yTop >= 0) Then
            
            m_PixelSA1D.pvData = m_DibPointer + yTop * m_dibStride
            
            For x = xLeft * 4 To xRight * 4 Step 4
                
                b = m_PixelLine(x)
                g = m_PixelLine(x + 1)
                r = m_PixelLine(x + 2)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                    
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_PixelLine(x + 3)
                    '    m_Alpha(a) = m_Alpha(a) - 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) - 1
                End If
                
            Next x
            
            'Infer the number of pixels removed from the histogram(s)
            m_NumOfPixels = m_NumOfPixels - (xRight - xLeft + 1)
                
        End If
        
        'If the *bottom* y-bound is within bounds, add a new line of pixels to the window.
        yBottom = m_Y + m_YBBottom
        If (yBottom <= m_FinalY) Then
            
            m_PixelSA1D.pvData = m_DibPointer + yBottom * m_dibStride
            
            For x = xLeft * 4 To xRight * 4 Step 4
                
                b = m_PixelLine(x)
                g = m_PixelLine(x + 1)
                r = m_PixelLine(x + 2)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_PixelLine(x + 3)
                    '    m_Alpha(a) = m_Alpha(a) + 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) + 1
                End If
                
            Next x
            
            'Infer the number of pixels added to the histogram(s)
            m_NumOfPixels = m_NumOfPixels + (xRight - xLeft) + 1
        
        End If
        
        'Regardless of pixels added or removed, we need to free our array reference before exiting
        PutMem4 VarPtrArray(m_PixelLine()), 0&
        
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        tmpY = m_Y
        
        'Start with trailing pixels
        For x = -m_XBLeft To m_XBRight
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpX = x + m_X
            
            'If this x-coordinate lies out of bounds, ignore it
            If (tmpX >= m_InitX) Then
            If (tmpX <= m_FinalX) Then
                
                tmpX = tmpX * 4
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YTop(x + m_XBLeft) - 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) - 1
                        m_Green(g) = m_Green(g) - 1
                        m_Red(r) = m_Red(r) - 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) - 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YBottom(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                    
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) + 1
                        m_Green(g) = m_Green(g) + 1
                        m_Red(r) = m_Red(r) + 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) + 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                End If
                
            End If
            End If
        
        Next x
        
    End If
    
    MoveYDown = m_NumOfPixels
    
End Function

Friend Function MoveYDown_Byte() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long
    Dim v As Long
    
    'Move our target pixel coordinate down
    m_Y = m_Y + 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
    
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        If (xLeft < m_InitX) Then xLeft = m_InitX
        xRight = m_X + m_XBRight
        If (xRight > m_FinalX) Then xRight = m_FinalX
        
        'Next, figure out Y bounds.
        yTop = (m_Y - m_YBTop) - 1
        yBottom = m_Y + m_YBBottom
        
        'To improve performance, we're going to wrap a 1D array around the current image scanline.  This doesn't provide a huge
        ' performance benefit on small kernels, but on large kernels, it can be meaningfully faster.
        ' (Note that the SafeArray header was already populated correctly in the initialization step, so we don't need to
        '  initialize it here.)
        PutMem4 VarPtrArray(m_ByteLine()), VarPtr(m_BytesSA1D)
        
        'If the *top* y-bound is within bounds, remove a line of pixels from the window.
        If (yTop >= 0) Then
            
            m_BytesSA1D.pvData = m_ByteArrayPtr + yTop * m_ArrayWidth
            
            For x = xLeft To xRight
                v = m_ByteLine(x)
                m_Luminance(v) = m_Luminance(v) - 1
                m_NumOfPixels = m_NumOfPixels - 1
            Next x
        
        End If
        
        'If the *bottom* y-bound is within bounds, add a new line of pixels to the window.
        If (yBottom <= m_FinalY) Then
            
            m_BytesSA1D.pvData = m_ByteArrayPtr + yBottom * m_ArrayWidth
            
            For x = xLeft To xRight
                v = m_ByteLine(x)
                m_Luminance(v) = m_Luminance(v) + 1
                m_NumOfPixels = m_NumOfPixels + 1
            Next x
        
        End If
        
        PutMem4 VarPtrArray(m_ByteLine()), 0&
        
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        tmpY = m_Y
        
        'Start with trailing pixels
        For x = -m_XBLeft To m_XBRight
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpX = x + m_X
            
            'If this x-coordinate lies out of bounds, ignore it
            If (tmpX >= m_InitX) Then
            If (tmpX <= m_FinalX) Then
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YTop(x + m_XBLeft) - 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) - 1
                    m_NumOfPixels = m_NumOfPixels - 1
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YBottom(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) + 1
                    m_NumOfPixels = m_NumOfPixels + 1
                End If
            
            End If
            End If
        
        Next x
        
    End If
    
    MoveYDown_Byte = m_NumOfPixels
    
End Function

Friend Function MoveYUp() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate up
    m_Y = m_Y - 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
        
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        If (xLeft < m_InitX) Then xLeft = m_InitX
        
        xRight = m_X + m_XBRight
        If (xRight > m_FinalX) Then xRight = m_FinalX
        
        'To improve performance, we're going to wrap a 1D array around the current image scanline.  This doesn't provide a huge
        ' performance benefit on small kernels, but on large kernels, it can be meaningfully faster.
        ' (Note that the SafeArray header was already populated correctly in the initialization step, so we don't need to
        '  initialize it here.)
        PutMem4 VarPtrArray(m_PixelLine()), VarPtr(m_PixelSA1D)
            
        'If the *bottom* y-bound is within bounds, remove a new line of pixels from the window.
        yBottom = m_Y + m_YBBottom + 1
        If (yBottom <= m_FinalY) Then
            
            m_PixelSA1D.pvData = m_DibPointer + yBottom * m_dibStride
            
            For x = xLeft * 4 To xRight * 4 Step 4
                
                b = m_PixelLine(x)
                g = m_PixelLine(x + 1)
                r = m_PixelLine(x + 2)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_PixelLine(x + 3)
                    '    m_Alpha(a) = m_Alpha(a) - 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) - 1
                End If
                
            Next x
            
            'Infer the number of pixels removed from the histogram(s)
            m_NumOfPixels = m_NumOfPixels - (xRight - xLeft + 1)
        
        End If
        
        'If the *top* y-bound is within bounds, add a line of pixels to the window.
        yTop = m_Y - m_YBTop
        If (yTop >= 0) Then
            
            m_PixelSA1D.pvData = m_DibPointer + yTop * m_dibStride
            
            For x = xLeft * 4 To xRight * 4 Step 4
                
                b = m_PixelLine(x)
                g = m_PixelLine(x + 1)
                r = m_PixelLine(x + 2)
                
                If (m_HistogramMode = PDPIM_RGBA) Then
                
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    'If m_AlphaSupport Then
                    '    a = m_PixelLine(x + 3)
                    '    m_Alpha(a) = m_Alpha(a) + 1
                    'End If
                    
                Else
                    a = GetLuminanceFromRGB(r, g, b)
                    m_Luminance(a) = m_Luminance(a) + 1
                End If
                
            Next x
            
            'Infer the number of pixels added to the histogram(s)
            m_NumOfPixels = m_NumOfPixels + (xRight - xLeft) + 1
        
        End If
        
        PutMem4 VarPtrArray(m_PixelLine()), 0&
        
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        tmpY = m_Y
        
        'Start with trailing pixels
        For x = -m_XBLeft To m_XBRight
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpX = x + m_X
            
            'If this x-coordinate lies out of bounds, ignore it
            If (tmpX >= m_InitX) Then
            If (tmpX <= m_FinalX) Then
                
                tmpX = tmpX * 4
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YBottom(x + m_XBLeft) + 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) - 1
                        m_Green(g) = m_Green(g) - 1
                        m_Red(r) = m_Red(r) - 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) - 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YTop(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    If (m_HistogramMode = PDPIM_RGBA) Then
                    
                        m_Blue(b) = m_Blue(b) + 1
                        m_Green(g) = m_Green(g) + 1
                        m_Red(r) = m_Red(r) + 1
                        
                        'If m_AlphaSupport Then
                        '    a = m_Pixels(tmpX + 3, tmpY)
                        '    m_Alpha(a) = m_Alpha(a) + 1
                        'End If
                        
                    Else
                        a = GetLuminanceFromRGB(r, g, b)
                        m_Luminance(a) = m_Luminance(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                End If
                
            End If
            End If
        
        Next x
        
    End If
    
    MoveYUp = m_NumOfPixels
    
End Function

Friend Function MoveYUp_Byte() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long
    Dim v As Long
    
    'Move our target pixel coordinate up
    m_Y = m_Y - 1
    
    'Rectangular regions get special, optimized treatment
    If (m_WindowShape = PDPRS_Rectangle) Then
        
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        If (xLeft < m_InitX) Then xLeft = m_InitX
        
        xRight = m_X + m_XBRight
        If (xRight > m_FinalX) Then xRight = m_FinalX
        
        'Next, figure out Y bounds.
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom + 1
        
        'To improve performance, we're going to wrap a 1D array around the current image scanline.  This doesn't provide a huge
        ' performance benefit on small kernels, but on large kernels, it can be meaningfully faster.
        ' (Note that the SafeArray header was already populated correctly in the initialization step, so we don't need to
        '  initialize it here.)
        PutMem4 VarPtrArray(m_ByteLine()), VarPtr(m_BytesSA1D)
            
        'If the *bottom* y-bound is within bounds, remove a new line of pixels from the window.
        If (yBottom <= m_FinalY) Then
            
            m_BytesSA1D.pvData = m_ByteArrayPtr + yBottom * m_ArrayWidth
            
            For x = xLeft To xRight
                v = m_ByteLine(x)
                m_Luminance(v) = m_Luminance(v) - 1
                m_NumOfPixels = m_NumOfPixels - 1
            Next x
        
        End If
        
        'If the *top* y-bound is within bounds, add a line of pixels to the window.
        If (yTop >= 0) Then
            
            m_BytesSA1D.pvData = m_ByteArrayPtr + yTop * m_ArrayWidth
            
            For x = xLeft To xRight
                v = m_ByteLine(x)
                m_Luminance(v) = m_Luminance(v) + 1
                m_NumOfPixels = m_NumOfPixels + 1
            Next x
        
        End If
        
        PutMem4 VarPtrArray(m_ByteLine()), 0&
        
    'Non-rectangular kernels require custom handling
    Else
    
        'Because kernels may be non-standard sizes, we don't know in advance if a given pixel needs to be added
        ' or removed.  The only way to know is to scan each boundary pixel in turn, and see if it lies in-bounds.
        Dim tmpY As Long, tmpX As Long
        tmpY = m_Y
        
        'Start with trailing pixels
        For x = -m_XBLeft To m_XBRight
        
            'Calculate the position of the trailing boundary pixel in this column
            tmpX = x + m_X
            
            'If this x-coordinate lies out of bounds, ignore it
            If (tmpX >= m_InitX) Then
            If (tmpX <= m_FinalX) Then
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YBottom(x + m_XBLeft) + 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) - 1
                    m_NumOfPixels = m_NumOfPixels - 1
                End If
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YTop(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) Then
                If (tmpY <= m_FinalY) Then
                    v = m_Bytes(tmpX, tmpY)
                    m_Luminance(v) = m_Luminance(v) + 1
                    m_NumOfPixels = m_NumOfPixels + 1
                End If
                End If
                
            End If
            End If
        
        Next x
        
    End If
    
    MoveYUp_Byte = m_NumOfPixels
    
End Function
