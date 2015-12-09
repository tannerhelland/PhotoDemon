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
'Copyright 2015-2015 by Tanner Helland
'Created: 04/December/14
'Last updated: 08/December/15
'Last update: add circular kernel support
'
'By the time the 7.0 release rolled around, a ridiculous number of PD effects were managing their own pixel iteration
' methods.  Most of these were a variant of a "sliding window" implementation, where a running histogram is maintained
' for a given pixel region, and when moving to the next pixel, instead of recalculating the entire region from scratch,
' the histograms from the previous pixel are simply updated against the new region, typically by subtracting a line of
' pixels on the left or top, and adding a new line of pixels on the right or bottom.
'
'These methods tended to use a single copy+paste chunk of code that was pretty advanced, supporting serpentine scanning
' and a number of other neat features, but maintenance quickly became unwieldy because any improvements to the method
' required copy+pasting the changes across dozens of functions.  Worse still, the function only supported rectangular
' regions, and a lot of image region functions look more natural when circular.
'
'So as part of the 7.0 release, this class was created.  This class is designed to abstract away the messy duties of
' iterating per-pixel regions, while supporting a bunch more features than PD's old implementation.  Because this class
' uses generic histograms, many functions can tap into it, without needing to modify the actual iteration code.
'
'Please note that this class has several dependencies throughout PD, including pdDIB, and some specialized enums
' (located in PD's central Public_Enums_and_Types module).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub SafeArrayLock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Sub SafeArrayUnlock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal Addr As Long, ByRef dstValue As Long) As Long

'Current (x, y) position of the iterator.  Call InitializeIterator() to reset these to (0, 0).
Private m_X As Long, m_Y As Long

'Initial and final loop boundaries
Private m_InitX As Long, m_InitY As Long
Private m_FinalX As Long, m_FinalY As Long

'Shape of the current iterator, including boundary distance
Private m_WindowShape As PD_PIXEL_REGION_SHAPE
Private m_XBLeft As Long, m_XBRight As Long, m_XBSize As Long
Private m_YBTop As Long, m_YBBottom As Long, m_YBSize As Long

'Other relevant DIB properties
Private m_DibPointer As Long, m_dibHeight As Long, m_DibStride As Long, m_DibColorDepth As Long, m_PixelWidth As Long

'Pixel array (alias only; the actual pixel bits are stored by the caller, and treated as read-only by this class)
Private m_PixelSA As SAFEARRAY2D, m_Pixels() As Byte

'If the function wants alpha tracked, this will be set to TRUE
Private m_AlphaSupport As Boolean

'Histogram arrays.  Managing these are a bit tricky, because these are simply "trick" wrappers against arrays
' provided by the caller.
Private m_RedSA As SAFEARRAY1D, m_GreenSA As SAFEARRAY1D, m_BlueSA As SAFEARRAY1D, m_AlphaSA As SAFEARRAY1D
Private m_Red() As Long, m_Green() As Long, m_Blue() As Long, m_Alpha() As Long

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
Friend Function InitializeIterator(ByRef targetDIB As pdDIB, ByVal xRadius As Long, ByVal yRadius As Long, Optional ByVal windowShape As PD_PIXEL_REGION_SHAPE = PDPRS_Rectangle) As Boolean

    'Reset the iterator coordinates and pixel count
    m_X = 0
    m_Y = 0
    m_NumOfPixels = 0
    
    'Cache loop boundaries
    m_InitX = 0
    m_InitY = 0
    m_FinalX = targetDIB.getDIBWidth - 1
    m_FinalY = targetDIB.getDIBHeight - 1
    
    'TODO: use magic to generate a shape map for each direction, including bounds
    ' (right now, during testing, only rectangular regions are supported)
    m_WindowShape = windowShape
    m_XBLeft = xRadius
    m_XBRight = xRadius
    m_YBTop = yRadius
    m_YBBottom = yRadius
    
    'Apply some failsafe dimension testing to the incoming bounds
    If m_XBRight > (m_FinalX - m_InitX) Then
        m_XBRight = (m_FinalX - m_InitX)
        m_XBLeft = m_XBRight
    End If
    
    If m_YBBottom > (m_FinalY - m_InitY) Then
        m_YBBottom = (m_FinalY - m_InitY)
        m_YBTop = m_YBBottom
    End If
    
    'Store the final kernel size
    m_XBSize = m_XBLeft + m_XBRight + 1
    m_YBSize = m_YBTop + m_YBBottom + 1
    
    'Retrieve other relevant DIB properties
    m_DibPointer = targetDIB.getActualDIBBits
    m_dibHeight = targetDIB.getDIBHeight
    m_DibStride = targetDIB.getDIBArrayWidth
    m_DibColorDepth = targetDIB.getDIBColorDepth
    m_PixelWidth = m_DibColorDepth \ 8
    
    'Set alpha to a default value, based on the source image's color depth
    If m_DibColorDepth = 32 Then m_AlphaSupport = True Else m_AlphaSupport = False
    
    InitializeIterator = True
    
End Function

'After you've initialized the iterator, call this function to setup the initial pixel region.  The caller must supply
' their own histogram arrays; we will wrap these with some "trick" internal array references, to avoid the need for
' passing these on every pixel request.
'
'IMPORTANT NOTE: PRIOR TO CALLING THIS FUNCTION, you must redim these arrays to range [0, 255].  Do not ReDim them until
' you have completed your function and freed the histograms safely (via ReleaseTargetHistograms, below).
'
'This function will return the pixel count of the first window in the image.  DO NOT PROCEED if it returns zero.
Friend Function LockTargetHistograms(ByRef dstRed() As Long, ByRef dstGreen() As Long, ByRef dstBlue() As Long, ByRef dstAlpha() As Long, Optional ByVal calcAlpha As Boolean = True) As Long
    
    'Alias our internal histogram arrays around the destination ones.  As you might expect, you MUST NOT attempt
    ' to erase or ReDim the target arrays until the iterator has finished.
    Alias1DArray dstRed, m_Red, m_RedSA
    Alias1DArray dstGreen, m_Green, m_GreenSA
    Alias1DArray dstBlue, m_Blue, m_BlueSA
    
    If calcAlpha And m_AlphaSupport Then
        Alias1DArray dstAlpha, m_Alpha, m_AlphaSA
    Else
        m_AlphaSupport = False
    End If
    
    'Point our internal 2D pixel array at the target DIB
    With m_PixelSA
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_DibStride
        .pvData = m_DibPointer
    End With
    CopyMemory ByVal VarPtrArray(m_Pixels()), VarPtr(m_PixelSA), 4&
    
    'Generate the initial window
    Select Case m_WindowShape
        
        Case PDPRS_Rectangle
            LockTargetHistograms = GenerateInitialWindow_Square()
        
        Case PDPRS_Circle
            
            'Generate an initial shape
            CreateShape PDPRS_Circle
            
            'Generate boundary indices
            FindBoundaries
            
            'Generate the initial histogram window
            LockTargetHistograms = GenerateInitialWindow_ArbitraryShape()
        
    End Select
    
End Function

'When the iterator is finished (due to any condition - success, error, etc), the caller MUST call this function to
' release our aliases to their histogram arrays and DIB.
Friend Function ReleaseTargetHistograms(ByRef dstRed() As Long, ByRef dstGreen() As Long, ByRef dstBlue() As Long, ByRef dstAlpha() As Long)

    Unalias1DArray dstRed, m_Red
    Unalias1DArray dstGreen, m_Green
    Unalias1DArray dstBlue, m_Blue
    If m_AlphaSupport Then Unalias1DArray dstAlpha, m_Alpha
    
    'While we're here, release our DIB reference, too
    CopyMemory ByVal VarPtrArray(m_Pixels()), 0&, 4&
    
End Function

'Point an internal 1D array at some other 1D array.  Any arrays aliased this way must be freed via Unalias1DArray,
' or VB will crash.
Private Sub Alias1DArray(ByRef orig1DArray() As Long, ByRef new1DArray() As Long, ByRef newArraySA As SAFEARRAY1D)
    
    'Retrieve a copy of the original 1D array's SafeArray struct
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig1DArray()), ptrSrc
    CopyMemory ByVal VarPtr(newArraySA), ByVal ptrSrc, LenB(newArraySA)
    
    'newArraySA now contains the full SafeArray of the original array.  Copy this over our current array.
    CopyMemory ByVal VarPtrArray(new1DArray()), VarPtr(newArraySA), 4&
    
    'Add a lock to the original array, to prevent potential crashes from unknowing users.  (Thanks to @Kroc for this tip.)
    SafeArrayLock ptrSrc
    
End Sub

'Counterpart to Alias1DArray, above.  Do NOT call this function on arrays that were not originally processed by that function.
Private Sub Unalias1DArray(ByRef orig1DArray() As Long, ByRef new1DArray() As Long)
    
    'Wipe the array pointer
    CopyMemory ByVal VarPtrArray(new1DArray), 0&, 4&
    
    'Remove a lock from the original array; this allows the user to safely release the array on their own terms
    Dim ptrSrc As Long
    GetMem4 VarPtrArray(orig1DArray()), ptrSrc
    SafeArrayUnlock ptrSrc
    
End Sub

'Create the reference map for a non-rectangular kernel
Private Sub CreateShape(ByVal srcShape As PD_PIXEL_REGION_SHAPE)
    
    If m_ShapeImage Is Nothing Then Set m_ShapeImage = New pdDIB
    m_ShapeImage.createBlank m_XBSize, m_YBSize, 24, 0&
    
    Select Case srcShape
        
        'At small sizes, the GDI+ circle rendering algorithm doesn't work very well.  To bypass this, we activate
        ' antialiasing when the search radius is only 1 px wide (for a net search diameter of 3).
        Case PDPRS_Circle
            Dim useAA As Boolean
            If (m_XBSize <= 3) Or (m_YBSize <= 3) Then useAA = True Else useAA = False
            GDI_Plus.GDIPlusFillEllipseToDC m_ShapeImage.getDIBDC, 0, 0, m_ShapeImage.getDIBWidth, m_ShapeImage.getDIBHeight, RGB(255, 255, 255), useAA, 255, True
            
    End Select
    
    'TODO: as a failsafe, scan the DIB and trim any empty lines, because they'll grossly mess up our calculations.
    '      (Weird shapes may not extend all the way to the edge of the target DIB, and it's difficult to predict this
    '       in advance.)
    
    'Convert the reference image into a 2D byte array, which is much faster to access and process.
    DIB_Handler.GetDIBGrayscaleMap m_ShapeImage, m_ReferenceMap
    m_ShapeImage.eraseDIB
    
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
        
            If m_ReferenceMap(x, y) > 0 Then Exit Do
            
            'Move to the next column
            x = x + 1
            
            'As a failsafe for blank lines, set a dummy value if the entire line is blank
            If x > m_XBSize - 1 Then
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
            If m_ReferenceMap(x, y) > 0 Then Exit Do
            x = x - 1
            If x < 0 Then
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
            If m_ReferenceMap(x, y) > 0 Then Exit Do
            y = y + 1
            If y > m_YBSize - 1 Then
                y = m_YBTop
                Exit Do
            End If
        Loop
        
        m_YTop(x) = y - m_YBTop
        
        y = m_YBSize - 1
        Do
            If m_ReferenceMap(x, y) > 0 Then Exit Do
            y = y - 1
            If y < 0 Then
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
    
    Dim x As Long, y As Long, QuickX As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    For y = m_InitY To m_InitY + m_YBBottom
    For x = m_InitX * m_PixelWidth To (m_InitX + m_XBRight) * m_PixelWidth Step m_PixelWidth
        
        b = m_Pixels(x, y)
        g = m_Pixels(x + 1, y)
        r = m_Pixels(x + 2, y)
        
        m_Blue(b) = m_Blue(b) + 1
        m_Green(g) = m_Green(g) + 1
        m_Red(r) = m_Red(r) + 1
        
        If m_AlphaSupport Then
            a = m_Pixels(x + 3, y)
            m_Alpha(a) = m_Alpha(a) + 1
        End If
        
        m_NumOfPixels = m_NumOfPixels + 1
    
    Next x
    Next y
    
    GenerateInitialWindow_Square = m_NumOfPixels
    
End Function

'Populate an initial window of values for a square shape.  The cope is optimized for this specific use-case, and it's not
' meant to be applied to other shapes!
Private Function GenerateInitialWindow_ArbitraryShape() As Long
    
    Dim x As Long, y As Long, QuickX As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    For y = m_InitY To m_InitY + m_YBBottom
    For x = m_InitX To m_InitX + m_XBRight
        
        'Only calculate pixels that lie inside the reference map
        If m_ReferenceMap((x - m_InitX) + m_XBLeft, (y - m_InitY) + m_YBTop) > 0 Then
            
            QuickX = x * m_PixelWidth
            b = m_Pixels(QuickX, y)
            g = m_Pixels(QuickX + 1, y)
            r = m_Pixels(QuickX + 2, y)
            
            m_Blue(b) = m_Blue(b) + 1
            m_Green(g) = m_Green(g) + 1
            m_Red(r) = m_Red(r) + 1
            
            If m_AlphaSupport Then
                a = m_Pixels(QuickX + 3, y)
                m_Alpha(a) = m_Alpha(a) + 1
            End If
            
            m_NumOfPixels = m_NumOfPixels + 1
            
        End If
    
    Next x
    Next y
    
    GenerateInitialWindow_ArbitraryShape = m_NumOfPixels
    
End Function

'After the caller has successfully processed a pixel, they can call these functions to move to the next pixel
' in the X or Y direction.  Because this class uses serpentine scanning, the caller is responsible for changing
' direction on each Y increment.
Friend Function MoveXRight() As Long
    
    Dim yTop As Long, yBottom As Long, y As Long
    Dim xLeft As Long, xRight As Long, x As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate to the right
    m_X = m_X + 1
    
    'Rectangular regions get special, optimized treatment
    If m_WindowShape = PDPRS_Rectangle Then
    
        'Figure out Y bounds first
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom
        If yTop < m_InitY Then yTop = m_InitY
        If yBottom > m_FinalY Then yBottom = m_FinalY
        
        'Next, figure out X bounds
        xLeft = (m_X - m_XBLeft) - 1
        xRight = m_X + m_XBRight
            
        'If the *left* x-bound is within bounds, remove a line of pixels from the window.
        If xLeft >= m_InitX Then
            
            x = xLeft * m_PixelWidth
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                m_Blue(b) = m_Blue(b) - 1
                m_Green(g) = m_Green(g) - 1
                m_Red(r) = m_Red(r) - 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, y)
                    m_Alpha(a) = m_Alpha(a) - 1
                End If
                
                m_NumOfPixels = m_NumOfPixels - 1
                
            Next y
        
        End If
        
        'If the *right* x-bound is within bounds, add a new line of pixels to the window.
        If xRight <= m_FinalX Then
            
            x = xRight * m_PixelWidth
            For y = yTop To yBottom
                
                b = m_Pixels(x, y)
                g = m_Pixels(x + 1, y)
                r = m_Pixels(x + 2, y)
                
                m_Blue(b) = m_Blue(b) + 1
                m_Green(g) = m_Green(g) + 1
                m_Red(r) = m_Red(r) + 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, y)
                    m_Alpha(a) = m_Alpha(a) + 1
                End If
                
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
                
                    tmpX = tmpX * m_PixelWidth
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                
                'Repeat the above steps, but for the matching (x) position on the right.
                tmpX = m_X + m_XRight(y + m_YBTop)
                
                'If this x-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                    tmpX = tmpX * m_PixelWidth
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                
            End If
        
        Next y
    
    End If
    
    MoveXRight = m_NumOfPixels
    
End Function

Friend Function MoveYDown() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long, y As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate down
    m_Y = m_Y + 1
    
    'Rectangular regions get special, optimized treatment
    If m_WindowShape = PDPRS_Rectangle Then
    
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        xRight = m_X + m_XBRight
        If xLeft < m_InitX Then xLeft = m_InitX
        If xRight > m_FinalX Then xRight = m_FinalX
        
        'Next, figure out Y bounds.
        yTop = (m_Y - m_YBTop) - 1
        yBottom = m_Y + m_YBBottom
        
        'If the *top* y-bound is within bounds, remove a line of pixels from the window.
        If yTop >= 0 Then
            
            For x = xLeft * m_PixelWidth To xRight * m_PixelWidth Step m_PixelWidth
                
                b = m_Pixels(x, yTop)
                g = m_Pixels(x + 1, yTop)
                r = m_Pixels(x + 2, yTop)
                
                m_Blue(b) = m_Blue(b) - 1
                m_Green(g) = m_Green(g) - 1
                m_Red(r) = m_Red(r) - 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, yTop)
                    m_Alpha(a) = m_Alpha(a) - 1
                End If
                
                m_NumOfPixels = m_NumOfPixels - 1
                
            Next x
        
        End If
        
        'If the *bottom* y-bound is within bounds, add a new line of pixels to the window.
        If yBottom <= m_FinalY Then
            
            For x = xLeft * m_PixelWidth To xRight * m_PixelWidth Step m_PixelWidth
                
                b = m_Pixels(x, yBottom)
                g = m_Pixels(x + 1, yBottom)
                r = m_Pixels(x + 2, yBottom)
                
                m_Blue(b) = m_Blue(b) + 1
                m_Green(g) = m_Green(g) + 1
                m_Red(r) = m_Red(r) + 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, yBottom)
                    m_Alpha(a) = m_Alpha(a) + 1
                End If
                
                m_NumOfPixels = m_NumOfPixels + 1
                
            Next x
        
        End If
        
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
            If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                tmpX = tmpX * m_PixelWidth
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YTop(x + m_XBLeft) - 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YBottom(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                    
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                
            End If
        
        Next x
        
    End If
    
    MoveYDown = m_NumOfPixels
    
End Function

Friend Function MoveYUp() As Long
    
    Dim xLeft As Long, xRight As Long, x As Long
    Dim yTop As Long, yBottom As Long, y As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Move our target pixel coordinate up
    m_Y = m_Y - 1
    
    'Rectangular regions get special, optimized treatment
    If m_WindowShape = PDPRS_Rectangle Then
        
        'Figure out X bounds first.
        xLeft = m_X - m_XBLeft
        xRight = m_X + m_XBRight
        If xLeft < m_InitX Then xLeft = m_InitX
        If xRight > m_FinalX Then xRight = m_FinalX
        
        'Next, figure out Y bounds.
        yTop = m_Y - m_YBTop
        yBottom = m_Y + m_YBBottom + 1
            
        'If the *bottom* y-bound is within bounds, remove a new line of pixels from the window.
        If yBottom <= m_FinalY Then
            
            For x = xLeft * m_PixelWidth To xRight * m_PixelWidth Step m_PixelWidth
                
                b = m_Pixels(x, yBottom)
                g = m_Pixels(x + 1, yBottom)
                r = m_Pixels(x + 2, yBottom)
                
                m_Blue(b) = m_Blue(b) - 1
                m_Green(g) = m_Green(g) - 1
                m_Red(r) = m_Red(r) - 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, yBottom)
                    m_Alpha(a) = m_Alpha(a) - 1
                End If
                
                m_NumOfPixels = m_NumOfPixels - 1
                
            Next x
        
        End If
        
        'If the *top* y-bound is within bounds, add a line of pixels to the window.
        If yTop >= 0 Then
            
            For x = xLeft * m_PixelWidth To xRight * m_PixelWidth Step m_PixelWidth
                
                b = m_Pixels(x, yTop)
                g = m_Pixels(x + 1, yTop)
                r = m_Pixels(x + 2, yTop)
                
                m_Blue(b) = m_Blue(b) + 1
                m_Green(g) = m_Green(g) + 1
                m_Red(r) = m_Red(r) + 1
                
                If m_AlphaSupport Then
                    a = m_Pixels(x + 3, yTop)
                    m_Alpha(a) = m_Alpha(a) + 1
                End If
                
                m_NumOfPixels = m_NumOfPixels + 1
                
            Next x
        
        End If
        
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
            If (tmpX >= m_InitX) And (tmpX <= m_FinalX) Then
                
                tmpX = tmpX * m_PixelWidth
                
                'This x-boundary potentially lies in-bounds.  Check the matching y-position.
                tmpY = m_Y + m_YBottom(x + m_XBLeft) + 1
                
                'If this y-coordinate lies in-bounds, this pixel can be removed from the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) - 1
                    m_Green(g) = m_Green(g) - 1
                    m_Red(r) = m_Red(r) - 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) - 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels - 1
                    
                End If
                
                'Repeat the above steps, but for the matching y-position on the bottom.
                tmpY = m_Y + m_YTop(x + m_XBLeft)
                
                'If this y-coordinate lies in-bounds, this pixel can be added to the kernel
                If (tmpY >= m_InitY) And (tmpY <= m_FinalY) Then
                
                    b = m_Pixels(tmpX, tmpY)
                    g = m_Pixels(tmpX + 1, tmpY)
                    r = m_Pixels(tmpX + 2, tmpY)
                    
                    m_Blue(b) = m_Blue(b) + 1
                    m_Green(g) = m_Green(g) + 1
                    m_Red(r) = m_Red(r) + 1
                    
                    If m_AlphaSupport Then
                        a = m_Pixels(tmpX + 3, tmpY)
                        m_Alpha(a) = m_Alpha(a) + 1
                    End If
                    
                    m_NumOfPixels = m_NumOfPixels + 1
                    
                End If
                
            End If
        
        Next x
        
    End If
    
    MoveYUp = m_NumOfPixels
    
End Function
