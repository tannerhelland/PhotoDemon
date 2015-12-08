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
'Last updated: 04/December/15
'Last update: initial build
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
'The initial guinea pig for this class is the Effects > Edge > Mean-Shift algorithm, because it's very straightforward,
' and it was added in 7.0 so there's no penalty if I break it irreparably!
'
'This class has several dependencies throughout PD, including pdDIB, and some specialized enums (located in PD's central
' Public_Enums_and_Types module).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal Addr As Long, ByRef dstValue As Long) As Long

'Current (x, y) position of the iterator.  Call InitializeIterator() to reset these to (0, 0).
Private m_X As Long, m_Y As Long

'Initial and final loop boundaries
Private m_InitX As Long, m_InitY As Long
Private m_FinalX As Long, m_FinalY As Long

'Shape of the current iterator, including boundary distance
Private m_WindowShape As PD_PIXEL_REGION_SHAPE
Private m_XBLeft As Long, m_XBRight As Long
Private m_YBTop As Long, m_YBBottom As Long

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

'This function is the first one you need to call.  It will initialize a bunch of internal bits against the target DIB,
' bits that are required prior to actually iterating through individual pixels.
'
'Returns: TRUE if successful.  DO NOT PROCEED with pixel iterating if the function returns FALSE.
Friend Function InitializeIterator(ByRef targetDIB As pdDIB, ByVal xRadius As Long, ByVal yRadius As Long, ByVal windowShape As PD_PIXEL_REGION_SHAPE) As Boolean

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
    m_WindowShape = PDPRS_Rectangle
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
    ' (TODO: eventually vary this by shape, obviously.)
    LockTargetHistograms = GenerateInitialWindow_Square()
    
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
    
End Sub

'Counterpart to Alias1DArray, above.  Do NOT call this function on arrays that were not originally processed by that function.
Private Sub Unalias1DArray(ByRef orig1DArray() As Long, ByRef new1DArray() As Long)
    CopyMemory ByVal VarPtrArray(new1DArray), 0&, 4&
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

'After the caller has successfully processed a pixel, they can call these functions to move to the next pixel
' in the X or Y direction.  Because this class uses serpentine scanning, the caller is responsible for changing
' direction on each Y increment.
Friend Function MoveXRight() As Long
    
    m_X = m_X + 1
    
    'Figure out Y bounds first.
    Dim yTop As Long, yBottom As Long, y As Long
    yTop = m_Y - m_YBTop
    yBottom = m_Y + m_YBBottom
    If yTop < m_InitY Then yTop = m_InitY
    If yBottom > m_FinalY Then yBottom = m_FinalY
    
    'Next, figure out X bounds.
    Dim xLeft As Long, xRight As Long, x As Long
    xLeft = (m_X - m_XBLeft) - 1
    xRight = m_X + m_XBRight
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    'If the *left* x-bound is within bounds, remove a line of pixels from the window.
    If xLeft >= 0 Then
        
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
    
    MoveXRight = m_NumOfPixels
    
End Function

Friend Function MoveYDown() As Long
    
    m_Y = m_Y + 1
    
    'Figure out X bounds first.
    Dim xLeft As Long, xRight As Long, x As Long
    xLeft = m_X - m_XBLeft
    xRight = m_X + m_XBRight
    If xLeft < m_InitX Then xLeft = m_InitX
    If xRight > m_FinalX Then xRight = m_FinalX
    
    'Next, figure out Y bounds.
    Dim yTop As Long, yBottom As Long, y As Long
    yTop = (m_Y - m_YBTop) - 1
    yBottom = m_Y + m_YBBottom
    
    Dim r As Long, g As Long, b As Long, a As Long
    
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
    
    MoveYDown = m_NumOfPixels
    
End Function

Friend Function MoveYUp() As Long

    m_Y = m_Y - 1
    
    'Figure out X bounds first.
    Dim xLeft As Long, xRight As Long, x As Long
    xLeft = m_X - m_XBLeft
    xRight = m_X + m_XBRight
    If xLeft < m_InitX Then xLeft = m_InitX
    If xRight > m_FinalX Then xRight = m_FinalX
    
    'Next, figure out Y bounds.
    Dim yTop As Long, yBottom As Long, y As Long
    yTop = m_Y - m_YBTop
    yBottom = m_Y + m_YBBottom + 1
    
    Dim r As Long, g As Long, b As Long, a As Long
    
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
    
    MoveYUp = m_NumOfPixels
    
End Function
