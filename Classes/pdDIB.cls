VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon DIB Wrapper class (formerly known as "pdLayer")
'Copyright 2012-2026 by Tanner Helland
'Created: 29/August/12
'Last updated: 24/March/25
'Last update: if we fail to create a blank surface at the target size, this class will now send out a
'             "cry for help"; this will trigger other surfaces in the app to suspend themselves to disk
'
'This class manages pixel data for almost everything inside PhotoDemon.  From image layers to UI
' elements, all are managed as pdDIB objects (a name that exists for historical reasons, but which
' may not always correspond to a literal Windows DIB object).
'
'To support legacy behavior - especially in UI elements - this class exposes a number of GDI-style
' properties.  For example, you can call the .GetDIBDC function to retrieve a GDI-compatible device
' context for use with functions like BitBlt.  Note, however, that this class may silently release
' that DC under a wide variety of circumstances, so you *cannot* cache GDI-type properties
' externally.  Instead, always call .Get-prefixed properties when you need GDI-compatible handles.
'
'This class also supports the concept of "suspending" the underlying surface to either a compressed
' memory stream or all the way down to the hard drive.  The caller doesn't need to track suspend state -
' just know that if the surface is suspended and you attempt to access it, there may be a brief pause
' as PD decompresses (or loads from disk) surface contents.  As long as you only access the underlying
' pixel data via safe wrapper functions (and do not cache return values between calls), you never have
' to worry about suspended data "not being there" - this class will silently manage all of that for you.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'For additional debug data, set this to TRUE.  (This should be set to FALSE in release builds.)
Private Const DIB_DEBUG_VERBOSE As Boolean = False

'Note that some critical structs, declares and enums are used in other ways throughout PD so they
' are not declared here.  To use this as a standalone class, you will need to do some copy+paste work
' from other PD modules.
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    xPelsPerMeter As Long
    yPelsPerMeter As Long
    ColorUsed As Long
    ColorImportant As Long
End Type

Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

'GDI APIs
Private Declare Function AlphaBlend Lib "gdi32" Alias "GdiAlphaBlend" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByRef lpBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long

'Variables related to the DIB
Private m_dibDC As Long                 'hDC for this DIB
Private m_dibHandle As Long             'Actual DIB handle for this DIB
Private m_dibHandleOriginal As Long     'Original handle when this DIB is first created (we must store this so we can properly clean up the DIB when we're finished)
Private m_dibBits As Long               'Pointer to the actual DIB pixel bits
Private m_dibHeader As BITMAPINFOHEADER 'Persistent DIB header; this will be used with WAPI to create the DIB initially

'The DIB's width and height
Private m_dibWidth As Long, m_dibHeight As Long

'The DIB's array width (m_dibWidth * 4 for 32-bit, varies for 24-bit due to DWORD-alignment)
Private m_dibStride As Long

'The DIB's color depth (should only ever be 24 or 32)
Private m_dibColorDepth As Long

'DIB pixel data is always backed against a file mapping object
' (see https://devblogs.microsoft.com/oldnewthing/20100108-00/?p=15343)
' This class manages its own handles silently, but it can also be configured to back pixel data
' against a dedicated on-disk file.
Private m_BackingFile As String, m_FileMM As pdFileMM

'Used when writing/reading DIB data to/from a file
Private Const DIB_IDENTIFIER As Long = &H726C4450  'The ASCII chars "PDlr"
Private Const DIB_FILE_VERSION_2015 As Long = &H1001&
Private Const DIB_FILE_VERSION_2016 As Long = &H1002&

'If this DIB was loaded from file, its original file format will be set here.  (By default, this value will be -1).
Private m_OriginalFIF As PD_IMAGE_FORMAT

'If this DIB has been tagged and/or hard-converted to a specific color profile, the profile's hash will
' be set, and the association will also be marked.  (Any used color profiles are stored in a central
' per-session cache, which spares us from needing to duplicate them all over the program.)
Private m_ColorProfileState As PD_ColorManagementState, m_ColorProfileHash As String

'DIBs created from loaded image files contain resolution data.  This data isn't important on a per-DIB basis, but the parent
' pdImage object will copy the resolution data from its first-loaded child DIB (if it has a resolution)
Private m_XResolution As Double, m_YResolution As Double, m_dibDPI As Double

'DIBs created from loaded image files will have their original color depth stored here.  Note that PD only works in 24/32 bpp
' mode at present, but this value may contain other bit-depths (e.g. 8bpp if the source data was a GIF).
Private m_originalColorDepth As Long

'If the alpha is currently premultiplied, this will be set to TRUE
Private m_IsAlphaPremultiplied As Boolean

'New in 9.0 is the option to "suspend" a DIB to a compressed memory stream.  The advantage of this is
' (obviously) less memory, but also freeing one or more GDI handles associated with a DIB.  Note that
' doing this repeatedly comes with obvious performance penalties, but for e.g. UI DIBs whose appearance
' doesn't change frequently, the performance savings can be large.  Also, header bits like DIB size can
' still be queried freely - only the pixel bits themselves get "suspended".
'
'(In 2025.3, suspend was updated to allow for suspending to compressed memory *or* disk; the latter can
' now be used when memory constraints are tight.)
Private m_IsSuspended As Boolean, m_SuspensionBuffer() As Byte, m_SizeSuspended As Long, m_SizeUnsuspended As Long
Private m_SuspendFormat As PD_CompressionFormat
Private m_SuspendFilename As String

'Get/set alpha premultiplication.
' IMPORTANT NOTE!  To make it explicitly clear that modifying this property DOES NOT ACTUALLY MODIFY THE IMAGE, the Set instruction is
' labeled differently.  It is only meant to be used by DIB creation functions, where the premultiplication state is explicitly known prior
' to writing DIB bits.  The counterpart SetAlphaPremultiplication function (which is found further down in this file) will actually
' modify image bits as necessary to create the desired premultiplication state.
Friend Function GetAlphaPremultiplication() As Boolean
    GetAlphaPremultiplication = m_IsAlphaPremultiplied
End Function

Friend Sub SetInitialAlphaPremultiplicationState(ByVal newState As Boolean)
    m_IsAlphaPremultiplied = newState
End Sub

'Get/set color management data
Friend Function GetColorManagementState() As PD_ColorManagementState
    GetColorManagementState = m_ColorProfileState
End Function

Friend Sub SetColorManagementState(ByVal newState As PD_ColorManagementState)
    m_ColorProfileState = newState
End Sub

Friend Function GetColorProfileHash() As String
    GetColorProfileHash = m_ColorProfileHash
End Function

Friend Sub SetColorProfileHash(ByRef newHash As String)
    m_ColorProfileHash = newHash
End Sub

'Get/set original color depth.  Note that this is set after tone-mapping and other actions have been applied
Friend Function GetOriginalColorDepth() As Long
    GetOriginalColorDepth = m_originalColorDepth
End Function

Friend Sub SetOriginalColorDepth(ByVal origColorDepth As Long)
    m_originalColorDepth = origColorDepth
End Sub

'Set the DPI of this DIB.  This is only relevant if this DIB has been created directly from an image file.
Friend Sub SetDPI(ByVal xRes As Double, ByVal yRes As Double)
    
    'Many image types do not store resolution information; default to 96 in this case
    If (xRes = 0#) Then xRes = 96#
    If (yRes = 0#) Then yRes = 96#
    
    m_XResolution = xRes
    m_YResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    m_dibDPI = (xRes + yRes) * 0.5

End Sub

'Even though we store separate x and y DPI, PD only deals in a single per-DIB DPI value
Friend Function GetDPI() As Double
    GetDPI = m_dibDPI
End Function

'Get/Set original DIB format
Friend Function GetOriginalFormat() As PD_IMAGE_FORMAT
    GetOriginalFormat = m_OriginalFIF
End Function

Friend Sub SetOriginalFormat(ByVal origFIF As PD_IMAGE_FORMAT)
    m_OriginalFIF = origFIF
End Sub

'Returns TRUE if the DIB has been suspended to either a compressed memory stream or a temp file
Friend Function IsSuspended(Optional ByVal checkSuspendedToDisk As Boolean = False) As Boolean
    IsSuspended = m_IsSuspended
    If (IsSuspended And checkSuspendedToDisk) Then IsSuspended = (LenB(m_SuspendFilename) <> 0)
End Function

'Given a pixel coordinate, return an RGBA quad for that coordinate.
' Returns: TRUE if the pixel lies inside DIB boundaries; FALSE otherwise.  Make sure to check this before using the RGBQUAD.
Friend Function GetPixelRGBQuad(ByVal x As Long, ByVal y As Long, ByRef dstQuad As RGBQuad) As Boolean
    
    'Before doing anything else, check to see if the x/y postition lies inside the DIB
    If (x >= 0) And (x < m_dibWidth) And (y >= 0) And (y < m_dibHeight) Then
        
        'The point lies inside the DIB, which means we need to figure out the color at this position
        GetPixelRGBQuad = True
        
        'Unsuspend this DIB (as necessary)
        If m_IsSuspended Then UnsuspendDIB
        
        Dim tmpData() As Byte, tSA As SafeArray2D
        Me.WrapArrayAroundDIB tmpData, tSA
        
        Dim xStride As Long
        xStride = x * (m_dibColorDepth \ 8)
        
        'Failsafe bounds check
        If ((xStride + 2) < m_dibStride) Then
        
            With dstQuad
                .Blue = tmpData(xStride, y)
                .Green = tmpData(xStride + 1, y)
                .Red = tmpData(xStride + 2, y)
                If (m_dibColorDepth = 32) Then .Alpha = tmpData(xStride + 3, y)
            End With
            
        End If
        
        Me.UnwrapArrayFromDIB tmpData
        
    'This coordinate does not lie inside the layer.
    Else
        GetPixelRGBQuad = False
    End If

End Function

'Given a pixel coordinate, set an RGBA quad to that coordinate.
' Returns: TRUE if the pixel lies inside DIB boundaries; FALSE otherwise.
Friend Function SetPixelRGBQuad(ByVal x As Long, ByVal y As Long, ByRef srcQuad As RGBQuad) As Boolean

    'Before doing anything else, check to see if the x/y postition lies inside the DIB
    If (x >= 0) And (x < m_dibWidth) And (y >= 0) And (y < m_dibHeight) Then
        
        'The point lies inside the DIB, which means we need to figure out the color at this position
        SetPixelRGBQuad = True
        
        'Unsuspend this DIB (as necessary)
        If m_IsSuspended Then UnsuspendDIB
        
        Dim tmpData() As Byte, tSA As SafeArray2D
        Me.WrapArrayAroundDIB tmpData, tSA
        
        Dim xStride As Long
        xStride = x * (m_dibColorDepth \ 8)
        
        'Failsafe bounds check
        If ((xStride + 2) < m_dibStride) Then
        
            With srcQuad
                tmpData(xStride, y) = .Blue
                tmpData(xStride + 1, y) = .Green
                tmpData(xStride + 2, y) = .Red
                If (m_dibColorDepth = 32) Then tmpData(xStride + 3, y) = .Alpha
            End With
            
        End If
        
        Me.UnwrapArrayFromDIB tmpData
        
    'This coordinate does not lie inside the layer.
    Else
        SetPixelRGBQuad = False
    End If

End Function

'Return a "thumbnail" of the DIB.  The returned DIB will always be 32-bpp, with transparent padding
' to ensure a square thumbnail size.
Friend Sub GetThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbSize As Long = 64)
    
    'Determine the thumbnail's actual width and height, and any x and y offset necessary to
    ' preserve the aspect ratio and center the image on the thumbnail.
    Dim thumbWidth As Long, thumbHeight As Long, thumbLeft As Single, thumbTop As Single
    PDMath.ConvertAspectRatio Me.GetDIBWidth, Me.GetDIBHeight, thumbSize, thumbSize, thumbWidth, thumbHeight
    
    'If the image is wider than it is tall, center the thumbnail vertically
    If (thumbWidth > thumbHeight) Then
        thumbLeft = 0!
        thumbTop = (thumbSize - thumbHeight) * 0.5
    
    '...otherwise, center it horizontally
    Else
        thumbTop = 0!
        thumbLeft = (thumbSize - thumbWidth) * 0.5
    End If
    
    'Prep the destination thumbnail
    If (dstThumbnailDIB Is Nothing) Then Set dstThumbnailDIB = New pdDIB
    If (dstThumbnailDIB.GetDIBWidth <> thumbSize) Or (dstThumbnailDIB.GetDIBHeight <> thumbSize) Or (dstThumbnailDIB.GetDIBWidth = 0) Then
        dstThumbnailDIB.CreateBlank thumbSize, thumbSize, 32, 0, 0
    Else
        dstThumbnailDIB.ResetDIB 0
    End If
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
        
    'Paint the thumbnail into place
    GDI_Plus.GDIPlus_StretchBlt dstThumbnailDIB, thumbLeft, thumbTop, thumbWidth, thumbHeight, Me, 0!, 0!, Me.GetDIBWidth, Me.GetDIBHeight, , UserPrefs.GetThumbnailInterpolationPref(), , , , True
    dstThumbnailDIB.SetInitialAlphaPremultiplicationState Me.GetAlphaPremultiplication()
    
End Sub

'Convenience functions for wrapping an array around this DIB's bits.
' You *must* call the Unwrap function prior to the array falling out of scope, or VB will crash.
Friend Sub WrapArrayAroundDIB(ByRef srcArray() As Byte, ByRef srcSafeArray As SafeArray2D)
    PrepInternalSafeArray srcSafeArray
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapArrayAroundDIB_1D(ByRef srcArray() As Byte, ByRef srcSafeArray As SafeArray1D)
    PrepInternalSafeArray_Scanline srcSafeArray, 0
    srcSafeArray.cElements = m_dibStride * m_dibHeight
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapArrayAroundScanline(ByRef srcArray() As Byte, ByRef srcSafeArray As SafeArray1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalSafeArray_Scanline srcSafeArray, dstScanLine
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub UnwrapArrayFromDIB(ByRef srcArray() As Byte)
    PutMem4 VarPtrArray(srcArray), 0&
End Sub

Friend Sub WrapLongArrayAroundDIB(ByRef srcArray() As Long, ByRef srcSafeArray As SafeArray2D)
    PrepInternalLongSafeArray srcSafeArray
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapLongArrayAroundDIB_1D(ByRef srcArray() As Long, ByRef srcSafeArray As SafeArray1D)
    PrepInternalLongSafeArray_Scanline srcSafeArray, 0
    srcSafeArray.cElements = m_dibWidth * m_dibHeight
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapLongArrayAroundScanline(ByRef srcArray() As Long, ByRef srcSafeArray As SafeArray1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalLongSafeArray_Scanline srcSafeArray, dstScanLine
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub UnwrapLongArrayFromDIB(ByRef srcArray() As Long)
    PutMem4 VarPtrArray(srcArray), 0&
End Sub

Friend Sub WrapRGBQuadArrayAroundDIB(ByRef srcArray() As RGBQuad, ByRef srcSafeArray As SafeArray2D)
    PrepInternalLongSafeArray srcSafeArray
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapRGBQuadArrayAroundDIB_1D(ByRef srcArray() As RGBQuad, ByRef srcSafeArray As SafeArray1D)
    PrepInternalLongSafeArray_Scanline srcSafeArray, 0
    srcSafeArray.cElements = m_dibWidth * m_dibHeight
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub WrapRGBQuadArrayAroundScanline(ByRef srcArray() As RGBQuad, ByRef srcSafeArray As SafeArray1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalLongSafeArray_Scanline srcSafeArray, dstScanLine
    PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
End Sub

Friend Sub UnwrapRGBQuadArrayFromDIB(ByRef srcArray() As RGBQuad)
    PutMem4 VarPtrArray(srcArray), 0&
End Sub

'Force all alpha bytes in the DIB to some preset value.  If the value is *not* 255, please remember to premultiply
' alpha accordingly, if the DIB will be rendered to the screen.
Friend Function ForceNewAlpha(ByVal newAlpha As Byte) As Boolean
    
    ForceNewAlpha = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If (m_dibColorDepth = 32) Then

        'Make sure this DIB isn't empty
        If (m_dibHandle <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            'Loop through the image and force each alpha value to the user's specified value
            Dim iData() As Byte, tmpSA As SafeArray1D
            Me.WrapArrayAroundDIB_1D iData, tmpSA
            
            Dim numPixels As Long
            numPixels = (m_dibWidth * m_dibHeight - 1) * 4
                
            'Loop through the image, checking alphas as we go
            Dim x As Long
            For x = 3 To numPixels + 3 Step 4
                iData(x) = newAlpha
            Next x
            
            Me.UnwrapArrayFromDIB iData
            ForceNewAlpha = True
            
            'If alpha is 255, mark this DIB as premultiplied
            If (newAlpha = 255) Then Me.SetInitialAlphaPremultiplicationState True
            
        Else
            Debug.Print "WARNING!  The target DIB is empty (DC or Width or Height = 0).  pdDIB.ForceNewAlpha failed."
        End If
        
    Else
        Debug.Print "WARNING!  You cannot call pdDIB.ForceNewAlpha on a 24-bpp DIB!"
    End If
    
End Function

'Translate an OLE color to an RGB Long.
' (Note that OleTranslateColor returns -1 if it fails; if that happens, default to white)
Private Function TranslateColor(ByVal colorRef As Long) As Long
    If OleTranslateColor(colorRef, 0, TranslateColor) Then TranslateColor = RGB(255, 255, 255)
End Function

Friend Sub FreeFromDC()
    If (m_dibDC <> 0) Then
        SelectObject m_dibDC, m_dibHandleOriginal
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
        m_dibHandleOriginal = 0
    End If
End Sub

'Return this DIB's color depth
Friend Function GetDIBColorDepth() As Long
    GetDIBColorDepth = m_dibColorDepth
End Function

'Return this DIB's array width
Friend Function GetDIBStride() As Long
    GetDIBStride = m_dibStride
End Function

'Return this DIB's width
Friend Function GetDIBWidth() As Long
    GetDIBWidth = m_dibWidth
End Function

'Return this DIB's height
Friend Function GetDIBHeight() As Long
    GetDIBHeight = m_dibHeight
End Function

'Return whether or not this DIB has image data associated with it.  Requires unsuspending first.
Friend Function HasImage() As Boolean
    If m_IsSuspended Then UnsuspendDIB
    HasImage = (m_dibHandle <> 0)
End Function

'Return whether the DIB is top-down (negative height) or bottom-up (positive height).  Only relevant
' when trading data with 3rd-party sources; internally, PD always assumes top-down.
Friend Function IsDIBTopDown() As Boolean
    IsDIBTopDown = (m_dibHeader.Height < 0)
End Function

'Return this DIB's hDC.  DIBs receive a DC by default, when they are first created.  However, the caller
' can remove the DIB from that DC by calling .FreeFromDC().  If you subsequently request a DC for the DIB,
' one will be auto-created for you.
Friend Function GetDIBDC() As Long
    
    GetDIBDC = m_dibDC
    If (GetDIBDC = 0) Then
    
        'Unsuspend this DIB (as necessary)
        If m_IsSuspended Then UnsuspendDIB
        
        If (m_dibHandle <> 0) Then
            m_dibDC = GDI.GetMemoryDC()
            m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
            GetDIBDC = m_dibDC
        End If
        
    End If
    
End Function

'Return a pointer to this DIB's actual DIB.  Useful primarily for GDI interop.  Requires unsuspending first.
Friend Function GetDIBHandle() As Long
    If m_IsSuspended Then UnsuspendDIB
    GetDIBHandle = m_dibHandle
End Function

'Return a pointer to this DIB's pixel data (no header).  Use dib stride and height to determine total bitmap size.
' Requires unsuspending first.
Friend Function GetDIBPointer() As Long
    If m_IsSuspended Then UnsuspendDIB
    GetDIBPointer = m_dibBits
End Function

'Return a pointer to an individual scanline.  For performance reasons, this function does not perform bounds checking -
' so use it carefully!  Requires unsuspending first.
Friend Function GetDIBScanline(ByVal targetScanline As Long) As Long
    If m_IsSuspended Then UnsuspendDIB
    GetDIBScanline = m_dibBits + (targetScanline * m_dibStride)
End Function

'Return a pointer to this DIB's header.  Does *NOT* require unsuspending first.
Friend Function GetDIBHeader() As Long
    GetDIBHeader = VarPtr(m_dibHeader)
End Function

Friend Sub CopyDIBHeader(ByRef dstHeader As BITMAPINFOHEADER)
    dstHeader = m_dibHeader
End Sub

'Quickly create a new DIB from an arbitrary DC.  No stretching is applied, by design; StretchBlt is unpredictable with
' 32-bpp data, so any stretching must be applied *after* creation (when PD has full control over the image data).
Friend Function CreateFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32, Optional ByVal isSourcePremultiplied As Boolean = False) As Boolean
    
    CreateFromDC = False
    
    'Make sure the DIB we're passed isn't empty
    If (srcDC <> 0) Then
        
        'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
        If (srcWidth = Me.GetDIBWidth) And (srcHeight = Me.GetDIBHeight) And (srcColorDepth = Me.GetDIBColorDepth) Then
            BitBlt Me.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
            If (srcColorDepth = 32) Then m_IsAlphaPremultiplied = isSourcePremultiplied Else m_IsAlphaPremultiplied = False
            CreateFromDC = True
        Else
        
            'Create a new, blank DIB the same size as the source DIB
            If Me.CreateBlank(srcWidth, srcHeight, srcColorDepth, 0, 255) Then
                BitBlt Me.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
                If (srcColorDepth = 32) Then m_IsAlphaPremultiplied = isSourcePremultiplied Else m_IsAlphaPremultiplied = False
                CreateFromDC = True
            End If
            
        End If
        
        'Our internal DC may not be immediately required; free it until requested again
        Me.FreeFromDC
        
    End If
    
End Function

'Make a copy of an existing DIB
Friend Function CreateFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal newWidth As Long = -1, Optional ByVal newHeight As Long = -1, Optional ByVal resizeModeIfNecessary As GP_InterpolationMode = GP_IM_HighQualityBicubic) As Boolean
    
    CreateFromExistingDIB = False
    
    'Make sure the DIB we're passed isn't empty
    If (srcDIB.GetDIBDC <> 0) Then
        
        'Prepare new width and height values as requested by the user
        If (newWidth < 0) Then newWidth = srcDIB.GetDIBWidth
        If (newHeight < 0) Then newHeight = srcDIB.GetDIBHeight
        
        'If the width and height values are not being changed, the transfer is simple
        If (newWidth = srcDIB.GetDIBWidth) And (newHeight = srcDIB.GetDIBHeight) Then
            
            'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
            If (newWidth = m_dibWidth) And (newHeight = m_dibHeight) And (srcDIB.GetDIBColorDepth = m_dibColorDepth) Then
                
                'Copy the source DIB's header to ensure full parity.  (Note that this only copies a BITMAPINFOHEADER,
                ' meaning descriptive info only - not things like the actual DIB bits pointer.)
                srcDIB.CopyDIBHeader m_dibHeader
                CreateFromExistingDIB = True
            
            'If our current DIB is a different size and/or bit-depth than the source image,
            ' create a new, blank DIB the same size as the source DIB
            Else
                CreateFromExistingDIB = Me.CreateBlank(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBColorDepth, 0, 0)
            End If
            
            'If creation and/or header-copy were successful, copy all image bits (BitBlt works for <=32-bpp data),
            ' then mirror current alpha premultiplication status.
            If CreateFromExistingDIB Then
                BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
            End If
        
        'If new width and height values are being specified, the transfer is a bit more complex
        Else
            
            'If our current size does not match, create a newly sized DIB in advance.
            If (Me.GetDIBWidth <> newWidth) Or (Me.GetDIBHeight <> newHeight) Or (Me.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
                Me.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth
            Else
                Me.ResetDIB 0
            End If
            
            CreateFromExistingDIB = GDI_Plus.GDIPlusResizeDIB(Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, resizeModeIfNecessary, P2_PO_Half)
                
            'Note that GDI+ *always* paints images with pre-multiplied alpha.  If our source image did *not*
            ' have premultiplied alpha, this will result in a mismatch; it's expected that the caller handles
            ' this case manually, because PD is currently designed to always assume premultiplication internally.
            
        End If
        
        'Free this DIB and the source DIB from their DCs, as they may not be required again for some time
        srcDIB.FreeFromDC
        Me.FreeFromDC
        
    End If
    
End Function

'Convert this DIB to 24bpp mode.  At present, PD only uses this under certain export circumstances;
' internally, all pdDIB objects are treated as 32-bpp to ensure alpha is available if we need it.
Friend Function ConvertTo24bpp(Optional ByVal newBackColor As Long = vbWhite) As Boolean
    
    ConvertTo24bpp = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function pointless.
    If (m_dibColorDepth = 32) Then
        
        'Unsuspend this DIB (as necessary)
        If m_IsSuspended Then UnsuspendDIB
        
        'Make sure this DIB isn't empty
        If (m_dibHandle <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.CreateFromExistingDIB Me
                        
            'Composite the temporary DIB against a white background, per convention.
            tmpDIB.CompositeBackgroundColor Colors.ExtractRed(newBackColor), Colors.ExtractGreen(newBackColor), Colors.ExtractBlue(newBackColor)
            
            'Now erase our own DIB
            Me.EraseDIB
            
            'Create a new DIB that's exactly the same size as the old one
            If Me.CreateBlank(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 24) Then
                BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                ConvertTo24bpp = True
            End If
            
            'Minimize GDI resources by freeing our DC
            Me.FreeFromDC
    
        End If
    
    'If the DIB is already 24-bpp, return TRUE as there's nothing left for us to do
    Else
        ConvertTo24bpp = True
    End If
    
End Function

'Convert this DIB to 32bpp mode
Friend Function ConvertTo32bpp(Optional ByVal newTransparency As Byte = 255) As Boolean
    
    ConvertTo32bpp = False
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    'Make sure this DIB isn't empty
    If (m_dibHandle <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) And (Me.GetDIBColorDepth <> 32) Then
        
        'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB Me
        
        'Now erase this DIB and create a new one that's exactly the same size, but 32-bit
        If Me.CreateBlank(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 32) Then
        
            'Copy the image data from the temporary DIB without modification, then free the source image
            ' (it's no longer required).
            BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
            Set tmpDIB = Nothing
            
            'Finally, we need to loop through the image and set all alpha values to 255. (BitBlt from a
            ' 24-bpp to 32-bpp target technically treats the destination alpha channel as "undefined",
            ' but in real-world practice it always forces alpha to 0.)
            Dim imgPixels() As Byte, tmpSA As SafeArray1D
            Dim x As Long, y As Long, xMax As Long
            xMax = (m_dibWidth * 4) - 1
            
            'Loop through the image, setting the alpha of each pixel to 255 (opaque)
            For y = 0 To m_dibHeight - 1
                Me.WrapArrayAroundScanline imgPixels, tmpSA, y
            For x = 3 To xMax Step 4
                imgPixels(x) = newTransparency
            Next x
            Next y
            
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            Me.UnwrapArrayFromDIB imgPixels
            
            'If the alpha value is not 255, apply premultiplication now
            If (newTransparency <> 255) Then Me.SetAlphaPremultiplication True Else Me.SetInitialAlphaPremultiplicationState True
            
            ConvertTo32bpp = True
            
        End If
        
        'Minimize GDI resources by freeing our DC
        Me.FreeFromDC
        
    End If
    
End Function

'Create a blank DIB. If no colorDepth is specified, it will default to 24bpp (16 million colors, no alpha-channel).
' If the DIB is 32bpp, an optional alpha parameter can be set for the DIB.
Friend Function CreateBlank(ByVal imgWidth As Long, ByVal imgHeight As Long, Optional ByVal colorDepth As Long = 24, Optional ByVal initialColor As Long = vbWhite, Optional ByVal initialAlpha As Long = 0) As Boolean
    
    On Error GoTo CouldNotCreateDIB
    
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush
                
    'By default, we try to avoid creating new DIBs whenever we can.  In the unlikely chance that this DIB already exists,
    ' let's compare its size and color depth to the requested dimensions - if they match, we can reuse this DIB as-is.
    Dim existingDIBIsOkay As Boolean, colorSetIsSkippable As Boolean
    existingDIBIsOkay = (imgWidth = m_dibWidth) And (imgHeight = m_dibHeight) And (colorDepth = m_dibColorDepth) And (m_dibHandle <> 0)
    
    If existingDIBIsOkay Then
        
        m_IsAlphaPremultiplied = False
        
        'This DIB is fine as-is!  Simply reset color and alpha-premultiplication to match the requested values.
        If (colorDepth = 24) Then
            If (initialColor <> vbWhite) And (initialColor <> vbBlack) Then initialColor = TranslateColor(initialColor)
            GDI.FillRectToDC Me.GetDIBDC, 0, 0, imgWidth, imgHeight, initialColor
        Else
            
            colorSetIsSkippable = (initialAlpha = 0) And (initialColor = 0)
            
            If colorSetIsSkippable Then
                FillMemory m_dibBits, m_dibStride * m_dibHeight, 0
            Else
                
                'The other "skippable" color combination is an opaque+white image
                colorSetIsSkippable = (initialAlpha = 255) And (initialColor = vbWhite)
                If colorSetIsSkippable Then
                    FillMemory m_dibBits, m_dibStride * m_dibHeight, 255
                Else
                
                    If (initialColor <> vbWhite) And (initialColor <> vbBlack) Then initialColor = TranslateColor(initialColor)
                    
                    Drawing2D.QuickCreateSurfaceFromDC cSurface, Me.GetDIBDC, False
                    cSurface.SetSurfaceCompositing P2_CM_Overwrite
                    Drawing2D.QuickCreateSolidBrush cBrush, initialColor, initialAlpha / 2.55
                    PD2D.FillRectangleF cSurface, cBrush, 0, 0, imgWidth + 1, imgHeight + 1
                    Set cSurface = Nothing
                    
                End If
                
            End If
                    
        End If
        
        'We probably created (or reused) an existing DC during the creation phase; if we did, clear it now
        Me.FreeFromDC
                
    Else
    
        'Erase any existing DIB data
        Me.EraseDIB
        
        'PhotoDemon only supports 24 and 32 BPP at present
        If (colorDepth <> 32) And (colorDepth <> 24) Then colorDepth = 32
            
        'Force the DIB to have a size of at least 1x1
        If (imgWidth < 1) Then imgWidth = 1
        If (imgHeight < 1) Then imgHeight = 1
        
        'Cache the requested color depth, width, and height; we'll refer to these constantly during DIB interactions
        m_dibColorDepth = colorDepth
        m_dibWidth = imgWidth
        m_dibHeight = imgHeight
        
        'Prepare the required header
        With m_dibHeader
            .Size = Len(m_dibHeader)
            .Planes = 1
            .BitCount = colorDepth
            .Width = imgWidth
            .Height = -imgHeight
            'As always, this value needs to be a multiple of four; with 32bpp that's automatic, with 24bpp it is not
            If (colorDepth = 32) Then
                m_dibStride = 4 * imgWidth
            Else
                m_dibStride = (imgWidth * 3 + 3) And &HFFFFFFFC
            End If
            .ImageSize = m_dibStride * imgHeight
        End With
        
        'If we don't already have a compatible DC, create one now
        If (Me.GetDIBDC = 0) Then m_dibDC = GDI.GetMemoryDC()
        
        If (m_dibDC <> 0) Then
            
            'Create a DIB using the system page file to back it
            If m_FileMM.StartNewFile(m_BackingFile, m_dibHeader.ImageSize, OptimizeNone) Then
                m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, m_FileMM.GetMMHandle(), 0)
            
            'Alternatively, you can let the system manage your backing setup for you
            Else
                m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, 0, 0)
            End If
            
            'If we failed to create the DIB, it may be a memory issue.  Send out a cry for help
            ' and see if we can cut memory usage somewhere else in the app.
            If (m_dibHandle = 0) Then
                
                PDDebug.LogAction "WARNING: failed to create requested DIB (" & m_dibWidth & "x" & m_dibHeight & ")"
                Message "Memory usage looks high; suspending some data to disk before continuing..."
                PDImages.StrategicMemoryReduction
                UIImages.FreeSharedCompressBuffer
                Set m_FileMM = New pdFileMM
                Message vbNullString
                
                'Try creating the DIB again
                If m_FileMM.StartNewFile(m_BackingFile, m_dibHeader.ImageSize, OptimizeNone) Then
                    m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, m_FileMM.GetMMHandle(), 0)
                
                'Alternatively, you can let the system manage your backing setup for you
                Else
                    m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, 0, 0)
                End If
                
            End If
            
            'If successful, select the newly created dib into our DC
            If (m_dibHandle <> 0) Then
                
                'Inside debug mode, update a global counter (for leak tracking)
                PDDebug.UpdateResourceTracker PDRT_hDIB, 1
                
                'Select the DIB into a DC so we can fill it with the requested background color
                m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
            
                'If the DIB is 24bpp, apply the background color now
                If (colorDepth = 24) Then
                    
                    If (initialColor <> 0) Then
                        
                        'The back color may or may not be a system color, so translate it just in case
                        If (initialColor <> vbWhite) Then initialColor = TranslateColor(initialColor)
                        GDI.FillRectToDC m_dibDC, 0, 0, imgWidth, imgHeight, initialColor
                        
                    End If
                        
                    m_IsAlphaPremultiplied = False
                
                '32bpp requires GDI+
                Else
                    
                    colorSetIsSkippable = (initialAlpha = 0) And (initialColor = 0)
                    
                    'Newly created DIBs will be black and transparent by default, so we can skip the fill
                    If (Not colorSetIsSkippable) Then
                        
                        'Look for opaque+white, which is also skippable
                        colorSetIsSkippable = (initialAlpha = 255) And (initialColor = vbWhite)
                        If colorSetIsSkippable Then
                            FillMemory m_dibBits, m_dibStride * m_dibHeight, 255
                        Else
                        
                            'The back color may or may not be a system color, so translate it just in case
                            If (initialColor <> vbWhite) Then initialColor = TranslateColor(initialColor)
                            
                            Drawing2D.QuickCreateSurfaceFromDC cSurface, m_dibDC, False
                            Drawing2D.QuickCreateSolidBrush cBrush, initialColor, initialAlpha / 2.55
                            PD2D.FillRectangleF cSurface, cBrush, 0, 0, imgWidth + 1, imgHeight + 1
                            Set cSurface = Nothing
                            
                        End If
                        
                    Else
                        'We could cheat and use FillMemory here, but by default, a newly initialized DIB is already
                        ' set to all-zeroes.
                        'FillMemory m_dibBits, m_dibStride * m_dibHeight, 0&
                    End If
                    
                    'Because alpha premultiplication is unknown, we'll assume it is FALSE
                    m_IsAlphaPremultiplied = False
                    
                End If
                
            'If DIB creation failed, clear out the work we've done so far
            Else
                PDDebug.LogAction "WARNING!  pdDIB failed to create a DIB at size (" & CStr(m_dibWidth) & "x" & CStr(m_dibHeight) & "x" & m_dibColorDepth & "bpp); last API error was #" & Err.LastDllError
                Me.EraseDIB
            End If
            
            'Our parent may not use this DIB right away, so free it from its DC; the DC will be auto-created when
            ' it's next requested.
            Me.FreeFromDC
            
        End If
        
    End If
    
CouldNotCreateDIB:
    
    'Return success contingent on whether we have a DIB pointer or not
    CreateBlank = (m_dibHandle <> 0)
    
End Function

'Replace a rectangle of the DIB with a fixed RGBA value.
' IMPORTANTLY: it is up to the caller to handle alpha premultiplication on the passed color + alpha value,
' based on the underlying premultiplied alpha state of the target DIB.
Friend Sub FillRectWithColor(ByRef dstRectF As RectF, Optional ByVal fillColor As Long = vbBlack, Optional ByVal fillAlpha As Single = 0!)

    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
        
    'If the DIB is 24bpp, apply the color using GDI
    If (m_dibColorDepth = 24) Then
        
        GDI.FillRectToDC Me.GetDIBDC, Int(dstRectF.Left), Int(dstRectF.Top), Int(dstRectF.Width) + 1, Int(dstRectF.Height) + 1, fillColor
        
        'Alpha premultiplication technically doesn't matter in 24-bpp mode, so reset it to its default value (FALSE)
        m_IsAlphaPremultiplied = False
        
    '32-bpp requires GDI+
    Else
        
        Dim fillAlphaL As Long
        fillAlphaL = Int(fillAlpha * 2.55! + 0.5!)
        
        Dim tmpSurface As Long, tmpBrush As Long
        tmpSurface = GDI_Plus.GetGDIPlusGraphicsFromDC(Me.GetDIBDC)
        tmpBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(fillColor, fillAlphaL)
        
        GDI_Plus.GDIPlus_GraphicsSetCompositingMode tmpSurface, GP_CM_SourceCopy
        GDI_Plus.GDIPlus_FillRectI tmpSurface, tmpBrush, Int(dstRectF.Left), Int(dstRectF.Top), Int(dstRectF.Width + 0.5!), Int(dstRectF.Height + 0.5!)
        GDI_Plus.ReleaseGDIPlusBrush tmpBrush
        GDI_Plus.ReleaseGDIPlusGraphics tmpSurface
        
        m_IsAlphaPremultiplied = (fillAlphaL = 255) Or ((fillAlphaL = 0) And (fillColor = 0))
        
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub

'Replace the entire DIB's contents with a fixed RGBA value
Friend Sub FillWithColor(Optional ByVal fillColor As Long = vbBlack, Optional ByVal fillAlpha As Single = 0!)

    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
        
    'If the DIB is 24bpp, apply the background color now
    If (m_dibColorDepth = 24) Then
        
        'There are two cases where we can "cheat" and simply fill the entire DIB area with a single value:
        ' filling with black (everything gets set to 0), and filling with white (everything gets set to 255).
        If (fillColor = vbBlack) Then
            Me.ResetDIB 0
        Else
            If (fillColor = vbWhite) Then
                Me.ResetDIB 255
            Else
                GDI.FillRectToDC Me.GetDIBDC, 0, 0, m_dibWidth + 1, m_dibHeight + 1, fillColor
            End If
        End If
        
        'Alpha premultiplication technically doesn't matter in 24-bpp mode, so reset it to its default value (FALSE)
        m_IsAlphaPremultiplied = False
        
    '32-bpp requires GDI+
    Else
        
        'Just like the 24-bpp block above, we can cheat with black and white values in 32-bpp mode too,
        ' BUT ONLY if the alpha byte value matches.
        Dim fillAlphaL As Long
        fillAlphaL = Int(fillAlpha * 2.55! + 0.5!)
        
        If ((fillColor = vbWhite) And (fillAlphaL = 255)) Then
            Me.ResetDIB 255
            m_IsAlphaPremultiplied = True
        ElseIf ((fillColor = vbBlack) And (fillAlphaL = 0)) Then
            Me.ResetDIB 0
            m_IsAlphaPremultiplied = True
        Else
            
            Dim tmpSurface As Long, tmpBrush As Long
            tmpSurface = GDI_Plus.GetGDIPlusGraphicsFromDC(Me.GetDIBDC)
            tmpBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(fillColor, fillAlpha * 2.55)
            
            GDI_Plus.GDIPlus_GraphicsSetCompositingMode tmpSurface, GP_CM_SourceCopy
            GDI_Plus.GDIPlus_FillRectI tmpSurface, tmpBrush, -1, -1, m_dibWidth + 2, m_dibHeight + 2
            GDI_Plus.ReleaseGDIPlusBrush tmpBrush
            GDI_Plus.ReleaseGDIPlusGraphics tmpSurface
            
            m_IsAlphaPremultiplied = (fillAlphaL = 255)
            
        End If
        
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub

'This will effectively reset everything related to this DIB, including all image data. Use cautiously!
Friend Sub EraseDIB(Optional ByVal alsoReleaseDC As Boolean = False)
    
    'If we have image data, clear it out
    If (m_dibHandle <> 0) Then
        
        If (m_dibDC <> 0) Then SelectObject m_dibDC, m_dibHandleOriginal
        DeleteObject m_dibHandle
        m_dibHandle = 0
        
        'Inside debug mode, update a global counter (for leak tracking)
        PDDebug.UpdateResourceTracker PDRT_hDIB, -1
        
    End If
    
    If (alsoReleaseDC And (m_dibDC <> 0)) Then
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
    End If
    
    'Reset all associated DIB section variables
    m_dibHandleOriginal = 0
    m_dibHandle = 0
    m_dibBits = 0

    'Reset DIB size
    m_dibWidth = 0
    m_dibHeight = 0
    
    'Reset alpha premultiplication
    m_IsAlphaPremultiplied = False
    
    'Release any suspension-related values
    m_SizeSuspended = 0
    m_SizeUnsuspended = 0
    Erase m_SuspensionBuffer
    m_IsSuspended = False
    If (LenB(m_SuspendFilename) <> 0) Then
        Files.FileDeleteIfExists m_SuspendFilename
        m_SuspendFilename = vbNullString
    End If
    
    'Finally, release the memory-mapped file backing the DIB, if one exists
    If (Not m_FileMM Is Nothing) Then m_FileMM.CloseFile
    
End Sub

'Maintain the DIB's dimensions, but reset all bytes to some value (typically 0).  FillMemory is used for maximum performance.
Friend Sub ResetDIB(Optional ByVal fillValue As Byte = 0)

    'Retrieve a pointer to the DIB array and the array's current size
    Dim tmpDIBPointer As Long, tmpDIBSize As Long
    RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBSize
    
    'Erase everything
    If (tmpDIBPointer <> 0) And (tmpDIBSize > 0) Then
        If (fillValue = 0) Then ZeroMemory tmpDIBPointer, tmpDIBSize Else FillMemory tmpDIBPointer, tmpDIBSize, fillValue
    End If
    
End Sub

'"Suspend" this DIB by converting it to a compressed memory stream and freeing all associated GDI handles.
Friend Sub SuspendDIB(Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal autoKeepIfLarge As Boolean = True, Optional ByVal suspendToDisk As Boolean = False)
    
    'If this DIB doesn't exist (or if it's already suspended), quit now
    If m_IsSuspended Then
        
        'If the caller wants us suspended to disk but we are *not*, go ahead and dump our compressed data to disk.
        If (suspendToDisk And (LenB(m_SuspendFilename) = 0)) Then
            If DIB_DEBUG_VERBOSE Then PDDebug.LogAction "DIB is already suspended, but not to disk; writing out to file now..."
            m_SuspendFilename = OS.UniqueTempFilename()
            If Files.FileCreateFromByteArray(m_SuspensionBuffer, m_SuspendFilename, True, True, m_SizeSuspended) Then
                Erase m_SuspensionBuffer
            Else
                m_SuspendFilename = vbNullString
            End If
        End If
        
        Exit Sub
        
    End If
    
    If (m_dibHandle = 0) Or (m_dibBits = 0) Then Exit Sub
    
    'Very large DIBs (like the main canvas backbuffer) are expensive to compress/decompress on-the-fly,
    ' even with a very fast engine like lz4.  (It doesn't help that we have to reallocate the memory at
    ' decompression time, or that we have to drop into APIs to move the data between buffers.)
    
    'As such, PD's various resource minimizers are set to automatically ignore suspend requests for
    ' DIBs larger than some arbitrary size (currently 128kb).  This prevents noticeable UI lags.
    ' That said, the caller is allowed to ignore this if they desperately need to cut memory.
    If autoKeepIfLarge And (m_dibHeader.ImageSize > 128000) Then Exit Sub
    
    'We need to store original and compressed sizes
    m_SizeUnsuspended = m_dibHeader.ImageSize
    m_SizeSuspended = m_SizeUnsuspended
    
    'Because of quirks with the order that some UI elements get loaded, PD may attempt to suspend
    ' some DIBs before various compression libraries are made available.  If this happens, we'll
    ' still perform a suspend, but we'll do it *without* compression.  (This saves on GDI objects
    ' and user-space handles, but does nothing for memory reduction.)
    m_SuspendFormat = cmpFormat
    If (Not Compression.IsFormatSupported(m_SuspendFormat)) Then m_SuspendFormat = cf_None
    
    'Request access to the shared UI compression buffer, then compress our pixel data
    Dim dstBufferPtr As Long, dstBufferSize As Long
    If suspendToDisk Then
        
        m_SuspendFilename = OS.UniqueTempFilename()
        
        Dim worstCaseSize As Long
        worstCaseSize = Compression.GetWorstCaseSize(m_SizeUnsuspended, m_SuspendFormat)
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, m_SuspendFilename, worstCaseSize, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'We'll try to compress the data when writing it to disk, but if we *can't*, we'll do an uncompressed
            ' dump instead.  (Compression requires a temporary buffer of sub-optimal size, while an uncompressed
            ' dump can be done scanline-by-scanline, using less RAM but more disk space.)
            dstBufferPtr = cStream.Peek_PointerOnly(0, worstCaseSize)
            If (dstBufferPtr = 0) Then
                
                'Shit, memory is too tight to allocate a compression buffer.
                ' Instead, let's dump the image to file a single line at a time with *no* compression.
                ' (This is the lowest-case memory usage scenario, and when done, we can release the full
                '  backing surface to reclaim a *lot* of memory.)
                
                'Start by cleaning up the existing file we tried to create
                cStream.StopStream True
                Files.FileDeleteIfExists m_SuspendFilename
                
                'Switch the suspend format to "uncompressed"
                m_SuspendFormat = cf_None
                
                'Start a file-backed (NON-memory-mapped) stream on the target file
                If cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadWrite, m_SuspendFilename, optimizeAccess:=OptimizeSequentialAccess) Then
                    
                    'Dump scanlines one-at-a-time.  (Vertical order is fine because PD uses upside-down DIBs.)
                    Dim y As Long
                    For y = 0 To Me.GetDIBHeight - 1
                        cStream.WriteBytesFromPointer m_dibBits + (y * Me.GetDIBStride), Me.GetDIBStride
                    Next y
                    
                    'End the stream (trimming to size is automatic in non-memory-mapped files)
                    cStream.StopStream True
                    PDDebug.LogAction "Suspended DIB WITHOUT compression to file " & m_SuspendFilename & ", " & m_SizeUnsuspended & " bytes saved"
                    
                    'If the suspension buffer was used previously, it may still be allocated; free it now
                    Erase m_SuspensionBuffer
                    
                '/If we can't even access the target file, we're truly screwed.  CRASH INCOMING
                Else
                    PDDebug.LogAction "Failed to open handle to suspend file - this session is going dooooown!"
                End If
                
            '/Allocation okay - compress away!
            Else
                
                dstBufferSize = worstCaseSize
                If Compression.CompressPtrToPtr(dstBufferPtr, dstBufferSize, m_dibBits, m_SizeUnsuspended, m_SuspendFormat) Then
                    
                    'Shrink the file size to match the actual compressed bytes, then release to disk
                    m_SizeSuspended = dstBufferSize
                    cStream.SetSizeExternally m_SizeSuspended
                    cStream.StopStream True
                    If DIB_DEBUG_VERBOSE Then PDDebug.LogAction "Suspended DIB to file " & m_SuspendFilename & ", " & m_SizeUnsuspended & " bytes saved"
                    
                    'If the suspension buffer was used previously, it may still be allocated; free it now
                    Erase m_SuspensionBuffer
                    
                'If compression failed, attempt a memory suspend instead
                Else
                    cStream.StopStream True
                    Files.FileDeleteIfExists m_SuspendFilename
                    m_SuspendFilename = vbNullString
                    suspendToDisk = False
                End If
                
            End If
            
        'Something went wrong with disk access; attempt a memory suspend instead
        Else
            m_SuspendFilename = vbNullString
            suspendToDisk = False
        End If
        
        Set cStream = Nothing
        
    End If
    
    If (Not suspendToDisk) Then
        
        m_SuspendFilename = vbNullString
        dstBufferPtr = UIImages.GetSharedCompressBuffer(dstBufferSize, m_SizeUnsuspended, m_SuspendFormat)
        If Compression.CompressPtrToPtr(dstBufferPtr, dstBufferSize, m_dibBits, m_SizeUnsuspended, m_SuspendFormat) Then
            
            'Create a precisely sized local buffer to hold the compressed data, then copy it over
            m_SizeSuspended = dstBufferSize
            If VBHacks.IsArrayInitialized(m_SuspensionBuffer) Then
                If (m_SizeSuspended > UBound(m_SuspensionBuffer) + 1) Then ReDim m_SuspensionBuffer(0 To m_SizeSuspended - 1) As Byte
            Else
                ReDim m_SuspensionBuffer(0 To m_SizeSuspended - 1) As Byte
            End If
            CopyMemoryStrict VarPtr(m_SuspensionBuffer(0)), dstBufferPtr, m_SizeSuspended
            
            'Suspend info is actually a little *too* verbose for me!
            'If DIB_DEBUG_VERBOSE Then PDDebug.LogAction "Suspended DIB to memory, " & (m_SizeUnsuspended - m_SizeSuspended) & " bytes saved"
            
        'Something went wrong with compression.  Abandon ship!
        Else
            m_SizeUnsuspended = 0
            m_SizeSuspended = 0
            Exit Sub
        End If
        
    End If
    
    '*Selectively* free only our local pixel bits
    If (m_dibDC <> 0) Then
        SelectObject m_dibDC, m_dibHandleOriginal
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
        m_dibHandleOriginal = 0
    End If
    
    DeleteObject m_dibHandle
    m_dibHandle = 0
    m_dibBits = 0
    
    'Inside debug mode, update a global counter (for leak tracking)
    PDDebug.UpdateResourceTracker PDRT_hDIB, -1
    
    'Free the underlying memory-mapped file instance
    If (Not m_FileMM Is Nothing) Then m_FileMM.CloseFile
    
    'This DIB has been "suspended"
    m_IsSuspended = True
    
End Sub

'Allow some outside caller to assume ownership of our DIB handle.
' IMPORTANT NOTE: the caller is obviously responsible of freeing the DIB after calling this function.
Friend Sub TransferDIBOwnership(ByRef dstDIB As Long, ByRef dstPtr As Long)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    'Make sure we're not selected into a DC
    Me.FreeFromDC
    
    'Transfer out the key DIB identifiers
    dstDIB = m_dibHandle
    dstPtr = m_dibBits
    
    'Treat our class as if everything has been freed - the DIB is the caller's problem now
    m_dibHandle = 0
    m_dibHandleOriginal = 0
    m_dibBits = 0
    
    m_dibWidth = 0
    m_dibHeight = 0
    
    m_IsAlphaPremultiplied = False
    
End Sub

'Take a suspended DIB, and restore it to a normal GDI DIB object.
Private Sub UnsuspendDIB()
    
    'Failsafe only
    If (Not m_IsSuspended) Then Exit Sub
    
    'At shutdown, unloading toolbars may cause focus changes that trigger redraws; these are unnecessary
    ' and we want to exit as fast as possible, so ignore unsuspend requests (cleanup will happen automatically
    ' at class release)
    If g_ProgramShuttingDown Then
        Me.EraseDIB
        Exit Sub
    End If
    
    'Make sure we have suspended data
    If (m_dibWidth = 0) Or (m_dibHeight = 0) Or (m_dibBits <> 0) Or (m_SizeSuspended = 0) Or (m_SizeUnsuspended = 0) Or (m_dibHeader.Size <> Len(m_dibHeader)) Then Exit Sub
    
    'The existing DIB header can (and should) be used as-is, by design.
    ' Suspending a DIB only clears its pixel bits - *not* any of its other state trackers.
    
    'If we don't already have a compatible DC, create one now
    If (m_dibDC = 0) Then m_dibDC = GDI.GetMemoryDC()
    If (m_dibDC <> 0) Then
        
        'Create a DIB using the system page file to back it
        If m_FileMM.StartNewFile(m_BackingFile, m_dibHeader.ImageSize, OptimizeNone) Then
            m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, m_FileMM.GetMMHandle(), 0)
        
        'Alternatively, you can let the system manage your backing setup for you
        Else
            m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, 0, 0)
        End If
        
        'TODO: check failure, and force memory reduction elsewhere as necessary (then try again)
        
        'If successful, select the newly created dib into our DC
        If (m_dibHandle <> 0) Then
            
            'Inside debug mode, update a global counter (for leak tracking)
            PDDebug.UpdateResourceTracker PDRT_hDIB, 1
            
            'Compression could have occurred to disk or memory
            If (LenB(m_SuspendFilename) <> 0) Then
                
                If DIB_DEBUG_VERBOSE Then PDDebug.LogAction "Unsuspending DIB from file " & m_SuspendFilename & ", " & m_SizeUnsuspended & " bytes expected"
                
                Dim cStream As pdStream
                Set cStream = New pdStream
                If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SuspendFilename, m_SizeSuspended, optimizeAccess:=OptimizeSequentialAccess) Then
                    
                    Dim ptrImageData As Long
                    ptrImageData = cStream.Peek_PointerOnly(0, m_SizeSuspended)
                    
                    'Dump the data directly into the zip!
                    If Compression.DecompressPtrToPtr(m_dibBits, m_SizeUnsuspended, ptrImageData, m_SizeSuspended, m_SuspendFormat) Then
                        
                        cStream.StopStream True
                        Files.FileDeleteIfExists m_SuspendFilename
                        
                        m_SuspendFilename = vbNullString
                        m_SizeSuspended = 0
                        m_SizeUnsuspended = 0
                        If DIB_DEBUG_VERBOSE Then PDDebug.LogAction "DIB unsuspended successfully!"
                        
                    Else
                        PDDebug.LogAction "WARNING!  pdDIB couldn't decompress DIB from disk."
                    End If
                    
                Else
                    PDDebug.LogAction "WARNING!  pdDIB couldn't unsuspend DIB from disk."
                End If
                
                Set cStream = Nothing
                
            Else
                
                'Decompress our local copy of pixel data into the newly created DIB
                If Compression.DecompressPtrToPtr(m_dibBits, m_SizeUnsuspended, VarPtr(m_SuspensionBuffer(0)), m_SizeSuspended, m_SuspendFormat) Then
                    
                    'A lot of different strategies are possible here, but for now, let's keep the suspend
                    ' buffer intact for small images.  These are likely to suspend/unsuspend many times
                    ' as part of UI rendering, and the cost of keeping the suspend buffer intact is less
                    ' than the cost of constant small re-allocations.
                    If (m_SizeSuspended > 16384) Then Erase m_SuspensionBuffer
                    
                    m_SizeSuspended = 0
                    m_SizeUnsuspended = 0
                    
                Else
                    PDDebug.LogAction "WARNING!  pdDIB couldn't unsuspend DIB from memory."
                End If
                
            End If
            
            'We will later use m_dibHandleOriginal to clear up the memory associated with this DIB
            m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
        
        'If DIB creation failed, clear out the work we've done so far
        Else
            PDDebug.LogAction "WARNING!  pdDIB.UnsuspendDIB failed to create a DIB at size (" & CStr(m_dibWidth) & "x" & CStr(m_dibHeight) & ":" & m_dibColorDepth & "); last API error was #" & Err.LastDllError
        End If
        
        'Our parent may not use this DIB right away, so free it from its DC; the DC will be auto-created when
        ' it's next requested.
        Me.FreeFromDC
        
    End If
    
    'This DIB has been restored to its original state!
    m_IsSuspended = False
    
End Sub

'INITIALIZE class
Private Sub Class_Initialize()

    'Reset all associated DIB section variables
    m_dibDC = 0
    m_dibHandle = 0
    m_dibHandleOriginal = 0
    m_dibBits = 0
    
    'Reset DIB size
    m_dibWidth = 0
    m_dibHeight = 0
    
    'Reset color management data
    m_ColorProfileState = cms_NoManagement
    m_ColorProfileHash = vbNullString
    
    'Create a memory-mapped file manager, and explicitly note that the underlying pixel data
    ' is *not* backed by a file (by default)
    Set m_FileMM = New pdFileMM
    m_BackingFile = vbNullString
    
End Sub

'TERMINATE class
Private Sub Class_Terminate()
    Me.EraseDIB True
End Sub

'Render this DIB to some other DIB, with automatic center- and fit- behavior.
' Do NOT supply a null DIB as the target; if you do, this function will crash.
' If the targetDIB is a backbuffer designed for screen display, enable color management as relevant.
' (Similarly, this function will auto-paint a checkerboard background if desired, but if this is suspended,
' the caller must handle any backdrop preparation.)
Friend Sub RenderToDIB(ByRef dstDIB As pdDIB, Optional ByVal colorManagementMatters As Boolean = True, Optional ByVal doNotStretchIfSmaller As Boolean = False, Optional ByVal suspendTransparencyGrid As Boolean = False)
    
    'Calculate destination width/height
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstDIB.GetDIBWidth
    dstHeight = dstDIB.GetDIBHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = m_dibWidth
    srcHeight = m_dibHeight
    
    'If the caller expects the source image to be small, they may prevent us from enlarging the image to fit
    Dim fitPrevented As Boolean
    fitPrevented = False
    
    If doNotStretchIfSmaller Then
        If (srcWidth < dstWidth) And (srcHeight < dstHeight) Then
            fitPrevented = True
            dstWidth = srcWidth
            dstHeight = srcHeight
        End If
    End If
    
    'Calculate the aspect ratio of this DIB and the target picture box, then calculate a proper fit-size
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0#) Then srcAspect = srcWidth / srcHeight Else srcAspect = 1#
    If (dstHeight > 0#) Then dstAspect = dstWidth / dstHeight Else dstAspect = 1#
    
    Dim newWidth As Long, newHeight As Long, previewX As Long, previewY As Long
    PDMath.ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
    
    If fitPrevented Then
        previewX = (dstDIB.GetDIBWidth - srcWidth) \ 2
        previewY = (dstDIB.GetDIBHeight - srcHeight) \ 2
    Else
        If (srcAspect > dstAspect) Then
            previewY = (dstHeight - newHeight) \ 2
            previewX = 0
        Else
            previewX = (dstWidth - newWidth) \ 2
            previewY = 0
        End If
    End If
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    'We now branch according to active color management.  (Whether to color-manage the image is left to the caller;
    ' only they have enough information to know if it's relevant for this op)
    If colorManagementMatters Then
        
        'Create a temporary DIB to house the stretched image.  Because the stretched image is typically smaller than
        ' the original image, this reduces color management time.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank newWidth, newHeight, m_dibColorDepth, 0, 0
        tmpDIB.SetInitialAlphaPremultiplicationState False
        GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, newWidth, newHeight, Me, 0, 0, m_dibWidth, m_dibHeight
        
        'Apply display color management now
        ColorManagement.ApplyDisplayColorManagement tmpDIB
        
        'Render the now-color-managed temporary image to the destination DIB
        If (Not suspendTransparencyGrid) Then GDI_Plus.GDIPlusFillDIBRect_Pattern dstDIB, previewX, previewY, newWidth, newHeight, g_CheckerboardPattern, fixBoundaryPainting:=True
        GDI_Plus.GDIPlus_StretchBlt dstDIB, previewX, previewY, newWidth, newHeight, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight
        
    'If color management is *not* required, we don't need a temporary DIB - just render directly onto the target.
    Else
        If (Not suspendTransparencyGrid) Then GDI_Plus.GDIPlusFillDIBRect_Pattern dstDIB, previewX, previewY, newWidth, newHeight, g_CheckerboardPattern, fixBoundaryPainting:=True
        GDI_Plus.GDIPlus_StretchBlt dstDIB, previewX, previewY, newWidth, newHeight, Me, 0, 0, m_dibWidth, m_dibHeight
    End If
    
    'Minimize GDI resources by freeing our internal DC
    Me.FreeFromDC
    
End Sub

'Load a DIB's header and image information from file.  If the file is compressed, it will automatically be decompressed.
' Optionally, if the DIB data is embedded inside another file, that handle can be passed to the this function, and it will attempt to
' load DIB data from the current file pointer location.  (In this case, just supply a blank filename.)
Friend Function CreateFromFile(ByRef srcFilename As String, Optional ByVal useEmbeddedLocation As Boolean = False, Optional ByVal hFile As Long = 0) As Boolean
    
    'pdFSO is used so we can support Unicode paths
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If this is not an embedded file, create and open it now
    If (Not useEmbeddedLocation) Then
        If (Not cFile.FileCreateHandle(srcFilename, hFile, True, False, OptimizeSequentialAccess)) Then
            PDDebug.LogAction "WARNING!  pdDIB.createFromFile failed to generate a handle for " & srcFilename & ".  Load abandoned."
            CreateFromFile = False
            Exit Function
        End If
    End If
    
    'Check to make sure this file actually contains DIB data
    Dim DIBIDCheck As Long
    cFile.FileReadData hFile, VarPtr(DIBIDCheck), 4&
        
    If (DIBIDCheck <> DIB_IDENTIFIER) Then
        PDDebug.LogAction "Failed to load embedded image; invalid header found."
        CreateFromFile = False
        If (Not useEmbeddedLocation) Then cFile.FileCloseHandle hFile
        Exit Function
    End If
    
    'In the future, we may need to branch due to version differences, but for now, we can make any adjustments in-line
    Dim DIBVersionCheck As Long
    cFile.FileReadData hFile, VarPtr(DIBVersionCheck), 4&
    
    'Retrieve color depth
    Dim fColorDepth As Long
    cFile.FileReadData hFile, VarPtr(fColorDepth), 4&
    
    'The newest, 2015 version of PD DIBs retrieves alpha premultiplication as well
    Dim localAlphaPremultiplied As Boolean
    If (DIBVersionCheck = DIB_FILE_VERSION_2016) Then
        Dim tmpAlphaCheck As Long
        cFile.FileReadData hFile, VarPtr(tmpAlphaCheck), 4&
        localAlphaPremultiplied = (tmpAlphaCheck <> 0)
        
    ElseIf (DIBVersionCheck = DIB_FILE_VERSION_2015) Then
        cFile.FileReadData hFile, VarPtr(localAlphaPremultiplied), 2&
        
    'If alpha premultiplication state is not stored, make a best guess based on color depth
    Else
        localAlphaPremultiplied = Not (fColorDepth = 24)
    End If
    
    'Retrieve DIB dimensions
    Dim fWidth As Long, fHeight As Long, fArrayWidth As Long
    cFile.FileReadData hFile, VarPtr(fWidth), 4&
    cFile.FileReadData hFile, VarPtr(fHeight), 4&
    cFile.FileReadData hFile, VarPtr(fArrayWidth), 4&
    
    'Make sure the dimensions are valid
    If (fWidth = 0) Or (fHeight = 0) Or (fArrayWidth = 0) Then
        PDDebug.LogAction "WARNING!  pdDIB.CreateFromFile encountered invalid width/height measurements (" & fWidth & "x" & fHeight & ", scanline width: " & fArrayWidth & ")"
    End If
    
    'Check compression.  As of 2016, this is no longer a true/false value - it is an enum describing a specific
    ' PD compression engine.
    Dim useDecompressionFormat As PD_CompressionFormat
    useDecompressionFormat = cf_None
    
    If (DIBVersionCheck = DIB_FILE_VERSION_2016) Then
        cFile.FileReadData hFile, VarPtr(useDecompressionFormat), 4&
    Else
        Dim toDecompress As Boolean
        cFile.FileReadData hFile, VarPtr(toDecompress), 2&
        If toDecompress Then useDecompressionFormat = cf_Zlib Else useDecompressionFormat = cf_None
    End If
    
    Dim fileData() As Byte
    
    'If the data *is* compressed, we need to retrieve two additional pieces of information: the original and compressed size
    ' of the entire pixel chunk.
    Dim origDataSize As Long, compressedDataSize As Long
    
    If (useDecompressionFormat <> cf_None) Then
        cFile.FileReadData hFile, VarPtr(origDataSize), 4&
        cFile.FileReadData hFile, VarPtr(compressedDataSize), 4&
        ReDim fileData(0 To compressedDataSize - 1) As Byte
        cFile.FileReadData hFile, VarPtr(fileData(0)), compressedDataSize
    End If
    
    'Note: if the pixel data is uncompressed, we can load it directly into the DIB itself, avoiding the need for
    ' a temporary copy, which is why fileData() is initialized conditionally.
    
    'Initialize a blank DIB at the requested size.  Besides giving us a memory target, this will also initialize all module-level
    ' parameters to match the file's settings.
    If Me.CreateBlank(fWidth, fHeight, fColorDepth) Then
        
        'Now it's time to overwrite the DIB memory region with the pixel data in the file.
        Dim tmpDIBPointer As Long, tmpDIBLength As Long
        Me.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
        
        'If the DIB data is compressed, decompress it now.  Note that we decompress directly into the DIB itself.
        If (useDecompressionFormat <> cf_None) Then
            
            'If the data is decompressed successfully, convert it to a 2D array and copy it into this class's DIB memory chunk
            If (Not Compression.DecompressPtrToPtr(tmpDIBPointer, tmpDIBLength, VarPtr(fileData(0)), compressedDataSize, useDecompressionFormat)) Then
                PDDebug.LogAction "Decompression was not successful (unspecified compression plugin error).  DIB processing abandoned."
                CreateFromFile = False
                Exit Function
            End If
        
        'If the DIB data is not compressed, copy it straight into the DIB
        Else
            cFile.FileReadData hFile, tmpDIBPointer, tmpDIBLength
        End If
        
        CreateFromFile = True
        
    Else
        PDDebug.LogAction "WARNING!  pdDIB.createFromFile failed to create a new DIB using the stored file settings.  Load abandoned."
        CreateFromFile = False
    End If
    
    'Reset alpha premultiplication to match the value embedded in the file
    m_IsAlphaPremultiplied = localAlphaPremultiplied
    
    'If this is an embedded file, leave the file open for further processing.  Otherwise, close it.
    If (Not useEmbeddedLocation) Then cFile.FileCloseHandle hFile
    
End Function

'Write this class's DIB information to file.  Compression can optionally be applied to the DIB data.
' Also, if the calling function wants to embed the DIB's data inside an existing file,
' specify the embedInFile parameter and pass a valid file handle, with the pointer already moved to
' the location desired for embedding.
Friend Function WriteToFile(ByRef dstFilename As String, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal embedInFile As Boolean = False, Optional ByRef hFile As Long = 0) As Boolean
    
    'Before proceeding, make sure this DIB actually exists
    If (Me.GetDIBWidth <= 0) Or (Me.GetDIBHeight <= 0) Then
        PDDebug.LogAction "WARNING!  pdDIB.WriteToFile was called on an empty DIB.  Fix this!"
        WriteToFile = False
        Exit Function
    End If
    
    'File interactions are handled via pdFSO, so we can easily support Unicode paths
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If we are not embedding the data inside an existing file, create a new file now.
    If (Not embedInFile) Then
        If cFile.FileExists(dstFilename) Then cFile.FileDelete dstFilename
        If (Not cFile.FileCreateHandle(dstFilename, hFile, True, True, OptimizeSequentialAccess)) Then
            PDDebug.LogAction "WARNING!  pdDIB.WriteToFile failed to create a file handle for " & dstFilename & ".  Write abandoned."
            WriteToFile = False
            Exit Function
        End If
    End If
    
    'Write identifiers first
    Dim tmpLong As Long
    
    tmpLong = DIB_IDENTIFIER
    cFile.FileWriteData hFile, VarPtr(tmpLong), 4&
    
    tmpLong = DIB_FILE_VERSION_2016
    cFile.FileWriteData hFile, VarPtr(tmpLong), 4&
    
    'Color depth
    cFile.FileWriteData hFile, VarPtr(Me.GetDIBColorDepth), 4&
    
    'Alpha premultiplication (yes, this is stored even for 24-bit images)
    If m_IsAlphaPremultiplied Then tmpLong = 1 Else tmpLong = 0
    cFile.FileWriteData hFile, VarPtr(tmpLong), 4&
    
    'Size
    cFile.FileWriteData hFile, VarPtr(Me.GetDIBWidth), 4&
    cFile.FileWriteData hFile, VarPtr(Me.GetDIBHeight), 4&
    cFile.FileWriteData hFile, VarPtr(Me.GetDIBStride), 4&
    
    'And finally, the pixel data.  Start by retrieving two key items of business:
    ' - a pointer to this DIB's raw pixel data, and...
    ' - the total length of that pixel data.
    ' (Note that this call will also unsuspend the DIB, as necessary)
    Dim tmpDIBPointer As Long, tmpDIBLength As Long
    Me.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
        
    'If compression has been requested, and zLib is available, we'll shrink the DIB data down before writing it.
    Dim dataWasCompressed As Boolean: dataWasCompressed = False
    If (cmpFormat <> cf_None) Then
        
        'Before writing anything to file, attempt compression and make sure it works
        Dim compressedImageData() As Byte, compressedSize As Long
        dataWasCompressed = Compression.CompressPtrToDstArray(compressedImageData, compressedSize, tmpDIBPointer, tmpDIBLength, cmpFormat)
        
        If dataWasCompressed Then
            
            cFile.FileWriteData hFile, VarPtr(cmpFormat), 4&
                
            'Write the original size of the array, followed by the compressed size, followed by the image data itself.  (The array sizes will
            ' help us automatically create intelligently-sized buffers at decompression time.)
            cFile.FileWriteData hFile, VarPtr(tmpDIBLength), 4&
            cFile.FileWriteData hFile, VarPtr(compressedSize), 4&
            cFile.FileWriteData hFile, VarPtr(compressedImageData(0)), compressedSize
            
            'pdDebug.LogAction "pdDIB.WriteToFile: (" & CStr(Me.GetDIBWidth) & "x" & CStr(Me.GetDIBHeight) & "x" & CStr(Me.GetDIBColorDepth) & "): " & CStr(compressedSize) & " bytes of pixel data (compressed; original was " & CStr(tmpDIBLength) & " bytes)"
                
        'On failure, we'll simply write out uncompressed bytes
        Else
            cmpFormat = cf_None
        End If
        
    End If
    
    'If compression has not been requested, or zLib is unavailable, we can dump the raw image data to file as-is.
    If (Not dataWasCompressed) Then
        cFile.FileWriteData hFile, VarPtr(cmpFormat), 4&
        cFile.FileWriteData hFile, tmpDIBPointer, tmpDIBLength
    End If
        
    'If this is not an embedding, close the file now
    If (Not embedInFile) Then cFile.FileCloseHandle hFile

End Function

'Write this class's DIB information to a valid BMP file.
Friend Sub WriteToBitmapFile(ByRef dstFilename As String)

    'First, make sure this DIB actually contains image data.
    ' (This step will also unsuspend the DIB, as necessary.)
    If (Not Me.HasImage()) Then Exit Sub
    
    'Calculate the size of a scanline
    Dim slWidth As Long
    
    If (Me.GetDIBColorDepth = 32) Then
        slWidth = Me.GetDIBStride
    Else
        slWidth = ((m_dibWidth * m_dibColorDepth + 31) \ 32) * 4
    End If
    
    'A valid bitmap header consists of two parts: a file header, and an image header.
    Dim bmpFileHeader As BITMAPFILEHEADER
    Dim bmpInfoHeader As BITMAPINFOHEADER
    
    'Build the file header first
    With bmpFileHeader
        
        'BMP identifier
        .Type = &H4D42
        
        'Length of the file
        .Size = LenB(bmpFileHeader) + LenB(bmpInfoHeader) + (slWidth * Me.GetDIBHeight)
        
        'Length of the header area (e.g. offset to the actual pixel data)
        .OffBits = LenB(bmpFileHeader)
        
    End With
    
    '...then the image header
    With bmpInfoHeader
        .Size = 40
        .Planes = 1
        .BitCount = Me.GetDIBColorDepth()
        .Width = Me.GetDIBWidth()
        .Height = Me.GetDIBHeight()
    End With
    
    'Finally, the image bytes.  We call GetDIBits here because DIBs are upside-down, but we need to store them rightside-up in the BMP.
    ' We could do this manually on a per-scanline basis, but this is more succinct.
    Dim iData() As Byte
    ReDim iData(0 To slWidth - 1, 0 To m_dibHeight - 1) As Byte
    GetDIBits Me.GetDIBDC, m_dibHandle, 0, m_dibHeight, iData(0, 0), bmpInfoHeader, 0&
    
    'With all our information in place, check to see if that file already exists. Delete it if necessary.
    Files.FileDeleteIfExists dstFilename
    
    'Writing the file is easy; just dump the two headers and image data into it!
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    Dim hFile As Long
    If cFile.FileCreateHandle(dstFilename, hFile, False, True, OptimizeSequentialAccess) Then
    
        With cFile
            
            'To avoid automatic struct alignment, we'll write out the header manually.
            .FileWriteData hFile, VarPtr(bmpFileHeader.Type), 2&
            .FileWriteData hFile, VarPtr(bmpFileHeader.Size), 4&

            Dim reservedBytes As Long
            .FileWriteData hFile, VarPtr(reservedBytes), 4&
            .FileWriteData hFile, VarPtr(bmpFileHeader.OffBits), 4&
            
            'With the header finished, we add the lightweight BMP header and the actual pixel data
            .FileWriteData hFile, VarPtr(bmpInfoHeader), LenB(bmpInfoHeader)
            .FileWriteData hFile, VarPtr(iData(0, 0)), slWidth * m_dibHeight
            
        End With
        
        'That's it!
        cFile.FileCloseHandle hFile
        PDDebug.LogAction "pdDIB.writeToBitmapFile successfully wrote " & dstFilename & "."
        
    Else
        PDDebug.LogAction "WARNING!  pdDIB.writeToBitmapFile failed to create a valid file handle for " & dstFilename & ".  Write abandoned."
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
        
End Sub

'Return a pointer to DIB byte 0, and return the net size of the DIB's pixel contents.  (PD frequently
' uses this to wrap arbitrary arrays or structs around DIB contents.)
Friend Sub RetrieveDIBPointerAndSize(ByRef dibPointer As Long, ByRef dibSize As Long)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
        
    dibPointer = m_dibBits
    dibSize = m_dibStride * m_dibHeight
    
End Sub

'Sometimes this class needs to access its own DIB bits. Here's how.
Private Sub PrepInternalSafeArray(ByRef dstSafeArray As SafeArray2D)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_dibStride
        .pvData = m_dibBits
    End With
    
End Sub

Private Sub PrepInternalSafeArray_Scanline(ByRef dstSafeArray As SafeArray1D, ByVal dstLine As Long)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_dibBits + (dstLine * m_dibStride)
    End With
    
End Sub

Private Sub PrepInternalLongSafeArray(ByRef dstSafeArray As SafeArray2D)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    With dstSafeArray
        .cbElements = 4
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_dibWidth
        .pvData = m_dibBits
    End With
    
End Sub

Private Sub PrepInternalLongSafeArray_Scanline(ByRef dstSafeArray As SafeArray1D, ByVal dstLine As Long)
    
    'Unsuspend this DIB (as necessary)
    If m_IsSuspended Then UnsuspendDIB
    
    With dstSafeArray
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibWidth
        .pvData = m_dibBits + (dstLine * m_dibStride)
    End With
    
End Sub

'Pre-composite an image with an alpha-channel against a background color.
Friend Sub CompositeBackgroundColor(Optional ByVal newR As Byte = 255, Optional ByVal newG As Byte = 255, Optional ByVal newB As Byte = 255)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then Exit Sub
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access),
    ' so we cheat and use a 1D array, which we reset between scanlines.  (The pointer retrieval will also
    ' unsuspend the underlying DIB, as necessary.)
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim dibPixels() As Byte, dibSA As SafeArray1D
    Me.WrapArrayAroundScanline dibPixels, dibSA, 0
    
    Dim x As Long, y As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'Because our alpha values are pre-multiplied, we can composite them against the background color via use of a look-up table.
    Dim rLookup(0 To 255) As Byte, gLookup(0 To 255) As Byte, bLookup(0 To 255) As Byte
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Populate a unique lookup table for each color, based on each possible alpha value (0 to 255)
    For x = 0 To 255
        tmpAlpha = 1# - (x * ONE_DIV_255)
        rLookup(x) = Int(newR * tmpAlpha)
        gLookup(x) = Int(newG * tmpAlpha)
        bLookup(x) = Int(newB * tmpAlpha)
    Next x
    
    'Loop through the image, compositing as we go
    For y = 0 To finalY
        
        'Point our 1D pixel array at the proper scanline
        dibSA.pvData = dibPointer + scanlineSize * y
        
    For x = 0 To finalX Step 4
        
        'Access the alpha data for this pixel
        checkAlpha = dibPixels(x + 3)
        
        'Ignore opaque pixels
        If (checkAlpha <> 255) Then
            
            'Handle transparent pixels specially (this improves performance)
            If (checkAlpha = 0) Then
                dibPixels(x) = newB
                dibPixels(x + 1) = newG
                dibPixels(x + 2) = newR
            Else
            
                'Use that alpha value to blend the current colors with the newly requested color
                If m_IsAlphaPremultiplied Then
                    dibPixels(x) = dibPixels(x) + bLookup(checkAlpha)
                    dibPixels(x + 1) = dibPixels(x + 1) + gLookup(checkAlpha)
                    dibPixels(x + 2) = dibPixels(x + 2) + rLookup(checkAlpha)
                Else
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = checkAlpha * ONE_DIV_255
                
                    'Use that alpha value to blend the current colors with the newly requested color
                    dibPixels(x) = Blend2Colors(dibPixels(x), newB, tmpAlpha)
                    dibPixels(x + 1) = Blend2Colors(dibPixels(x + 1), newG, tmpAlpha)
                    dibPixels(x + 2) = Blend2Colors(dibPixels(x + 2), newR, tmpAlpha)
                    
                End If
                
            End If
            
            dibPixels(x + 3) = 255
            
        End If
        
    Next x
    Next y
    
    'With our alpha channel complete, point dibPixels() away from the DIB
    Me.UnwrapArrayFromDIB dibPixels
    
    'A composited image is always premultiplied
    Me.SetInitialAlphaPremultiplicationState True
    
End Sub

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function Blend2Colors(ByVal Color1 As Long, ByVal Color2 As Long, ByRef mixRatio As Double) As Byte
    Blend2Colors = mixRatio * (Color1 - Color2) + Color2
End Function

'Copy the alpha values from another pdDIB object to this one. This is useful when the alpha channel for this DIB must be lost
' during a transformation (typically something involving FreeImage or GDI+), and a temporary DIB was made to preserve the alpha
' data. If the image is currently 24bpp, this function will first convert it to 32bpp before copying the alpha data.
'
'IMPORTANT NOTE: this function does not deal with premultiplied alpha whatsoever.  It simply copies alpha data as-is.
' You must manually set the alpha premultiplication property yourself, and you must also make sure that any number of
' un/premultiply steps are applied to both DIBs to keep the alpha values properly in sync.  The sole exception to this
' occurs if you pass the binaryAlphaOnly parameter as TRUE; this means the alpha mask only contains 0 and 255 values,
' and if that's the case, this function will handle premultiplication for you.

'IMPORTANT NOTE: to keep this function fast, NO ARRAY BOUNDS CHECKING IS DONE. Make sure that the passed pdDIB object is
' THE SAME SIZE (including scanline alignment) as this DIB, or you will experience critical errors.
Friend Sub CopyAlphaFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal binaryAlphaOnly As Boolean = False)
    
    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then Exit Sub
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access),
    ' so we cheat and use a 1D array, which we reset between scanlines.  The pointer retrieval call will
    ' also unsuspend each DIB, as necessary.
    Dim srcScanlineSize As Long, srcPointer As Long, dstScanlineSize As Long, dstPointer As Long
    dstScanlineSize = Me.GetDIBStride
    dstPointer = Me.GetDIBPointer
    srcScanlineSize = srcDIB.GetDIBStride
    srcPointer = srcDIB.GetDIBPointer
    
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    Me.WrapArrayAroundScanline dstPixels, dstSA, 0
    
    Dim srcPixels() As Byte, srcSA As SafeArray1D
    srcDIB.WrapArrayAroundScanline srcPixels, srcSA, 0
    
    Dim x As Long, y As Long
    
    'Loop through the image, compositing as we go
    For y = 0 To finalY
        
        'Point our 1D pixel arrays at the proper scanlines
        dstSA.pvData = dstPointer + dstScanlineSize * y
        srcSA.pvData = srcPointer + srcScanlineSize * y
        
        'Split handling based on premultiplication
        If binaryAlphaOnly Then
            
            For x = 0 To finalX Step 4
                If (srcPixels(x + 3) = 255) Then
                    dstPixels(x + 3) = 255
                Else
                    dstPixels(x) = 0
                    dstPixels(x + 1) = 0
                    dstPixels(x + 2) = 0
                    dstPixels(x + 3) = 0
                End If
            Next x
        
        Else
            For x = 0 To finalX Step 4
                dstPixels(x + 3) = srcPixels(x + 3)
            Next x
        End If
        
    Next y
    
    'With our alpha channel complete, point dibPixels() away from the DIB
    Me.UnwrapArrayFromDIB dstPixels
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    If binaryAlphaOnly Then Me.SetInitialAlphaPremultiplicationState True
    
End Sub

'This routine can be used to either apply or remove premultiplied alpha values.
' Optionally, you can also supply a pointer to a valid RectF struct; if present, only this region will be updated.
Friend Sub SetAlphaPremultiplication(Optional ByVal applyPremultiplication As Boolean = False, Optional ByVal ignoreEmbeddedValue As Boolean = False, Optional ByVal ptrToRectF As Long = 0)

    'This function doesn't matter if the image isn't 32bpp
    If (m_dibColorDepth <> 32) Then Exit Sub
    
    'If alpha premultiplication already matches the requested state, exit now
    If (m_IsAlphaPremultiplied = applyPremultiplication) And (Not ignoreEmbeddedValue) Then
        Debug.Print "WARNING! Alpha premultiplication = " & applyPremultiplication & " is pointless, as image is already in that state.  (" & m_dibWidth & "x" & m_dibHeight & "x" & m_dibColorDepth & ")  Abandoning request."
        Exit Sub
    End If
    
    Dim allowedToProceed As Boolean: allowedToProceed = True
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    If (ptrToRectF = 0) Then
        initX = 0
        initY = 0
        finalX = (m_dibWidth - 1) * 4
        finalY = (m_dibHeight - 1)
    Else
        
        Dim tmpRectF As RectF
        CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
        PDMath.GetIntClampedRectF tmpRectF
        initX = tmpRectF.Left
        initY = tmpRectF.Top
        finalX = initX + tmpRectF.Width
        finalY = initY + tmpRectF.Height
        
        'Perform a bunch of failsafe checks on boundaries
        If (initX > (m_dibWidth - 1)) Then allowedToProceed = False
        If (initY > (m_dibHeight - 1)) Then allowedToProceed = False
        If (finalX < initX) Then allowedToProceed = False
        If (finalY < initY) Then allowedToProceed = False
        
        If allowedToProceed Then
            If (initX < 0) Then initX = 0
            If (initY < 0) Then initY = 0
            If (finalX > (m_dibWidth - 1)) Then finalX = m_dibWidth - 1
            If (finalY > (m_dibHeight - 1)) Then finalY = m_dibHeight - 1
        End If
        
        initX = initX * 4
        finalX = finalX * 4
                
    End If
    
    If allowedToProceed Then
        
        Const ONE_DIV_255 As Single = 1! / 255!
        
        'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
        ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
        Dim intToFloat(0 To 255) As Single
        Dim i As Long
        For i = 0 To 255
            If applyPremultiplication Then
                intToFloat(i) = CSng(i) * ONE_DIV_255
            Else
                If (i <> 0) Then intToFloat(i) = CSng(255! / CDbl(i))
            End If
        Next i
        
        '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access),
        ' so we cheat and use a 1D array, which we reset between scanlines.  The pointer retrieval will also
        ' unsuspend this DIB, as necessary.
        Dim scanlineSize As Long, dibPointer As Long
        scanlineSize = Me.GetDIBStride
        dibPointer = Me.GetDIBPointer
        
        Dim dibPixels() As Byte, dibSA As SafeArray1D
        Me.WrapArrayAroundScanline dibPixels, dibSA, initY
        
        Dim x As Long, y As Long
        Dim r As Long, g As Long, b As Long
        Dim tmpAlpha As Byte, tmpAlphaModifier As Single
        
        'Loop through the image, converting alpha as we go
        For y = initY To finalY
            
            'Point our 1D pixel array at the proper scanline.  The odd construction of this line
            ' is meant to work around issues near the 2GB boundary.
            If (y > initY) Then dibSA.pvData = VBHacks.UnsignedAdd(dibSA.pvData, scanlineSize)
            
        For x = initX To finalX Step 4
            
            'Retrieve alpha for the current pixel
            tmpAlpha = dibPixels(x + 3)
            
            'Branch according to applying or removing premultiplication
            If applyPremultiplication Then
            
                'When applying premultiplication, we can ignore fully opaque pixels
                If (tmpAlpha <> 255) Then
                
                    'We can shortcut the calculation of full transparent pixels (they are made black)
                    If (tmpAlpha = 0) Then
                        dibPixels(x) = 0
                        dibPixels(x + 1) = 0
                        dibPixels(x + 2) = 0
                    Else
                
                        b = dibPixels(x)
                        g = dibPixels(x + 1)
                        r = dibPixels(x + 2)
                        
                        tmpAlphaModifier = intToFloat(tmpAlpha)
                        
                        'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                        r = Int(r * tmpAlphaModifier + 0.5!)
                        g = Int(g * tmpAlphaModifier + 0.5!)
                        b = Int(b * tmpAlphaModifier + 0.5!)
                        
                        dibPixels(x) = b
                        dibPixels(x + 1) = g
                        dibPixels(x + 2) = r
                        
                    End If
                
                End If
            
            Else
                
                'When removing premultiplication, we can ignore fully opaque and fully transparent pixels.
                ' (Note that VB doesn't short-circuit AND statements, so we manually nest the IFs.)
                If (tmpAlpha <> 255) Then
                    If (tmpAlpha <> 0) Then
                    
                        b = dibPixels(x)
                        g = dibPixels(x + 1)
                        r = dibPixels(x + 2)
                        
                        tmpAlphaModifier = intToFloat(tmpAlpha)
                        
                        'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                        r = Int(r * tmpAlphaModifier + 0.5!)
                        g = Int(g * tmpAlphaModifier + 0.5!)
                        b = Int(b * tmpAlphaModifier + 0.5!)
                        
                        'Unfortunately, OOB checks are necessary for malformed DIBs
                        If (r > 255) Then r = 255
                        If (g > 255) Then g = 255
                        If (b > 255) Then b = 255
                        
                        dibPixels(x) = b
                        dibPixels(x + 1) = g
                        dibPixels(x + 2) = r
                        
                    End If
                End If
            
            End If
            
        Next x
        Next y
        
        'With our alpha channel complete, point dibPixels() away from the DIB
        Me.UnwrapArrayFromDIB dibPixels
        
        'Mark the new premultiplication state
        m_IsAlphaPremultiplied = applyPremultiplication
        
    End If
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Friend Sub AlphaBlendToDC(ByVal dstDC As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    Dim bfParams As Long
    If (newWidth = 0) Then newWidth = m_dibWidth
    If (newHeight = 0) Then newHeight = m_dibHeight
    
    If (m_dibColorDepth = 32) Then
        
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Also, raise a warning if premultiplication is not set
        If (Not m_IsAlphaPremultiplied) Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDC being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
        
    AlphaBlend dstDC, dstX, dstY, newWidth, newHeight, Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, bfParams
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Friend Sub AlphaBlendToDCEx(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customAlpha As Long = 255)
    
    Dim bfParams As Long
    
    If (m_dibColorDepth = 32) Then
    
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Raise a warning if premultiplication is not set
        If (Not m_IsAlphaPremultiplied) Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDCEx being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
    
    AlphaBlend dstDC, dstX, dstY, dstWidth, dstHeight, Me.GetDIBDC, srcX, srcY, srcWidth, srcHeight, bfParams
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub
