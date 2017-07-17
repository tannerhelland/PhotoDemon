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
'Copyright 2012-2017 by Tanner Helland
'Created: 29/August/12
'Last updated: 18/April/17
'Last update: store the original file format in a proper PDIF enum
'
'This class is the core of PhotoDemon.  Any image interactions are handled through this lightweight Windows DIB wrapper, and as you
' might expect, a number of other capabilities are built into the class.
'
'Note that anything you can do with an hDC property you can do with this class - simply use the getDIBDC function to
' return the DIB's hDC, then do with it what you please. All functions are heavily commented and should be self-explanatory.
'
'SPECIAL THANKS
'In building this class, I utilized a number of other DIB classes for reference and testing. Special thanks to:
' Carles PV's iBMP project: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Steve McMahon's DIB/SafeArray analysis: http://www.vbaccelerator.com/home/VB/Code/vbMedia/DIB_Sections/True_Colour_DIBSection/article.asp
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'DIB Types
'In PD, GDI bitmaps are declared inside the GDI module; to use this as a standalone class, uncomment this type dec
'Private Type GDI_Bitmap
'    Type As Long
'    Width As Long
'    Height As Long
'    WidthBytes As Long
'    Planes As Integer
'    BitsPerPixel As Integer
'    Bits As Long
'End Type

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
    Colorused As Long
    ColorImportant As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

'Rectangle type for use with SetRect/FillRect API calls
Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data).
' (Inside PhotoDemon, we declare these types elsewhere, as they're used in a whole bunch of functions, but to use this class
'  as a standalone instance you will need to uncomment these.)
'Public Type SAFEARRAYBOUND
'    cElements As Long
'    lBound As Long
'End Type
'
'Public Type SAFEARRAY2D
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'    Bounds(1) As SAFEARRAYBOUND
'End Type

'Drawing API functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCHBLT_COLORONCOLOR As Long = 3
Private Const STRETCHBLT_HALFTONE As Long = 4

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DIB API functions
' (Note that these are currently declared in EffectPrep as well)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'Object API functions
Private Const OBJ_BITMAP As Long = 7
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long

'AlphaBlend API call
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

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

'Used when writing/reading DIB data to/from a file
Private Const DIB_IDENTIFIER As Long = &H726C4450  'The ASCII chars "PDlr"
Private Const DIB_FILE_VERSION_2012 As Long = &H1000&
Private Const DIB_FILE_VERSION_2015 As Long = &H1001&
Private Const DIB_FILE_VERSION_2016 As Long = &H1002&

'Remember the original color of the pixels made transparent, then restore it after the "magic magenta" has served its purpose
Private m_OriginalTransparentColor As Long

'If this DIB was loaded from file, its original file format will be set here.  (By default, this value will be -1).
Private m_OriginalFIF As PD_IMAGE_FORMAT

'DIBs may require ICC profile handling.  All ICC Profile interactions are handled through a publicly available ICCProfile class,
' on a per-DIB basis.
Public ICCProfile As pdICCProfile

'DIBs created from loaded image files contain resolution data.  This data isn't important on a per-DIB basis, but the parent
' pdImage object will copy the resolution data from its first-loaded child DIB (if it has a resolution)
Private m_XResolution As Double, m_YResolution As Double, m_dibDPI As Double

'Some image formats (e.g. PNG) will specify a preferred background color.  If available at load-time, that color value
' will be stored here.
Private m_OriginalBackgroundColor As Long

'DIBs created from loaded image files will have their original color depth stored here.  Note that PD only works in 24/32 bpp
' mode at present, but this value may contain other bit-depths (e.g. 8bpp if the source data was a GIF).
Private m_OriginalColorDepth As Long

'If FreeImage is active, it allows us to retrieve even more detailed color depth information (vs GDI+, which downsamples unpredictably).
' While I don't really make use of this value at present, it could be helpful in the future, for example when editing icons (which contain
' multiple icons at varying sizes and original color depths).
Private m_OriginalFreeImageColorDepth As Long

'If the alpha is currently premultiplied, this will be set to TRUE
Private m_IsAlphaPremultiplied As Boolean

'Get/set alpha premultiplication.
' IMPORTANT NOTE!  To make it explicitly clear that modifying this property DOES NOT ACTUALLY MODIFY THE IMAGE, the Set instruction is
' labeled differently.  It is only meant to be used by DIB creation functions, where the premultiplication state is explicitly known prior
' to writing DIB bits.  The counterpart setAlphaPremultiplication function (which is found further down in this file) will actually
' modify image bits as necessary to create the desired premultiplication state.
Public Function GetAlphaPremultiplication() As Boolean
    GetAlphaPremultiplication = m_IsAlphaPremultiplied
End Function

Public Sub SetInitialAlphaPremultiplicationState(ByVal newState As Boolean)
    m_IsAlphaPremultiplied = newState
End Sub

'Get/set original color depth.  Note that this is set after tone-mapping and other actions have been applied
Public Function GetOriginalColorDepth() As Long
    GetOriginalColorDepth = m_OriginalColorDepth
End Function

Public Sub SetOriginalColorDepth(ByVal origColorDepth As Long)
    m_OriginalColorDepth = origColorDepth
End Sub

'Get/set original file color depth.  This is set prior to any tone-mapping or explicit FreeImage conversions.  As such, it may be quite large.
Public Function GetOriginalFreeImageColorDepth() As Long
    GetOriginalFreeImageColorDepth = m_OriginalFreeImageColorDepth
End Function

Public Sub SetOriginalFreeImageColorDepth(ByVal origFIColorDepth As Long)
    m_OriginalFreeImageColorDepth = origFIColorDepth
End Sub

'If this DIB was created from an image file, and that file specified a preferred background color, use this
' function to copy the value into this DIB.
Public Sub SetBackgroundColor(ByVal newBackColor As Long)
    m_OriginalBackgroundColor = newBackColor
End Sub

Public Function GetBackgroundColor() As Long
    GetBackgroundColor = m_OriginalBackgroundColor
End Function

'Set the DPI of this DIB.  This is only relevant if this DIB has been created directly from an image file.
Public Sub SetDPI(ByVal xRes As Double, ByVal yRes As Double, Optional ByVal sourceIsFreeImage As Boolean = False)
    
    'Many image types do not store resolution information; default to 96 in this case
    If xRes = 0 Then xRes = 96
    If yRes = 0 Then yRes = 96
    
    'FreeImage is pretty damn stupid when it comes to DPI.  If no DPI information is found, it will return "72" by default.
    ' If the source of this call is FreeImage, we want to replace 72 with a more modern default of 96.  Obviously this creates
    ' a problem when images have an actual resolution of 72 DPI, as there is no way to tell if FreeImage set that by default,
    ' or if it is the actual DPI value of the file.  If this occurs, we can use ExifTool to retrieve an actual value as
    ' stored in the file.
    If sourceIsFreeImage Then
        If xRes = 72 Then xRes = 96
        If yRes = 72 Then yRes = 96
    End If
    
    m_XResolution = xRes
    m_YResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    m_dibDPI = (xRes + yRes) \ 2

End Sub

'Even though we store separate x and y DPI, PD only deals in a single per-DIB DPI value
Public Function GetDPI() As Double
    GetDPI = m_dibDPI
End Function

'Get/Set original DIB format
Public Function GetOriginalFormat() As PD_IMAGE_FORMAT
    GetOriginalFormat = m_OriginalFIF
End Function

Public Sub SetOriginalFormat(ByVal origFIF As PD_IMAGE_FORMAT)
    m_OriginalFIF = origFIF
End Sub

'Return the original transparent color (set as part of the MarkAlpha routine).
Public Function GetOriginalTransparentColor() As Long
    GetOriginalTransparentColor = m_OriginalTransparentColor
End Function

'Given a pixel coordinate, return an RGBA quad for that coordinate.
' Returns: TRUE if the pixel lies inside DIB boundaries; FALSE otherwise.  Make sure to check this before using the RGBQUAD.
Friend Function GetPixelRGBQuad(ByVal x As Long, ByVal y As Long, ByRef dstQuad As RGBQUAD) As Boolean

    'Before doing anything else, check to see if the x/y postition lies inside the DIB
    If (x >= 0) And (x < m_dibWidth) And (y >= 0) And (y < m_dibHeight) Then
        
        'The point lies inside the DIB, which means we need to figure out the color at this position
        GetPixelRGBQuad = True
        
        Dim tmpData() As Byte, tSA As SAFEARRAY2D
        Me.WrapArrayAroundDIB tmpData, tSA
        
        Dim quickX As Long
        quickX = x * (m_dibColorDepth \ 8)
        
        'Failsafe bounds check
        If ((quickX + 2) < m_dibStride) Then
        
            With dstQuad
                .Blue = tmpData(quickX, y)
                .Green = tmpData(quickX + 1, y)
                .Red = tmpData(quickX + 2, y)
                If m_dibColorDepth = 32 Then .alpha = tmpData(quickX + 3, y)
            End With
            
        End If
        
        Me.UnwrapArrayFromDIB tmpData
        
    'This coordinate does not lie inside the layer.
    Else
        GetPixelRGBQuad = False
    End If

End Function

'Convenience functions for wrapping an array around this DIB's bits.  You *must* call the Unwrap function prior to the array
' falling out of scope, or VB will crash.
Friend Sub WrapArrayAroundDIB(ByRef srcArray() As Byte, ByRef srcSafeArray As SAFEARRAY2D)
    PrepInternalSafeArray srcSafeArray
    CopyMemory ByVal VarPtrArray(srcArray()), VarPtr(srcSafeArray), 4&
End Sub

Friend Sub WrapArrayAroundScanline(ByRef srcArray() As Byte, ByRef srcSafeArray As SAFEARRAY1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalSafeArray_Scanline srcSafeArray, dstScanLine
    CopyMemory ByVal VarPtrArray(srcArray()), VarPtr(srcSafeArray), 4&
End Sub

Public Sub UnwrapArrayFromDIB(ByRef srcArray() As Byte)
    CopyMemory ByVal VarPtrArray(srcArray), 0&, 4&
End Sub

Friend Sub WrapLongArrayAroundDIB(ByRef srcArray() As Long, ByRef srcSafeArray As SAFEARRAY2D)
    PrepInternalLongSafeArray srcSafeArray
    CopyMemory ByVal VarPtrArray(srcArray()), VarPtr(srcSafeArray), 4&
End Sub

Friend Sub WrapLongArrayAroundScanline(ByRef srcArray() As Long, ByRef srcSafeArray As SAFEARRAY1D, Optional ByVal dstScanLine As Long = 0)
    PrepInternalLongSafeArray_Scanline srcSafeArray, dstScanLine
    CopyMemory ByVal VarPtrArray(srcArray()), VarPtr(srcSafeArray), 4&
End Sub

Public Sub UnwrapLongArrayFromDIB(ByRef srcArray() As Long)
    CopyMemory ByVal VarPtrArray(srcArray), 0&, 4&
End Sub

'Force all alpha bytes in the DIB to some preset value.  If the value is *not* 255, please remember to premultiply
' alpha accordingly, if the DIB will be rendered to the screen.
Public Function ForceNewAlpha(ByVal newAlpha As Byte) As Boolean
    
    ForceNewAlpha = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If (m_dibColorDepth = 32) Then

        'Make sure this DIB isn't empty
        If (m_dibHandle <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            'Loop through the image and force each alpha value to the user's specified value
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
            Me.WrapArrayAroundDIB iData, tmpSA
            
            Dim x As Long, y As Long, finalX As Long
            finalX = (m_dibWidth - 1) * 4
                
            'Loop through the image, checking alphas as we go
            For y = 0 To m_dibHeight - 1
            For x = 0 To finalX Step 4
                iData(x + 3, y) = newAlpha
            Next x
            Next y
            
            Me.UnwrapArrayFromDIB iData
            ForceNewAlpha = True
            
        Else
            Debug.Print "WARNING!  The target DIB is empty (DC or Width or Height = 0).  pdDIB.ForceNewAlpha failed."
        End If
        
    Else
        Debug.Print "WARNING!  You cannot call pdDIB.ForceNewAlpha on a 24-bpp DIB!"
    End If
    
End Function

'Translate an OLE color to an RGB Long
Private Function TranslateColor(ByVal colorRef As Long) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, TranslateColor) Then
        TranslateColor = RGB(255, 255, 255)
    End If
End Function

Public Sub FreeFromDC()
    If (m_dibDC <> 0) Then
        SelectObject m_dibDC, m_dibHandleOriginal
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
    End If
End Sub

'Return this DIB's color depth
Public Function GetDIBColorDepth() As Long
    GetDIBColorDepth = m_dibColorDepth
End Function

'Return this DIB's array width
Public Function GetDIBStride() As Long
    GetDIBStride = m_dibStride
End Function

'Return this DIB's width
Public Function GetDIBWidth() As Long
    GetDIBWidth = m_dibWidth
End Function

'Return this DIB's height
Public Function GetDIBHeight() As Long
    GetDIBHeight = m_dibHeight
End Function

'Return whether or not this DIB has image data associated with it
Public Function HasImage() As Boolean
    HasImage = (m_dibHandle <> 0)
End Function

'Return whether the DIB is top-down (negative height) or bottom-up (positive height)
Public Function IsDIBTopDown() As Boolean
    IsDIBTopDown = CBool(m_dibHeader.Height < 0)
End Function

'Return this DIB's hDC.  DIBs receive a DC by default, at creation-time.  However, the caller can remove the DIB from its DC by
' calling the FreeFromDC() sub.  If you subsequently request a DC for the DIB, one will be auto-created for you.
Public Function GetDIBDC() As Long
    If (m_dibDC <> 0) Then
        GetDIBDC = m_dibDC
    Else
        If (m_dibHandle <> 0) Then
            m_dibDC = GDI.GetMemoryDC()
            m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
            GetDIBDC = m_dibDC
        Else
            GetDIBDC = 0
        End If
    End If
End Function

'Return a pointer to this DIB's actual DIB
Public Function GetDIBHandle() As Long
    GetDIBHandle = m_dibHandle
End Function

'Return a pointer to this DIB's pixel data (no header)
Public Function GetDIBPointer() As Long
    GetDIBPointer = m_dibBits
End Function

'Return a pointer to an individual scanline.  For performance reasons, this function does not perform bounds checking -
' so use it carefully!
Public Function GetDIBScanline(ByVal targetScanline As Long) As Long
    GetDIBScanline = m_dibBits + (targetScanline * m_dibStride)
End Function

'Return a pointer to this DIB's header
Public Function GetDIBHeader() As Long
    GetDIBHeader = VarPtr(m_dibHeader)
End Function

Friend Sub CopyDIBHeader(ByRef dstHeader As BITMAPINFOHEADER)
    dstHeader = m_dibHeader
End Sub

'Quickly create a new DIB from an arbitrary DC.  No stretching is applied, by design; StretchBlt is unpredictable with
' 32-bpp data, so any stretching must be applied *after* creation (when PD has full control over the image data).
Public Function CreateFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32, Optional ByVal isSourcePremultiplied As Boolean = False) As Boolean
    
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
            If CreateBlank(srcWidth, srcHeight, srcColorDepth, 0, 255) Then
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
Public Function CreateFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal newWidth As Long = -1, Optional ByVal newHeight As Long = -1, Optional ByVal useHQStretching As Boolean = True) As Boolean
    
    CreateFromExistingDIB = False
    
    'Make sure the DIB we're passed isn't empty
    If (srcDIB.GetDIBDC <> 0) Then
        
        'Prepare new width and height values as requested by the user
        If (newWidth = -1) Then newWidth = srcDIB.GetDIBWidth
        If (newHeight = -1) Then newHeight = srcDIB.GetDIBHeight
        
        'If the width and height values are not being changed, the transfer is simple
        If (newWidth = srcDIB.GetDIBWidth) And (newHeight = srcDIB.GetDIBHeight) Then
            
            'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
            If (newWidth = m_dibWidth) And (newHeight = m_dibHeight) And (srcDIB.GetDIBColorDepth = m_dibColorDepth) Then
            
                'Copy the image data without modification
                BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                
                'Copy the source DIB's header as well.  (Note that this only copies a BITMAPINFOHEADER,
                ' meaning descriptive info only - not things like the DIB bits pointer.)
                srcDIB.CopyDIBHeader m_dibHeader
                
                'Reset our alpha premultiplication flag to match
                m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
                
                CreateFromExistingDIB = True
                
            Else
            
                'Create a new, blank DIB the same size as the source DIB
                If Me.CreateBlank(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBColorDepth, 0, 0) Then
                    
                    'Copy the image data without modification
                    BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                    
                    'Reset our alpha premultiplication flag to match
                    m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
                    
                    CreateFromExistingDIB = True
                    
                End If
                
            End If
        
        'If new width and height values are being specified, the transfer is a bit more complex
        Else
            
            'If our current size does not match, create a newly sized DIB in advance.
            If (Me.GetDIBWidth <> newWidth) Or (Me.GetDIBHeight <> newHeight) Or (Me.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then Me.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth
            
            Dim resizeMode As GP_InterpolationMode
            If useHQStretching Then resizeMode = GP_IM_HighQualityBicubic Else resizeMode = GP_IM_NearestNeighbor
            GDI_Plus.GDIPlusResizeDIB Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, resizeMode
                
            'Reset alpha premultiplication to match
            m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
            CreateFromExistingDIB = True
            
        End If
        
        'Free this DIB and the source DIB from their DCs, as they may not be required again for some time
        srcDIB.FreeFromDC
        Me.FreeFromDC
        
    End If
    
End Function

'Convert this DIB to 24bpp mode
Public Function ConvertTo24bpp(Optional ByVal newBackColor As Long = vbWhite) As Boolean
    
    ConvertTo24bpp = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function pointless.
    If (m_dibColorDepth = 32) Then

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
Public Function ConvertTo32bpp(Optional ByVal newTransparency As Byte = 255) As Boolean
    
    ConvertTo32bpp = False
    
    'Make sure this DIB isn't empty
    If (m_dibHandle <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then

        Dim iData() As Byte, tmpSA As SAFEARRAY2D
        Dim x As Long, y As Long, quickX As Long
        
        Dim newTransFloat As Double
        newTransFloat = newTransparency / 255

        If (Me.GetDIBColorDepth = 32) Then
            
            'If we are already 32bpp, with premultiplied alpha, unpremultiply the temporary DIB now
            If Me.GetAlphaPremultiplication Then Me.SetAlphaPremultiplication False
            
            'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
            ' be completely transparent (and we don't want that!)
            PrepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
            
            'Loop through the image, setting the alpha of each pixel to 255 (opaque)
            For y = 0 To m_dibHeight - 1
            For x = 0 To m_dibWidth - 1
                quickX = x * 4
                iData(quickX + 3, y) = iData(quickX + 3, y) * newTransFloat
            Next x
            Next y

            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            'Per PD-convention, re-apply premultiplication now
            Me.SetAlphaPremultiplication True
            
            ConvertTo32bpp = True
            
        Else
        
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.CreateFromExistingDIB Me
            
            'Now erase our own DIB
            Me.EraseDIB
    
            'Create a new DIB that's exactly the same size as the old one
            If Me.CreateBlank(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 32) Then
                
                'Copy the image data from the temporary DIB without modification
                BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                
                'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
                ' be completely transparent (and we don't want that!)
                PrepInternalSafeArray tmpSA
                CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
                
                'Loop through the image, setting the alpha of each pixel to 255 (opaque)
                For y = 0 To m_dibHeight - 1
                For x = 0 To m_dibWidth - 1
                    iData(x * 4 + 3, y) = newTransparency
                Next x
                Next y
                
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                
                'If the alpha value is not 255, apply premultiplication now
                If (newTransparency <> 255) Then Me.SetAlphaPremultiplication True Else Me.SetInitialAlphaPremultiplicationState True
                
                ConvertTo32bpp = True
                
            End If
            
        End If
        
        'Minimize GDI resources by freeing our DC
        Me.FreeFromDC
        
    End If
    
End Function

'Give this DIB a picture from a standard VB picture object
Public Function CreateFromPicture(ByRef srcPicture As StdPicture, Optional forceWhiteBackground As Boolean = False) As Boolean

    CreateFromPicture = False
    
    'Make sure the picture we're passed isn't empty
    If (Not srcPicture Is Nothing) Then
    
        'Make sure the picture is actually a picture
        If (GetObjectType(srcPicture) = OBJ_BITMAP) Then
        
            'Select the picture's attributes into a bitmap object
            Dim tmpBitmap As GDI_Bitmap
            GetObject srcPicture.Handle, Len(tmpBitmap), tmpBitmap
            
            'Use that bitmap object to create a new, blank DIB of the same size
            Dim targetColorDepth As Long
            If (tmpBitmap.BitsPerPixel = 32) Then targetColorDepth = 32 Else targetColorDepth = 24
            If Me.CreateBlank(tmpBitmap.Width, tmpBitmap.Height, targetColorDepth, , 255) Then
            
                'Create a new DC
                Dim tmpDC As Long
                tmpDC = GDI.GetMemoryDC()
                
                'If successful, select the object into that DC
                If (tmpDC <> 0) Then
                
                    'Temporary holder for the object selection
                    Dim oldBitmap As Long
                    oldBitmap = SelectObject(tmpDC, srcPicture.Handle)
                    
                    'Use BitBlt to copy the pixel data to this DIB
                    BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, tmpDC, 0, 0, vbSrcCopy
                    
                    'Now that we have the pixel data, erase all temporary objects
                    SelectObject tmpDC, oldBitmap
                    GDI.FreeMemoryDC tmpDC
                    
                    'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                    ' black background. We typically want the background to be white, so perform a composite if requested.
                    If forceWhiteBackground And (m_dibColorDepth = 32) Then CompositeBackgroundColor
                    
                    CreateFromPicture = True
                    
                    'Minimize GDI resources by freeing our DC
                    Me.FreeFromDC
                    
                End If
            
            End If
            
        End If
        
    End If
    
End Function

'Create a blank DIB. If no colorDepth is specified, it will default to 24bpp (16 million colors, no alpha-channel).
' If the DIB is 32bpp, an optional alpha parameter can be set for the DIB.
Public Function CreateBlank(ByVal imgWidth As Long, ByVal imgHeight As Long, Optional ByVal colorDepth As Long = 24, Optional ByVal initialColor As Long = vbWhite, Optional ByVal initialAlpha As Long = 0) As Boolean
    
    On Error GoTo CouldNotCreateDIB
    
    Dim cPainter As pd2DPainter, cSurface As pd2DSurface, cBrush As pd2DBrush
                
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
                FillMemory m_dibBits, m_dibStride * m_dibHeight, 0&
            Else
                
                If (initialColor <> vbWhite) And (initialColor <> vbBlack) Then initialColor = TranslateColor(initialColor)
                
                Drawing2D.QuickCreatePainter cPainter
                Drawing2D.QuickCreateSurfaceFromDC cSurface, Me.GetDIBDC, False
                Drawing2D.QuickCreateSolidBrush cBrush, initialColor, initialAlpha / 2.55
                cPainter.FillRectangleF cSurface, cBrush, 0, 0, imgWidth + 1, imgHeight + 1
                Set cSurface = Nothing
                
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
            
            'Create a DIB
            m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, 0, 0)
            
            'If successful, select the newly created dib into our DC
            If (m_dibHandle <> 0) Then
                
                'Inside debugging mode, increment a global DIB counter.
                #If DEBUGMODE = 1 Then
                    If (m_dibHandle <> 0) Then g_DIBsCreated = g_DIBsCreated + 1
                #End If
                
                'We will later use m_dibHandleOriginal to clear up the memory associated with this DIB
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
                        
                        'The back color may or may not be a system color, so translate it just in case
                        If (initialColor <> vbWhite) Then initialColor = TranslateColor(initialColor)
                        
                        Drawing2D.QuickCreatePainter cPainter
                        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_dibDC, False
                        Drawing2D.QuickCreateSolidBrush cBrush, initialColor, initialAlpha / 2.55
                        cPainter.FillRectangleF cSurface, cBrush, 0, 0, imgWidth + 1, imgHeight + 1
                        Set cSurface = Nothing
                        
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
                Me.EraseDIB
            End If
            
            'Our parent may not use this DIB right away, so free it from its DC; the DC will be auto-created when
            ' it's next requested.
            Me.FreeFromDC
            
        End If
        
    End If
    
CouldNotCreateDIB:
    
    'Return success contingent on whether we have a DIB pointer or not
    CreateBlank = CBool(m_dibHandle <> 0)
    
End Function

'Replace the entire DIB's contents with some solid color
Public Sub FillWithColor(Optional ByVal fillColor As Long = vbBlack, Optional ByVal fillAlpha As Single = 0#)

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
        
        'Just like the 24-bpp block above, we can chat with black and white values in 32-bpp mode too - BUT ONLY
        ' if the alpha byte value matches.
        Dim fillAlphaL As Long
        fillAlphaL = fillAlpha * 2.55
        
        If ((fillColor = vbWhite) And (fillAlphaL = 255)) Then
            Me.ResetDIB 255
        Else
            
            Me.ResetDIB 0
            
            If ((fillColor <> vbBlack) And (fillAlphaL <> 0)) Then
                
                Dim tmpSurface As Long, tmpBrush As Long
                tmpSurface = GDI_Plus.GetGDIPlusGraphicsFromDC(Me.GetDIBDC)
                tmpBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(fillColor, fillAlpha * 2.55)
                GDI_Plus.GDIPlus_GraphicsSetCompositingMode tmpSurface, GP_CM_SourceCopy
                GDI_Plus.GDIPlus_FillRectI tmpSurface, tmpBrush, -1, -1, m_dibWidth + 2, m_dibHeight + 2
                GDI_Plus.ReleaseGDIPlusBrush tmpBrush
                GDI_Plus.ReleaseGDIPlusGraphics tmpSurface
                
                m_IsAlphaPremultiplied = CBool(fillAlphaL = 255)
                
            Else
                m_IsAlphaPremultiplied = True
            End If
            
        End If
        
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub

'This will effectively reset everything related to this DIB, including all image data. Use cautiously!
Public Sub EraseDIB(Optional ByVal alsoReleaseDC As Boolean = False)

    'If we have image data, clear it out
    If (m_dibHandle <> 0) Then
        
        If (m_dibDC <> 0) Then SelectObject m_dibDC, m_dibHandleOriginal
        DeleteObject m_dibHandle
        m_dibHandle = 0
        
        'Inside debug mode, decrement a global DIB counter
        #If DEBUGMODE = 1 Then
            g_DIBsDestroyed = g_DIBsDestroyed + 1
        #End If
    
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
    
End Sub

'Maintain the DIB's dimensions, but reset all bytes to some value (typically 0).  FillMemory is used for maximum performance.
Public Sub ResetDIB(Optional ByVal fillValue As Byte = 0)

    'Retrieve a pointer to the DIB array and the array's current size
    Dim tmpDIBPointer As Long, tmpDIBSize As Long
    RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBSize
    
    'Erase everything
    If (tmpDIBPointer <> 0) And (tmpDIBSize > 0) Then FillMemory tmpDIBPointer, tmpDIBSize, fillValue
    
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
    
    'Initialize the ICC profile manager
    Set ICCProfile = New pdICCProfile
    
End Sub

'TERMINATE class
Private Sub Class_Terminate()
    
    EraseDIB True
    
    'Release the ICC profile manager
    Set ICCProfile = Nothing
    
End Sub

'Draw this DIB to a picture box. The image will be automatically centered and sized to fit.
Public Sub RenderToPictureBox(ByRef dstPicture As PictureBox, Optional ByVal colorManagementMatters As Boolean = True, Optional ByVal doNotStretchIfSmaller As Boolean = False, Optional ByVal suspendTransparencyGrid As Boolean = False)

    'Erase any existing picture
    If dstPicture.AutoRedraw Then
        dstPicture.Picture = LoadPicture("")
    Else
        dstPicture.Cls
    End If
    
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
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
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0) Then srcAspect = srcWidth / srcHeight Else srcAspect = 1
    If (dstHeight > 0) Then dstAspect = dstWidth / dstHeight Else dstAspect = 1
    
    Dim dWidth As Long, dHeight As Long, previewX As Long, previewY As Long
    ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, dWidth, dHeight
    
    If fitPrevented Then
        previewX = (dstPicture.ScaleWidth - srcWidth) / 2
        previewY = (dstPicture.ScaleHeight - srcHeight) / 2
    Else
        If (srcAspect > dstAspect) Then
            previewY = CLng((dstHeight - dHeight) / 2)
            previewX = 0
        Else
            previewX = CLng((dstWidth - dWidth) / 2)
            previewY = 0
        End If
    End If
        
    'We now branch according to active color management.  This is a hackish fix designed to work around some current
    ' performance issues with WCS; I'm working on an LittleCMS-based fix for the 7.0 release.
    If colorManagementMatters Then
        
        'Create a temporary DIB to house the stretched image.  Because the stretched image is typically smaller than
        ' the original image, this reduces color management time.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank dWidth, dHeight, m_dibColorDepth, 0, 0
        
        If (m_dibColorDepth = 24) Then
        
            If (dstPicture.ScaleWidth < m_dibWidth) Or (dstPicture.ScaleHeight < m_dibHeight) Then
                SetStretchBltMode tmpDIB.GetDIBDC, STRETCHBLT_HALFTONE
            Else
                SetStretchBltMode tmpDIB.GetDIBDC, STRETCHBLT_COLORONCOLOR
            End If
            
            StretchBlt tmpDIB.GetDIBDC, 0, 0, dWidth, dHeight, Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, vbSrcCopy
        
        Else
            GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, dWidth, dHeight, Me, 0, 0, m_dibWidth, m_dibHeight
        End If
        
        'Apply display color management now
        ColorManagement.ApplyDisplayColorManagement tmpDIB
        
        'Render the image to the picture box
        If (m_dibColorDepth = 24) Then
            BitBlt dstPicture.hDC, previewX, previewY, dWidth, dHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        Else
            If (Not suspendTransparencyGrid) Then GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, previewX, previewY, dWidth, dHeight, g_CheckerboardPattern, dstPicture.hDC
            GDI_Plus.GDIPlus_StretchBlt Nothing, previewX, previewY, dWidth, dHeight, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, , , dstPicture.hDC
        End If
        
    'If color management is *not* required, we don't need a temporary DIB - just render directly onto the target.
    Else
        
        'For 24bpp images, simply copy over the current image data
        If (m_dibColorDepth = 24) Then
        
            'If the target area is smaller than the DIB, request halftoning
            If (dstPicture.ScaleWidth < m_dibWidth) Or (dstPicture.ScaleHeight < m_dibHeight) Then
                SetStretchBltMode dstPicture.hDC, STRETCHBLT_HALFTONE
            Else
                SetStretchBltMode dstPicture.hDC, STRETCHBLT_COLORONCOLOR
            End If
            
            StretchBlt dstPicture.hDC, previewX, previewY, dWidth, dHeight, Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, vbSrcCopy
        
        'For 32bpp images, we must first apply a background checkerboard pattern, then alpha-blend the resized
        ' image onto it.
        Else
            If (Not suspendTransparencyGrid) Then GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, previewX, previewY, dWidth, dHeight, g_CheckerboardPattern, dstPicture.hDC
            GDI_Plus.GDIPlus_StretchBlt Nothing, previewX, previewY, dWidth, dHeight, Me, 0, 0, m_dibWidth, m_dibHeight, , , dstPicture.hDC
        End If
        
    End If
        
    If dstPicture.AutoRedraw Then
        dstPicture.Picture = dstPicture.Image
        dstPicture.Refresh
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
    
End Sub

'Load a DIB's header and image information from file.  If the file is compressed, it will automatically be decompressed.
' Optionally, if the DIB data is embedded inside another file, that handle can be passed to the this function, and it will attempt to
' load DIB data from the current file pointer location.  (In this case, just supply a blank filename.)
Public Function CreateFromFile(ByRef srcFilename As String, Optional ByVal useEmbeddedLocation As Boolean = False, Optional ByVal hFile As Long = 0) As Boolean
    
    'pdFSO is used so we can support Unicode paths
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If this is not an embedded file, create and open it now
    If (Not useEmbeddedLocation) Then
        If (Not cFile.FileCreateHandle(srcFilename, hFile, True, False, OptimizeSequentialAccess)) Then
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdDIB.createFromFile failed to generate a handle for " & srcFilename & ".  Load abandoned."
            #End If
            CreateFromFile = False
            Exit Function
        End If
    End If
    
    'Check to make sure this file actually contains DIB data
    Dim DIBIDCheck As Long
    cFile.FileReadData hFile, VarPtr(DIBIDCheck), 4&
        
    If (DIBIDCheck <> DIB_IDENTIFIER) Then
        Message "Failed to load embedded image; invalid header found."
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
        localAlphaPremultiplied = CBool(tmpAlphaCheck <> 0)
        
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
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdDIB.CreateFromFile encountered invalid width/height measurements (" & fWidth & "x" & fHeight & ", scanline width: " & fArrayWidth & ")"
        #End If
    Else
        #If DEBUGMODE = 1 Then
            'pdDebug.LogAction "pdDIB.CreateFromFile() detected an incoming buffer with measurements (" & fWidth & "x" & fHeight & ", scanline width: " & fArrayWidth & ")"
        #End If
    End If
    
    'Check compression.  As of 2016, this is no longer a true/false value - it is an enum describing a specific
    ' PD compression engine.
    Dim useDecompressionEngine As PD_COMPRESSION_ENGINES
    useDecompressionEngine = PD_CE_NoCompression
    
    If (DIBVersionCheck = DIB_FILE_VERSION_2016) Then
        cFile.FileReadData hFile, VarPtr(useDecompressionEngine), 4&
    Else
        Dim toDecompress As Boolean
        cFile.FileReadData hFile, VarPtr(toDecompress), 2&
        If toDecompress Then useDecompressionEngine = PD_CE_ZLib Else useDecompressionEngine = PD_CE_NoCompression
    End If
    
    Dim fileData() As Byte
    
    'If the data *is* compressed, we need to retrieve two additional pieces of information: the original and compressed size
    ' of the entire pixel chunk.
    Dim origDataSize As Long, compressedDataSize As Long
    
    If (useDecompressionEngine <> PD_CE_NoCompression) Then
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
        
        'If the DIB data is compressed, decompress it now.  Note that we decompress directly into the DIB itself, to avoid the
        ' need for a costly CopyMemory call (and extra DIB-sized allocation).
        If (useDecompressionEngine <> PD_CE_NoCompression) Then
            
            'If the data is decompressed successfully, convert it to a 2D array and copy it into this class's DIB memory chunk
            If (Not Compression.DecompressPtrToPtr(tmpDIBPointer, tmpDIBLength, VarPtr(fileData(0)), compressedDataSize, useDecompressionEngine)) Then
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Decompression was not successful (unspecified compression plugin error).  DIB processing abandoned."
                #End If
                CreateFromFile = False
                Exit Function
            End If
        
        'If the DIB data is not compressed, copy it straight into the DIB
        Else
            cFile.FileReadData hFile, tmpDIBPointer, tmpDIBLength
        End If
        
        CreateFromFile = True
        
    Else
        #If DEBUGMODE = 1 Then
            If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.createFromFile failed to create a new DIB using the stored file settings.  Load abandoned."
        #End If
        CreateFromFile = False
    End If
    
    'Reset alpha premultiplication to match the value embedded in the file
    m_IsAlphaPremultiplied = localAlphaPremultiplied
    
    'If this is an embedded file, leave the file open for further processing.  Otherwise, close it.
    If (Not useEmbeddedLocation) Then cFile.FileCloseHandle hFile
    
End Function

'Write this class's DIB information to file.  Compression can optionally be applied to the DIB data.
' Also, if the calling function wants to embed the DIB's data inside an existing file, specify the embedInFile parameter and pass a valid
' file handle, with the pointer already moved to the location desired for embedding.
Public Function WriteToFile(ByRef dstFilename As String, Optional ByVal useCompressionEngine As PD_COMPRESSION_ENGINES = PD_CE_Lz4, Optional ByVal embedInFile As Boolean = False, Optional ByRef hFile As Long = 0) As Boolean
    
    'Before proceeding, make sure this DIB actually exists
    If (Me.GetDIBWidth <= 0) Or (Me.GetDIBHeight <= 0) Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdDIB.WriteToFile was called on an empty DIB.  Fix this!"
        #End If
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
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdDIB.WriteToFile failed to create a file handle for " & dstFilename & ".  Write abandoned."
            #End If
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
    
    'And finally, the pixel data.  Start by retrieving two key items of business: a pointer to this DIB's raw pixel data,
    ' and the total length of that pixel data.
    Dim tmpDIBPointer As Long, tmpDIBLength As Long
    Me.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
        
    'If compression has been requested, and zLib is available, we'll shrink the DIB data down before writing it.
    Dim dataWasCompressed As Boolean: dataWasCompressed = False
    If (useCompressionEngine <> PD_CE_NoCompression) Then
        
        'Before writing anythign to file, attempt compression and make sure it works
        Dim compressedImageData() As Byte, compressedSize As Long
        dataWasCompressed = Compression.CompressPtrToDstArray(compressedImageData, compressedSize, tmpDIBPointer, tmpDIBLength, useCompressionEngine)
        
        If dataWasCompressed Then
            
            cFile.FileWriteData hFile, VarPtr(useCompressionEngine), 4&
                
            'Write the original size of the array, followed by the compressed size, followed by the image data itself.  (The array sizes will
            ' help us automatically create intelligently-sized buffers at decompression time.)
            cFile.FileWriteData hFile, VarPtr(tmpDIBLength), 4&
            cFile.FileWriteData hFile, VarPtr(compressedSize), 4&
            cFile.FileWriteData hFile, VarPtr(compressedImageData(0)), compressedSize
            
            #If DEBUGMODE = 1 Then
                'pdDebug.LogAction "pdDIB.WriteToFile: (" & CStr(Me.GetDIBWidth) & "x" & CStr(Me.GetDIBHeight) & "x" & CStr(Me.GetDIBColorDepth) & "): " & CStr(compressedSize) & " bytes of pixel data (compressed; original was " & CStr(tmpDIBLength) & " bytes)"
            #End If
                
        'On failure, we'll simply write out uncompressed bytes
        Else
            useCompressionEngine = PD_CE_NoCompression
        End If
        
    End If
    
    'If compression has not been requested, or zLib is unavailable, we can dump the raw image data to file as-is.
    If (Not dataWasCompressed) Then
        cFile.FileWriteData hFile, VarPtr(useCompressionEngine), 4&
        cFile.FileWriteData hFile, tmpDIBPointer, tmpDIBLength
    End If
        
    'If this is not an embedding, close the file now
    If (Not embedInFile) Then cFile.FileCloseHandle hFile

End Function

'Write this class's DIB information to a valid BMP file.
Public Sub WriteToBitmapFile(ByRef dstFilename As String)

    'First, make sure this DIB actually contains image data
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
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    If cFile.FileExists(dstFilename) Then cFile.FileDelete dstFilename
    
    'Writing the file is easy; just dump the two headers and image data into it!
    
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

        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdDIB.writeToBitmapFile successfully wrote " & dstFilename & "."
        #End If

    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdDIB.writeToBitmapFile failed to create a valid file handle for " & dstFilename & ".  Write abandoned."
        #End If
    End If
    
    'Minimize GDI resources by freeing our DC
    Me.FreeFromDC
        
End Sub

'As a helper to external functions, this sub can be used to fill a destination array with a copy the image's current data.
' Note that this is raw image bytes - no header whatsoever - so it's assumed the caller knows how to deal with the data.
Public Sub CopyImageBytesIntoStream(ByRef dstArray() As Byte)

    'Retrieve a pointer to the current DIB bits (in VB format).  Technically we could just use the m_dibBits pointer
    ' and calculate bounds manually, but this provides a foolproof way to do the same thing using existing code.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Using that temporary as our source, convert the image to the supplied 1D array.
    Dim aSize As Long
    aSize = (UBound(iData, 1) + 1) * (UBound(iData, 2) + 1) - 1
    ReDim dstArray(0 To aSize) As Byte
    CopyMemory dstArray(0), iData(0, 0), aSize
    
    CopyMemory ByVal VarPtrArray(iData), 0&, 4

End Sub

'Shortcut to copyImageBytesIntoStream; this function gives a pointer to DIB byte 0, and the size of the DIB's contents.  This can be used
' to mimic byte array behavior, without the overhead of actually initializing a VB array and copying the DIB into it.
Public Sub RetrieveDIBPointerAndSize(ByRef dibPointer As Long, ByRef dibSize As Long)
    dibPointer = m_dibBits
    dibSize = m_dibStride * m_dibHeight
End Sub

'The counterpart to copyImageBytesIntoStream, above.  This function will blindly copy an array of bytes over the current
' image's data.  For it to work, the DIB must have been correctly created at the right size and color depth BEFORE CALLING
' THIS FUNCTION.  Otherwise, you will get errors.
Public Sub CopyStreamOverImageArray(ByRef srcArray() As Byte)
    CopyMemory ByVal m_dibBits, srcArray(0), UBound(srcArray) + 1
End Sub

'Sometimes this class needs to access its own DIB bits. Here's how.
Private Sub PrepInternalSafeArray(ByRef dstSafeArray As SAFEARRAY2D)
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

Private Sub PrepInternalSafeArray_Scanline(ByRef dstSafeArray As SAFEARRAY1D, ByVal dstLine As Long)
    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_dibBits + (dstLine * m_dibStride)
    End With
End Sub

Private Sub PrepInternalLongSafeArray(ByRef dstSafeArray As SAFEARRAY2D)
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

Private Sub PrepInternalLongSafeArray_Scanline(ByRef dstSafeArray As SAFEARRAY1D, ByVal dstLine As Long)
    With dstSafeArray
        .cbElements = 4
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = m_dibWidth
        .pvData = m_dibBits + (dstLine * m_dibStride)
    End With
End Sub

'2D array access is slow, so PD tries to use 1D arrays when possible.  This function initializes a base SafeArray, but the caller
' will need to manually update it to point at a desired line.
Private Sub PrepInternalSafeArray1D(ByRef dstSafeArray As SAFEARRAY1D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_dibBits
    End With
End Sub

'Sometimes this class needs to access the DIB bits of other DIBs. Here's how.
Private Sub PrepExternalSafeArray(ByRef srcDIB As pdDIB, ByRef dstSafeArray As SAFEARRAY2D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBStride
        .pvData = srcDIB.GetDIBPointer
    End With
End Sub

'Pre-composite an image with an alpha-channel against a background color. Until PhotoDemon is capable of rendering transparent
' images itself, this is necessary to give transparent images a white background.
Public Sub CompositeBackgroundColor(Optional ByVal newR As Byte = 255, Optional ByVal newG As Byte = 255, Optional ByVal newB As Byte = 255)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then Exit Sub
    
'    #If DEBUGMODE = 1 Then
'        Debug.Print "Compositing DIB of size " & m_dibWidth & "x" & m_dibHeight & ".  Alpha premultiplication = " & m_IsAlphaPremultiplied & "."
'    #End If
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
    PrepInternalSafeArray1D dibSA
    
    Dim x As Long, y As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'Because our alpha values are pre-multiplied, we can composite them against the background color via use of a look-up table.
    Dim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    'Populate a unique lookup table for each color, based on each possible alpha value (0 to 255)
    For x = 0 To 255
        tmpAlpha = 1 - (x / 255)
        rLookup(x) = Int(newR * tmpAlpha)
        gLookUp(x) = Int(newG * tmpAlpha)
        bLookup(x) = Int(newB * tmpAlpha)
    Next x
    
    'Loop through the image, compositing as we go
    For y = 0 To finalY
        
        'Point our 1D pixel array at the proper scanline
        dibSA.pvData = dibPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(dibPixels()), VarPtr(dibSA), 4
        
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
                    dibPixels(x + 1) = dibPixels(x + 1) + gLookUp(checkAlpha)
                    dibPixels(x + 2) = dibPixels(x + 2) + rLookup(checkAlpha)
                Else
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = checkAlpha / 255
                
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
    CopyMemory ByVal VarPtrArray(dibPixels), 0&, 4

End Sub

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function Blend2Colors(ByVal Color1 As Long, ByVal Color2 As Long, ByRef mixRatio As Double) As Byte
    Blend2Colors = mixRatio * (Color1 - Color2) + Color2
End Function

'Quick and dirty function to reverse scanlines in the DIB
Public Sub ReverseScanlines(Optional ByVal updateTopDownTracker As Boolean = True)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalY As Long
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim tmpScanLine() As Byte
    ReDim tmpScanLine(0 To scanlineSize - 1) As Byte
    
    Dim srcPointer As Long, dstPointer As Long
    
    'Loop through the image, converting alpha as we go
    Dim y As Long
    For y = 0 To (finalY \ 2)
        
        'Copy the source line to a temporary array
        dstPointer = VarPtr(tmpScanLine(0))
        srcPointer = dibPointer + scanlineSize * (finalY - y)
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
        'Copy the destination line over the source line
        dstPointer = srcPointer
        srcPointer = dibPointer + scanlineSize * y
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
        'Copy the temporary array into the destination line
        dstPointer = srcPointer
        srcPointer = VarPtr(tmpScanLine(0))
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
    Next y
    
    If updateTopDownTracker Then m_dibHeader.Height = m_dibHeader.Height * -1

End Sub

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
Public Sub CopyAlphaFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal binaryAlphaOnly As Boolean = False)
    
    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then Exit Sub
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim srcScanlineSize As Long, srcPointer As Long, dstScanlineSize As Long, dstPointer As Long
    dstScanlineSize = Me.GetDIBStride
    dstPointer = Me.GetDIBPointer
    srcScanlineSize = srcDIB.GetDIBStride
    srcPointer = srcDIB.GetDIBPointer
    
    Dim dstPixels() As Byte, dstSA As SAFEARRAY1D
    PrepInternalSafeArray1D dstSA
    
    Dim srcPixels() As Byte, srcSA As SAFEARRAY1D
    PrepInternalSafeArray1D srcSA
    srcSA.cElements = srcScanlineSize
    srcSA.pvData = srcPointer
    
    Dim x As Long, y As Long
    
    'Loop through the image, compositing as we go
    For y = 0 To finalY
        
        'Point our 1D pixel arrays at the proper scanlines
        dstSA.pvData = dstPointer + dstScanlineSize * y
        CopyMemory ByVal VarPtrArray(dstPixels()), VarPtr(dstSA), 4
        
        srcSA.pvData = srcPointer + srcScanlineSize * y
        CopyMemory ByVal VarPtrArray(srcPixels()), VarPtr(srcSA), 4
        
        'Split handling based on premultiplication
        If binaryAlphaOnly Then
            
            For x = 0 To finalX Step 4
                If srcPixels(x + 3) = 255 Then
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
    CopyMemory ByVal VarPtrArray(dstPixels), 0&, 4
    CopyMemory ByVal VarPtrArray(srcPixels), 0&, 4
    
    If binaryAlphaOnly Then Me.SetInitialAlphaPremultiplicationState True
    
End Sub

'DIBs created from GDI+ contain pre-multiplied alpha values. These are great for high-performance blting, but terrible for
' photo editing. This routine can be used to either apply or remove premultiplied alpha from an image.
'
'Optionally, you can also supply a pointer to a valid RectF struct; if present, only this region will be updated.
Public Sub SetAlphaPremultiplication(Optional ByVal applyPremultiplication As Boolean = False, Optional ByVal ignoreEmbeddedValue As Boolean = False, Optional ByVal ptrToRectF As Long = 0)

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
        
        Dim tmpRectF As RECTF
        CopyMemory_Strict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
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
    
        'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
        ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
        Dim intToFloat() As Single
        ReDim intToFloat(0 To 255) As Single
        Dim i As Long
        For i = 0 To 255
            If applyPremultiplication Then
                intToFloat(i) = i / 255
            Else
                If i <> 0 Then intToFloat(i) = 255 / i
            End If
        Next i
        
        '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
        ' 1D array, which we reset between scanlines.
        Dim scanlineSize As Long, dibPointer As Long
        scanlineSize = Me.GetDIBStride
        dibPointer = Me.GetDIBPointer
        
        Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
        PrepInternalSafeArray1D dibSA
        
        Dim x As Long, y As Long
        Dim r As Long, g As Long, b As Long
        Dim tmpAlpha As Byte, tmpAlphaModifier As Double
        
        'Loop through the image, converting alpha as we go
        For y = initY To finalY
            
            'Point our 1D pixel array at the proper scanline
            dibSA.pvData = dibPointer + scanlineSize * y
            CopyMemory ByVal VarPtrArray(dibPixels()), VarPtr(dibSA), 4
            
        For x = initX To finalX Step 4
            
            'Retrieve alpha for the current pixel
            tmpAlpha = dibPixels(x + 3)
            
            'Branch according to applying or removing premultiplication
            If applyPremultiplication Then
            
                'When applying premultiplication, we can ignore fully opaque pixels
                If (tmpAlpha <> 255) Then
                
                    'We can shortcut the calculation of full transparent pixels (they are made black)
                    If tmpAlpha = 0 Then
                        dibPixels(x) = 0
                        dibPixels(x + 1) = 0
                        dibPixels(x + 2) = 0
                    Else
                
                        b = dibPixels(x)
                        g = dibPixels(x + 1)
                        r = dibPixels(x + 2)
                        
                        tmpAlphaModifier = intToFloat(tmpAlpha)
                        
                        'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                        r = (r * tmpAlphaModifier)
                        g = (g * tmpAlphaModifier)
                        b = (b * tmpAlphaModifier)
                        
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
                        r = (r * tmpAlphaModifier)
                        g = (g * tmpAlphaModifier)
                        b = (b * tmpAlphaModifier)
                        
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
        CopyMemory ByVal VarPtrArray(dibPixels), 0&, 4
        
        'Mark the new premultiplication state
        m_IsAlphaPremultiplied = applyPremultiplication
        
    End If
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub AlphaBlendToDC(ByVal dstDC As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
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
Public Sub AlphaBlendToDCEx(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customAlpha As Long = 255)
    
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
