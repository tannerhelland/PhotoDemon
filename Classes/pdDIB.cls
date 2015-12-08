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
'Copyright 2012-2015 by Tanner Helland
'Created: 29/August/12
'Last updated: 28/October/15
'Last update: if we've already created a DC for this class, attempt to reuse it instead of deleting it and creating a new one
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
Private Type BITMAP
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPerPixel As Integer
    Bits As Long
End Type

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

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Private Type SAFEARRAYBOUND
    cElements As Long
    lBound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type


'Drawing API functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCHBLT_COLORONCOLOR As Long = 3
Private Const STRETCHBLT_HALFTONE As Long = 4

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DIB API functions
' (Note that these are currently declared in FastDrawing as well)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'Object API functions
Private Const OBJ_BITMAP As Long = 7
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Rectangle objects to be used with brushes
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'Brush creation
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal HPALETTE As Long, ByRef cColorRef As Long) As Long

'AlphaBlend API call
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'Variables related to the DIB
'hDC for this DIB
Private m_dibDC As Long
'Actual DIB handle for this DIB
Private m_dibHandle As Long
'Original handle when this DIB is first created (we must store this so we can properly clean up the DIB when we're finished)
Private m_dibHandleOriginal As Long
'Pointer to the actual DIB bits
Private m_dibBits As Long
'Persistent DIB header; this will be used with WAPI to create the DIB initially
Private m_dibHeader As BITMAPINFOHEADER

'The DIB's width and height
Private m_dibWidth As Long, m_dibHeight As Long

'The DIB's array width (m_dibWidth * 4 for 32-bit, varies for 24-bit due to DWORD-alignment)
Private m_dibArrayWidth As Long

'The DIB's color depth (should only ever be 24 or 32)
Private m_DibColorDepth As Long

'Used when writing/reading DIB data to/from a file
Private Const DIB_IDENTIFIER As Long = &H726C4450  'PDlr
Private Const DIB_FILE_VERSION_2012 As Long = &H1000&
Private Const DIB_FILE_VERSION_2015 As Long = &H1001&

'When converting a 32bpp DIB to 8bpp with transparency, we need to know the location of a transparent pixel.
' This location is passed to FreeImage, and it designates the palette index at that location as the transparent index.
Private m_trnsX As Long, m_trnsY As Long

'Remember the original color of the pixels made transparent, then restore it after the "magic magenta" has served its purpose
Private m_OriginalTransparentColor As Long

'If this DIB was loaded from file, its original file format will be set here.  (By default, this value will be -1).
Private m_OriginalFIF As Long

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
Public Function getAlphaPremultiplication() As Boolean
    getAlphaPremultiplication = m_IsAlphaPremultiplied
End Function

Public Sub setInitialAlphaPremultiplicationState(ByVal newState As Boolean)
    m_IsAlphaPremultiplied = newState
End Sub

'Get/set original color depth.  Note that this is set after tone-mapping and other actions have been applied
Public Function getOriginalColorDepth() As Long
    getOriginalColorDepth = m_OriginalColorDepth
End Function

Public Sub setOriginalColorDepth(ByVal origColorDepth As Long)
    m_OriginalColorDepth = origColorDepth
End Sub

'Get/set original file color depth.  This is set prior to any tone-mapping or explicit FreeImage conversions.  As such, it may be quite large.
Public Function getOriginalFreeImageColorDepth() As Long
    getOriginalFreeImageColorDepth = m_OriginalFreeImageColorDepth
End Function

Public Sub setOriginalFreeImageColorDepth(ByVal origFIColorDepth As Long)
    m_OriginalFreeImageColorDepth = origFIColorDepth
End Sub

'If this DIB was created from an image file, and that file specified a preferred background color, use this
' function to copy the value into this DIB.
Public Sub setBackgroundColor(ByVal newBackColor As Long)
    m_OriginalBackgroundColor = newBackColor
End Sub

Public Function getBackgroundColor() As Long
    getBackgroundColor = m_OriginalBackgroundColor
End Function

'Set the DPI of this DIB.  This is only relevant if this DIB has been created directly from an image file.
Public Sub setDPI(ByVal xRes As Double, ByVal yRes As Double, Optional ByVal sourceIsFreeImage As Boolean = False)
    
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
Public Function getDPI() As Double
    getDPI = m_dibDPI
End Function

'Get/Set original DIB format
Public Function getOriginalFormat() As Long
    getOriginalFormat = m_OriginalFIF
End Function

Public Sub setOriginalFormat(ByVal origFIF As Long)
    m_OriginalFIF = origFIF
End Sub

'Return the original transparent color (set as part of the MarkAlpha routine).
Public Function getOriginalTransparentColor() As Long
    getOriginalTransparentColor = m_OriginalTransparentColor
End Function

'Return the location of a transparent pixel. This will be set as part of the applyAlphaCutoff routine.
Public Sub getTransparentLocation(ByRef srcX As Long, ByRef srcY As Long)
    srcX = m_trnsX
    srcY = m_trnsY
End Sub

'Given a pixel coordinate, return an RGBA quad for that coordinate.
' Returns: TRUE if the pixel lies inside DIB boundaries; FALSE otherwise.  Make sure to check this before using the RGBQUAD.
Friend Function GetPixelRGBQuad(ByVal x As Long, ByVal y As Long, ByRef dstQuad As RGBQUAD) As Boolean

    'Before doing anything else, check to see if the x/y postition lies inside the DIB
    If (x >= 0) And (x < m_dibWidth) And (y >= 0) And (y < m_dibHeight) Then
        
        'The point lies inside the DIB, which means we need to figure out the color at this position
        GetPixelRGBQuad = True
        
        Dim tmpData() As Byte
        Dim tSA As SAFEARRAY2D
        prepInternalSafeArray tSA
        CopyMemory ByVal VarPtrArray(tmpData()), VarPtr(tSA), 4&
        
        Dim QuickX As Long
        QuickX = x * (m_DibColorDepth \ 8)
        
        'Failsafe bounds check
        If ((QuickX + 2) < m_dibArrayWidth) Then
        
            With dstQuad
                .Red = tmpData(QuickX + 2, y)
                .Green = tmpData(QuickX + 1, y)
                .Blue = tmpData(QuickX, y)
                If m_DibColorDepth = 32 Then .Alpha = tmpData(QuickX + 3, y)
            End With
            
        End If
        
        CopyMemory ByVal VarPtrArray(tmpData), 0&, 4
    
    'This coordinate does not lie inside the layer.
    Else
        GetPixelRGBQuad = False
    End If

End Function

'This function will apply an "alpha-cutoff" to a 32bpp image. Any alpha values above the cutoff will be replaced with
' magic magenta (253,0,253) and the location of a transparent pixel (m_trnsX, m_trnsY) will be marked.
Public Sub applyAlphaCutoff(Optional ByVal cutOff As Long = 127, Optional ByVal forceColor As Boolean = True, Optional ByVal newBackgroundColor As Long = 16777215)

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If m_DibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (m_dibDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            If cutOff = 0 Then
                convertTo24bpp newBackgroundColor
                Exit Sub
            End If
            
            'Loop through the image and force each alpha value to 0 and 255 based on the supplied cutoff.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
            
            Dim chkAlpha As Byte
            Dim tmpAlpha As Double
            
            m_trnsX = -1
            
            'Retrieve RGB values from the new background color, which we'll use to composite semi-transparent pixels
            Dim backR As Byte, backG As Byte, backB As Byte
            backR = CByte(ExtractR(newBackgroundColor))
            backG = CByte(ExtractG(newBackgroundColor))
            backB = CByte(ExtractB(newBackgroundColor))
                
            'Loop through the image, checking alphas as we go
            For x = 0 To m_dibWidth - 1
                QuickX = x * 4
            For y = 0 To m_dibHeight - 1
                
                chkAlpha = iData(QuickX + 3, y)
                
                'If the alpha value is less than the cutoff, mark this pixel
                If chkAlpha < cutOff Then
                
                    'Remember this location if we haven't already
                    If m_trnsX = -1 Then
                        m_trnsX = x
                        m_trnsY = y
                        m_OriginalTransparentColor = RGB(iData(QuickX + 2, y), iData(QuickX + 1, y), iData(QuickX, y))
                    End If
                    
                    'When previewing this effect, we don't actually want to change pixel colors. Otherwise,
                    ' HALFTONE-mode resampling may bleed the magenta into the preview and give poor results.
                    If forceColor Then
                        'Color this pixel with magic magenta (253,0,253)
                        iData(QuickX + 2, y) = 253
                        iData(QuickX + 1, y) = 0
                        iData(QuickX, y) = 253
                    End If
                    
                    'Blank the alpha channel
                    iData(QuickX + 3, y) = 0
                    
                'If the pixel is not beneath the cut-off, and not fully opaque, composite it against white
                ElseIf chkAlpha < 255 Then
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = CDbl(chkAlpha) / 255
            
                    'Use that alpha value to blend the current color against white
                    iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), backR, tmpAlpha)
                    iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), backG, tmpAlpha)
                    iData(QuickX, y) = Blend2Colors(iData(QuickX, y), backB, tmpAlpha)
                    
                    'Opaque-ify the alpha channel
                    iData(QuickX + 3, y) = 255
                
                End If
                
            Next y
                
            Next x
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            Erase iData
                
        End If
        
    End If
    
End Sub

'Force all alpha bytes in the DIB to some preset value.  If the value is *not* 255, please remember to premultiply
' alpha accordingly, if the DIB will be rendered to the screen.
Public Sub ForceNewAlpha(ByVal newAlpha As Byte)

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If m_DibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (m_dibDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            'Loop through the image and force each alpha value to the user's specified value
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
                
            'Loop through the image, checking alphas as we go
            For x = 0 To m_dibWidth - 1
                QuickX = x * 4
            For y = 0 To m_dibHeight - 1
                iData(QuickX + 3, y) = newAlpha
            Next y
            Next x
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
                
        End If
        
    End If
    
End Sub

'Translate an OLE color to an RGB Long
Private Function TranslateColor(ByVal colorRef As Long) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, TranslateColor) Then
        TranslateColor = RGB(255, 255, 255)
    End If
End Function

'Return this DIB's color depth
Public Function getDIBColorDepth() As Long
    getDIBColorDepth = m_DibColorDepth
End Function

'Return this DIB's array width
Public Function getDIBArrayWidth() As Long
    getDIBArrayWidth = m_dibArrayWidth
End Function

'Return this DIB's width
Public Function getDIBWidth() As Long
    getDIBWidth = m_dibWidth
End Function

'Return this DIB's height
Public Function getDIBHeight() As Long
    getDIBHeight = m_dibHeight
End Function

'Return whether or not this DIB has image data associated with it
Public Function hasImage() As Boolean
    hasImage = (m_dibHandle <> 0)
End Function

'Return this DIB's hDC
Public Function getDIBDC() As Long
    getDIBDC = m_dibDC
End Function

'Return a pointer to this DIB's actual DIB
Public Function getDIBHandle() As Long
    getDIBHandle = m_dibHandle
End Function

'Return a pointer to this DIB's pixel data (no header)
Public Function getActualDIBBits() As Long
    getActualDIBBits = m_dibBits
End Function

'Return a pointer to this DIB's header
Public Function getDIBHeader() As Long
    getDIBHeader = VarPtr(m_dibHeader)
End Function

'Make a copy of an existing DIB
Public Function createFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal newWidth As Long = -1, Optional ByVal newHeight As Long = -1, Optional useHalftoning As Boolean = True) As Boolean

    'Make sure the DIB we're passed isn't empty
    If srcDIB.getDIBDC <> 0 Then
        
        'Prepare new width and height values as requested by the user
        If newWidth = -1 Then newWidth = srcDIB.getDIBWidth
        If newHeight = -1 Then newHeight = srcDIB.getDIBHeight
        
        'If the width and height values are not being changed, the transfer is simple
        If (newWidth = srcDIB.getDIBWidth) And (newHeight = srcDIB.getDIBHeight) Then
            
            'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
            If (newWidth = m_dibWidth) And (newHeight = m_dibHeight) And (srcDIB.getDIBColorDepth = m_DibColorDepth) Then
            
                'Copy the image data without modification
                BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
                
                'Reset our alpha premultiplication flag to match
                m_IsAlphaPremultiplied = srcDIB.getAlphaPremultiplication
                
                createFromExistingDIB = True
                Exit Function
            
            Else
            
                'Create a new, blank DIB the same size as the source DIB
                If createBlank(srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBColorDepth) Then
                    
                    'Copy the image data without modification
                    BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
                    
                    'Reset our alpha premultiplication flag to match
                    m_IsAlphaPremultiplied = srcDIB.getAlphaPremultiplication
                    
                    createFromExistingDIB = True
                    Exit Function
                    
                End If
                
            End If
        
        'If new width and height values are being specified, the transfer is a bit more complex
        Else
        
            'Create a new, blank DIB at the requested size
            If createBlank(newWidth, newHeight, srcDIB.getDIBColorDepth) Then

                'We now break up the resize request into two branches: 24bpp data, and 32bpp data.
                If srcDIB.getDIBColorDepth = 24 Then
                
                    'If either dimension of the new image will be smaller than the source, request halftoning
                    If useHalftoning Then
                        SetStretchBltMode m_dibDC, STRETCHBLT_HALFTONE
                    Else
                        SetStretchBltMode m_dibDC, STRETCHBLT_COLORONCOLOR
                    End If
    
                    'Resize and copy the image data
                    StretchBlt m_dibDC, 0, 0, newWidth, newHeight, srcDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, vbSrcCopy
                    
                '32bpp data requires special attention, because StretchBlt won't halftone alpha data.
                Else
                    
                    'Instead of halftoning, switch between interpolation and nearest-neighbor
                    If useHalftoning Then
                        GDIPlusResizeDIB Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, InterpolationModeHighQualityBicubic
                    Else
                        GDIPlusResizeDIB Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, InterpolationModeNearestNeighbor
                    End If
                
                End If
                
                'Reset alpha premultiplication to match
                m_IsAlphaPremultiplied = srcDIB.getAlphaPremultiplication
                
                createFromExistingDIB = True
                Exit Function
                        
            End If
        
        End If
            
        createFromExistingDIB = False
    
    End If
    
    createFromExistingDIB = False

End Function

'Convert this DIB to 24bpp mode
Public Function convertTo24bpp(Optional ByVal newBackColor As Long = vbWhite) As Boolean

    'Make sure this DIB is 32bpp. If it isn't, running this function pointless.
    If m_DibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (m_dibDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
    
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB Me
                        
            'Composite the temporary DIB against a white background, per convention.
            tmpDIB.CompositeBackgroundColor ExtractR(newBackColor), ExtractG(newBackColor), ExtractB(newBackColor)
            
            'Now erase our own DIB
            eraseDIB
            
            'Create a new DIB that's exactly the same size as the old one
            If createBlank(tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, 24) Then
                
                'Copy the image data from the temporary DIB without modification
                BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
                
                'Erase the temporary DIB
                tmpDIB.eraseDIB
                Set tmpDIB = Nothing
                
                'Exit
                convertTo24bpp = True
                Exit Function
                
            End If
    
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong
    convertTo24bpp = False

End Function

'Convert this DIB to 32bpp mode
Public Function convertTo32bpp(Optional ByVal newTransparency As Byte = 255) As Boolean

    'Make sure this DIB isn't empty
    If (m_dibDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then

        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        Dim x As Long, y As Long, QuickX As Long
        
        Dim newTransFloat As Double
        newTransFloat = newTransparency / 255

        If Me.getDIBColorDepth = 32 Then
            
            'If we are already 32bpp, with premultiplied alpha, unpremultiply the temporary DIB now
            If Me.getAlphaPremultiplication Then Me.SetAlphaPremultiplication False
            
            'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
            ' be completely transparent (and we don't want that!)
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
            
            'Loop through the image, setting the alpha of each pixel to 255 (opaque)
            For x = 0 To m_dibWidth - 1
                QuickX = x * 4
            For y = 0 To m_dibHeight - 1
                
                iData(QuickX + 3, y) = iData(QuickX + 3, y) * newTransFloat
                
            Next y
            Next x

            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            'Per PD-convention, re-apply premultiplication now
            Me.SetAlphaPremultiplication True
            
            'Exit
            convertTo32bpp = True
            Exit Function
            
        Else
        
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB Me
            
            'Now erase our own DIB
            eraseDIB
    
            'Create a new DIB that's exactly the same size as the old one
            If Me.createBlank(tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, 32) Then
                
                'Copy the image data from the temporary DIB without modification
                BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
                
                'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
                ' be completely transparent (and we don't want that!)
                prepInternalSafeArray tmpSA
                CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
                
                'Loop through the image, setting the alpha of each pixel to 255 (opaque)
                For x = 0 To m_dibWidth - 1
                    QuickX = x * 4
                For y = 0 To m_dibHeight - 1
                    
                    iData(QuickX + 3, y) = newTransparency
                    
                Next y
                Next x
    
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                
                'If the alpha value is not 255, or if we started with a 32bpp image, apply premultiplication now
                If (newTransparency <> 255) Then Me.SetAlphaPremultiplication True
                
                'Exit
                convertTo32bpp = True
                Exit Function
                
            End If
            
        End If
        
    End If
    
    
    'If we made it to this line, something went horribly wrong
    convertTo32bpp = False

End Function

'Give this DIB a picture from a standard VB picture object
Public Function CreateFromPicture(ByRef srcPicture As StdPicture, Optional forceWhiteBackground As Boolean = False) As Boolean

    'Make sure the picture we're passed isn't empty
    If (Not srcPicture Is Nothing) Then
    
        'Make sure the picture is actually a picture
        If GetObjectType(srcPicture) = OBJ_BITMAP Then
        
            'Select the picture's attributes into a bitmap object
            Dim tmpBitmap As BITMAP
            GetObject srcPicture.Handle, Len(tmpBitmap), tmpBitmap
            
            'Use that bitmap object to create a new, blank DIB of the same size
            If createBlank(tmpBitmap.Width, tmpBitmap.Height, tmpBitmap.BitsPerPixel) Then
            
                'Create a new DC
                Dim tmpDC As Long
                tmpDC = Drawing.GetMemoryDC()
                
                'If successful, select the object into that DC
                If tmpDC <> 0 Then
                
                    'Temporary holder for the object selection
                    Dim oldBitmap As Long
                    oldBitmap = SelectObject(tmpDC, srcPicture.Handle)
                    
                    'Use BitBlt to copy the pixel data to this DIB
                    BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDC, 0, 0, vbSrcCopy
                    
                    'Now that we have the pixel data, erase all temporary objects
                    SelectObject tmpDC, oldBitmap
                    Drawing.FreeMemoryDC tmpDC
                    
                    'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                    ' black background. We typically want the background to be white, so perform a composite if requested.
                    If forceWhiteBackground And (m_DibColorDepth = 32) Then CompositeBackgroundColor
                    
                    'Success!
                    CreateFromPicture = True
                    Exit Function
                    
                End If
            
                CreateFromPicture = False
            
            End If
            
            CreateFromPicture = False
        
        End If
        
        CreateFromPicture = False
    
    End If
    
    CreateFromPicture = False

End Function

'Create a blank DIB. If no colorDepth is specified, it will default to 24bpp (16 million colors, no alpha-channel).
' If the DIB is 32bpp, an optional alpha parameter can be set for the DIB.
Public Function createBlank(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal colorDepth As Long = 24, Optional ByVal BackColor As Long = vbWhite, Optional ByVal dibAlpha As Long = 0) As Boolean
    
    'Erase any existing DIB data
    eraseDIB
    
    'PhotoDemon only supports 24 and 32 BPP at present
    If (colorDepth <> 32) And (colorDepth <> 24) Then colorDepth = 24
        
    'The back color may or may not be a system color, so translate it just in case
    If (BackColor <> vbWhite) Then BackColor = TranslateColor(BackColor)
    
    'Force the DIB to have a size of at least 1x1
    If iWidth <= 0 Then iWidth = 1
    If iHeight <= 0 Then iHeight = 1
    
    'Remember the requested color depth, width, and height
    m_DibColorDepth = colorDepth
    m_dibWidth = iWidth
    m_dibHeight = iHeight
    
    'Prepare the required header
    With m_dibHeader
        .Size = Len(m_dibHeader)
        .Planes = 1
        .BitCount = colorDepth
        .Width = iWidth
        .Height = -iHeight
        'As always, this value needs to be a multiple of four; with 32bpp that's automatic, with 24bpp it is not
        If colorDepth = 32 Then
            m_dibArrayWidth = 4 * iWidth
        Else
            m_dibArrayWidth = (iWidth * 3 + 3) And &HFFFFFFFC
        End If
        .ImageSize = m_dibArrayWidth * iHeight
    End With
    
    'If we don't already have a compatible DC, create one now
    If m_dibDC = 0 Then
        m_dibDC = Drawing.GetMemoryDC()
        
    'If we already have a DC, empty it
    Else
        
        If m_dibHandle <> 0 Then
            SelectObject m_dibDC, m_dibHandleOriginal
            DeleteObject m_dibHandle
        End If
        
    End If
    
    If m_dibDC <> 0 Then
    
        'Activate color management for this DC.
        ' NOTE!  Due to rejiggering of PD's color management engine, this is no longer necessary.  Color management is only applied
        '         at image load-time (to convert into sRGB), and when drawing to the screen - but when drawing to the screen, all
        '         management is handled at the destination.  For performance reasons, we no longer color-manage DIBs themselves,
        '         which allows us to move RGB/A data between them much more efficiently.
        'turnOnColorManagementForDC m_dibDC
        
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
                
                Dim dibRect As RECT
                SetRect dibRect, 0, 0, iWidth, iHeight
            
                Dim hBrush As Long
                hBrush = CreateSolidBrush(BackColor)
                FillRect m_dibDC, dibRect, hBrush
                DeleteObject hBrush
                
                'Also, reset alpha premultiplication to false
                m_IsAlphaPremultiplied = False
                
            Else
            
                'Finally, set the backColor and/or transparency
                GDIPlusFillDIBRect Me, 0, 0, iWidth, iHeight, BackColor, dibAlpha
                
                'Because alpha premultiplication is unknown, we'll assume it is FALSE
                m_IsAlphaPremultiplied = False
                
            End If
            
        'If DIB creation failed, clear out the work we've done so far
        Else
            eraseDIB
        End If
        
    End If
    
    'Return success contingent on whether we have a DIB pointer or not
    createBlank = CBool(m_dibHandle <> 0)
    
End Function

'This will effectively reset everything related to this DIB, including all image data. Use cautiously!
Public Sub eraseDIB(Optional ByVal alsoReleaseDC As Boolean = False)

    'If we have image data, clear it out
    If m_dibHandle <> 0 Then
        
        If m_dibDC <> 0 Then SelectObject m_dibDC, m_dibHandleOriginal
        DeleteObject m_dibHandle
        
        'Inside debug mode, decrement a global DIB counter
        #If DEBUGMODE = 1 Then
            g_DIBsDestroyed = g_DIBsDestroyed + 1
        #End If
    
    End If
    
    If alsoReleaseDC And (m_dibDC <> 0) Then
        Drawing.FreeMemoryDC m_dibDC
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
Public Sub resetDIB(Optional ByVal fillValue As Byte = 0)

    'Retrieve a pointer to the DIB array and the array's current size
    Dim tmpDIBPointer As Long, tmpDibSize As Long
    retrieveDIBPointerAndSize tmpDIBPointer, tmpDibSize
    
    'Erase everything
    If (tmpDIBPointer <> 0) And (tmpDibSize > 0) Then FillMemory tmpDIBPointer, tmpDibSize, fillValue
    
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
    
    eraseDIB True
    
    'Release the ICC profile manager
    Set ICCProfile = Nothing
    
End Sub

'Draw this DIB to a picture box. The image will be automatically centered and sized to fit.
Public Sub RenderToPictureBox(ByRef dstPicture As PictureBox, Optional ByVal colorManagementMatters As Boolean = True, Optional ByVal doNotStretchIfSmaller As Boolean = False)

    'Erase any existing picture
    If dstPicture.AutoRedraw Then dstPicture.Picture = LoadPicture("")
    
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
    If srcHeight > 0 Then srcAspect = srcWidth / srcHeight Else srcAspect = 1
    If dstHeight > 0 Then dstAspect = dstWidth / dstHeight Else dstAspect = 1
        
    Dim dWidth As Long, dHeight As Long, previewX As Long, previewY As Long
    convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, dWidth, dHeight
    
    If fitPrevented Then
        previewX = (dstPicture.ScaleWidth - srcWidth) / 2
        previewY = (dstPicture.ScaleHeight - srcHeight) / 2
    Else
        If srcAspect > dstAspect Then
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
        
        'Create a temporary DIB to house the stretched image.  This is required for proper GDI-based color management.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createBlank dWidth, dHeight, 24
        
        'For 24bpp images, simply copy over the current image data
        If m_DibColorDepth = 24 Then
        
            'If the target area is smaller than the DIB, request halftoning
            If (dstPicture.ScaleWidth < m_dibWidth) Or (dstPicture.ScaleHeight < m_dibHeight) Then
                SetStretchBltMode tmpDIB.getDIBDC, STRETCHBLT_HALFTONE
            Else
                SetStretchBltMode tmpDIB.getDIBDC, STRETCHBLT_COLORONCOLOR
            End If
            
            StretchBlt tmpDIB.getDIBDC, 0, 0, dWidth, dHeight, m_dibDC, 0, 0, m_dibWidth, m_dibHeight, vbSrcCopy
        
        'For 32bpp images, we must first apply a background checkerboard pattern, then alpha-blend the resized
        ' image onto it.
        Else
            GDI_Plus.GDIPlusFillDIBRect_Pattern tmpDIB, 0, 0, dWidth, dHeight, g_CheckerboardPattern
            GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, dWidth, dHeight, Me, 0, 0, m_dibWidth, m_dibHeight
        End If
        
        'Activate color management for the destination picture box
        AssignDefaultColorProfileToObject dstPicture.hWnd, dstPicture.hDC
        TurnOnColorManagementForDC dstPicture.hDC
        
        'Render the image to the picture box
        BitBlt dstPicture.hDC, previewX, previewY, dWidth, dHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
    Else
        
        'For 24bpp images, simply copy over the current image data
        If m_DibColorDepth = 24 Then
        
            'If the target area is smaller than the DIB, request halftoning
            If (dstPicture.ScaleWidth < m_dibWidth) Or (dstPicture.ScaleHeight < m_dibHeight) Then
                SetStretchBltMode dstPicture.hDC, STRETCHBLT_HALFTONE
            Else
                SetStretchBltMode dstPicture.hDC, STRETCHBLT_COLORONCOLOR
            End If
            
            StretchBlt dstPicture.hDC, previewX, previewY, dWidth, dHeight, m_dibDC, 0, 0, m_dibWidth, m_dibHeight, vbSrcCopy
        
        'For 32bpp images, we must first apply a background checkerboard pattern, then alpha-blend the resized
        ' image onto it.
        Else
            GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, previewX, previewY, dWidth, dHeight, g_CheckerboardPattern, dstPicture.hDC
            GDI_Plus.GDIPlus_StretchBlt Nothing, previewX, previewY, dWidth, dHeight, Me, 0, 0, m_dibWidth, m_dibHeight, , , dstPicture.hDC
        End If
        
    End If
        
    If dstPicture.AutoRedraw Then
        dstPicture.Picture = dstPicture.Image
        dstPicture.Refresh
    End If
    
End Sub

'Load a DIB's header and image information from file.  If the file is compressed, it will automatically be decompressed.
' Optionally, if the DIB data is embedded inside another file, that handle can be passed to the this function, and it will attempt to
' load DIB data from the current file pointer location.  (In this case, just supply a blank filename.)
Public Function CreateFromFile(ByRef srcFilename As String, Optional ByVal useEmbeddedLocation As Boolean = False, Optional ByVal hFile As Long = 0) As Boolean
    
    'pdFSO is used so we can support Unicode paths
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If this is not an embedded file, create and open it now
    If Not useEmbeddedLocation Then
        
        If Not cFile.CreateFileHandle(srcFilename, hFile, True, False, OptimizeSequentialAccess) Then
            
            #If DEBUGMODE = 1 Then
                If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.createFromFile failed to generate a handle for " & srcFilename & ".  Load abandoned."
            #End If
            
            CreateFromFile = False
            Exit Function
            
        End If
        
    End If
    
    'Check to make sure this file actually contains DIB data
    Dim DIBIDCheck As Long
    cFile.ReadDataFromFile hFile, VarPtr(DIBIDCheck), 4&
        
    If (DIBIDCheck <> DIB_IDENTIFIER) Then
        
        Message "Failed to load embedded image; invalid header found."
        CreateFromFile = False
        
        If Not useEmbeddedLocation Then cFile.CloseFileHandle hFile
        Exit Function
        
    End If
    
    'In the future, we may need to branch due to version differences, but for now, we can make any adjustments in-line
    Dim DIBVersionCheck As Long
    cFile.ReadDataFromFile hFile, VarPtr(DIBVersionCheck), 4&
    
    'Retrieve color depth
    Dim fColorDepth As Long
    cFile.ReadDataFromFile hFile, VarPtr(fColorDepth), 4&
    
    'The newest, 2015 version of PD DIBs retrieves alpha premultiplication as well
    Dim localAlphaPremultiplied As Boolean
    If DIBVersionCheck = DIB_FILE_VERSION_2015 Then
    
        cFile.ReadDataFromFile hFile, VarPtr(localAlphaPremultiplied), 2&
        
    'If alpha premultiplication state is not stored, make a best guess based on color depth
    Else
        
        If fColorDepth = 24 Then
            localAlphaPremultiplied = False
        Else
            localAlphaPremultiplied = True
        End If
        
    End If
    
    'Retrieve DIB dimensions
    Dim fWidth As Long, fHeight As Long, fArrayWidth As Long
    cFile.ReadDataFromFile hFile, VarPtr(fWidth), 4&
    cFile.ReadDataFromFile hFile, VarPtr(fHeight), 4&
    cFile.ReadDataFromFile hFile, VarPtr(fArrayWidth), 4&
    
    'Is the DIB data compressed?
    Dim toDecompress As Boolean
    cFile.ReadDataFromFile hFile, VarPtr(toDecompress), 2&
    
    Dim fileData() As Byte
    
    'If the data *is* compressed, we need to retrieve two additional pieces of information: the original and compressed size
    ' of the entire pixel chunk.
    Dim origDataSize As Long, compressedDataSize As Long
    
    'Also, as a failsafe, make sure the DIB has dimensions prior to loading pixel data.
    If (fWidth > 0) And (fHeight > 0) Then
    
        If toDecompress Then
        
            'If zLib is not enabled for some reason, abandon the load
            If Not g_ZLibEnabled Then
                Message "DIB data must be decompressed, but zLib is disabled or missing.  Please install the zLib plugin to continue."
                CreateFromFile = False
                Exit Function
            End If
            
            'Retrieve chunk sizes
            cFile.ReadDataFromFile hFile, VarPtr(origDataSize), 4&
            cFile.ReadDataFromFile hFile, VarPtr(compressedDataSize), 4&
            
            'In one fell swoop, load the entire pixel chunk into a temporary array
            ReDim fileData(0 To compressedDataSize - 1) As Byte
            cFile.ReadDataFromFile hFile, VarPtr(fileData(0)), compressedDataSize
            
        Else
            
            'If the pixel data is uncompressed, we can load it directly into the DIB itself, avoiding the need for
            ' a temporary copy.  Do nothing at this stage.
                        
        End If
    
    End If
    
    'Initialize a blank DIB at the requested size.  Besides giving us a memory target, this will also initialize all module-level
    ' parameters to match the file's settings.
    If createBlank(fWidth, fHeight, fColorDepth) Then
        
        'Now it's time to overwrite the DIB memory region with the pixel data in the file.
        Dim tmpDIBPointer As Long, tmpDIBLength As Long
        Me.retrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
        
        'If the DIB data is compressed, decompress it now.  Note that we decompress directly into the DIB itself, to avoid the
        ' need for a costly CopyMemory call.
        If toDecompress Then
            
            'If the data is decompressed succesfully, convert it to a 2D array and copy it into this class's DIB memory chunk
            If Plugin_zLib_Interface.decompressNakedPointers(VarPtr(fileData(0)), compressedDataSize, tmpDIBPointer, tmpDIBLength) Then
            
                'Immediately release our copy of the compressed file data
                Erase fileData
            
            'If decompression failed, abandon further processing
            Else
                
                Message "Decompression was not successful (unspecified zLib plugin error).  DIB processing abandoned."
                CreateFromFile = False
                Exit Function
                
            End If
        
        'If the DIB data is not compressed, copy it straight into the DIB
        Else
            
            'In one fell swoop, load the entire pixel chunk into our already-created DIB
            cFile.ReadDataFromFile hFile, tmpDIBPointer, tmpDIBLength
            
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
    If Not useEmbeddedLocation Then cFile.CloseFileHandle hFile
    
End Function

'Write this class's DIB information to file.  Compression can optionally be applied to the DIB data.
' Also, if the calling function wants to embed the DIB's data inside an existing file, specify the embedInFile parameter and pass a valid
' file handle, with the pointer already moved to the location desired for embedding.
Public Function WriteToFile(ByRef dstFilename As String, Optional ByVal toCompress As Boolean = True, Optional ByVal embedInFile As Boolean = False, Optional ByRef hFile As Long) As Boolean
    
    'Before proceeding, make sure this DIB actually exists
    If (Me.getDIBWidth <= 0) Or (Me.getDIBHeight <= 0) Then
        
        #If DEBUGMODE = 1 Then
            If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.writeToFile was called on an empty DIB.  Fix this!"
        #End If
        
        WriteToFile = False
        Exit Function
        
    End If
    
    'File interactions are handled via pdFSO, so we can easily support Unicode paths
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If we are not embedding the data inside an existing file, create a new file now.
    If Not embedInFile Then
        
        'Delete any existing file (overwrite)
        If cFile.FileExist(dstFilename) Then cFile.KillFile dstFilename
        
        'Open a new file
        If Not cFile.CreateFileHandle(dstFilename, hFile, True, True, OptimizeSequentialAccess) Then
        
            #If DEBUGMODE = 1 Then
                If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.writeToFile failed to create a file handle for " & dstFilename & ".  Write abandoned."
            #End If
            
            WriteToFile = False
            Exit Function
        
        End If
        
    End If
    
    'Write identifiers first
    Dim tmpLong As Long
    
    tmpLong = DIB_IDENTIFIER
    cFile.WriteDataToFile hFile, VarPtr(tmpLong), 4&
    
    tmpLong = DIB_FILE_VERSION_2015
    cFile.WriteDataToFile hFile, VarPtr(tmpLong), 4&
    
    'Color depth
    cFile.WriteDataToFile hFile, VarPtr(Me.getDIBColorDepth), 4&
    
    'Alpha premultiplication (yes, this is stored even for 24-bit images)
    cFile.WriteDataToFile hFile, VarPtr(m_IsAlphaPremultiplied), 2&
    
    'Size
    cFile.WriteDataToFile hFile, VarPtr(Me.getDIBWidth), 4&
    cFile.WriteDataToFile hFile, VarPtr(Me.getDIBHeight), 4&
    cFile.WriteDataToFile hFile, VarPtr(Me.getDIBArrayWidth), 4&
        
    'And finally, the pixel data.  Start by retrieving two key items of business: a pointer to this DIB's raw pixel data,
    ' and the total length of that pixel data.
    Dim tmpDIBPointer As Long, tmpDIBLength As Long
    Me.retrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
        
    'If compression has been requested, and zLib is available, we'll shrink the DIB data down before writing it.
    If toCompress Then
        
        If g_ZLibEnabled Then
            
            'Compress that 1D array
            Dim compressedImageData() As Byte, compressedSize As Long
            
            'If compression was successful, write out the compressed data to file
            If Plugin_zLib_Interface.compressNakedPointerToArray(tmpDIBPointer, tmpDIBLength, compressedImageData, compressedSize) Then
            
                'Mark the data as compressed
                cFile.WriteDataToFile hFile, VarPtr(toCompress), 2&
                
                'Write the original size of the array, followed by the compressed size, followed by the image data itself.  (The array sizes will
                ' help us automatically create intelligently-sized buffers at decompression time.)
                cFile.WriteDataToFile hFile, VarPtr(tmpDIBLength), 4&
                cFile.WriteDataToFile hFile, VarPtr(compressedSize), 4&
                cFile.WriteDataToFile hFile, VarPtr(compressedImageData(0)), compressedSize
                            
            'If compression failed, write uncompressed data instead.
            Else
                
                Message "Compression failed (unspecified zLib plugin error).  Uncompressed data written instead."
                toCompress = False
                cFile.WriteDataToFile hFile, VarPtr(toCompress), 2&
                cFile.WriteDataToFile hFile, tmpDIBPointer, tmpDIBLength
                
            End If
            
        Else
            
            #If DEBUGMODE = 1 Then
                If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.writeToFile requested zLib compression, but zLib is missing or disabled.  Data will not be compressed!"
            #End If
            
            toCompress = False
            cFile.WriteDataToFile hFile, VarPtr(toCompress), 2&
            cFile.WriteDataToFile hFile, tmpDIBPointer, tmpDIBLength
            
        End If
        
        'Immediately free our temporary array
        Erase compressedImageData
        
    'If compression has not been requested, or zLib is unavailable, we can dump the raw image data to file as-is.
    Else
        
        cFile.WriteDataToFile hFile, VarPtr(toCompress), 2&
        cFile.WriteDataToFile hFile, tmpDIBPointer, tmpDIBLength
        
    End If
        
    'If this is not an embedding, close the file now
    If Not embedInFile Then cFile.CloseFileHandle hFile

End Function

'Write this class's DIB information to a valid BMP file.
Public Sub WriteToBitmapFile(ByRef dstFilename As String)

    'First, make sure this DIB actually contains image data
    If Not Me.hasImage() Then Exit Sub
    
    'Calculate the size of a scanline
    Dim slWidth As Long
    
    If Me.getDIBColorDepth = 32 Then
        slWidth = Me.getDIBArrayWidth
    Else
        slWidth = ((m_dibWidth * m_DibColorDepth + 31) \ 32) * 4
    End If
    
    'A valid bitmap header consists of two parts: a file header, and an image header.
    Dim bmpFileHeader As BITMAPFILEHEADER
    Dim bmpInfoHeader As BITMAPINFOHEADER
    
    'Build the file header first
    With bmpFileHeader
        
        'BMP identifier
        .Type = &H4D42
        
        'Length of the file
        .Size = LenB(bmpFileHeader) + LenB(bmpInfoHeader) + (slWidth * Me.getDIBHeight)
        
        'Length of the header area (e.g. offset to the actual pixel data)
        .OffBits = LenB(bmpFileHeader)
        
    End With
    
    '...then the image header
    With bmpInfoHeader
        .Size = 40
        .Planes = 1
        .BitCount = Me.getDIBColorDepth()
        .Width = Me.getDIBWidth()
        .Height = Me.getDIBHeight()
    End With
    
    'Finally, the image bytes.  We call GetDIBits here because DIBs are upside-down, but we need to store them rightside-up in the BMP.
    ' We could do this manually on a per-scanline basis, but this is more succinct.
    Dim iData() As Byte
    ReDim iData(0 To slWidth - 1, 0 To m_dibHeight - 1) As Byte
    GetDIBits m_dibDC, m_dibHandle, 0, m_dibHeight, iData(0, 0), bmpInfoHeader, 0&
    
    'With all our information in place, check to see if that file already exists. Delete it if necessary.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(dstFilename) Then cFile.KillFile dstFilename
    
    'Writing the file is easy; just dump the two headers and image data into it!
    
    Dim hFile As Long
    If cFile.CreateFileHandle(dstFilename, hFile, False, True, OptimizeSequentialAccess) Then
    
        With cFile
            
            'To avoid automatic struct alignment, we'll write out the header manually.
            .WriteDataToFile hFile, VarPtr(bmpFileHeader.Type), 2&
            .WriteDataToFile hFile, VarPtr(bmpFileHeader.Size), 4&

            Dim reservedBytes As Long
            .WriteDataToFile hFile, VarPtr(reservedBytes), 4&
            .WriteDataToFile hFile, VarPtr(bmpFileHeader.OffBits), 4&
            
            'With the header finished, we add the lightweight BMP header and the actual pixel data
            .WriteDataToFile hFile, VarPtr(bmpInfoHeader), LenB(bmpInfoHeader)
            .WriteDataToFile hFile, VarPtr(iData(0, 0)), slWidth * m_dibHeight
            
        End With
        
        'That's it!
        cFile.CloseFileHandle hFile

        #If DEBUGMODE = 1 Then
            If Not pdDebug Is Nothing Then pdDebug.LogAction "pdDIB.writeToBitmapFile successfully wrote " & dstFilename & "."
        #End If

    Else
        #If DEBUGMODE = 1 Then
            If Not pdDebug Is Nothing Then pdDebug.LogAction "WARNING!  pdDIB.writeToBitmapFile failed to create a valid file handle for " & dstFilename & ".  Write abandoned."
        #End If
    End If
        
End Sub

'As a helper to external functions, this sub can be used to fill a destination array with a copy the image's current data.
' Note that this is raw image bytes - no header whatsoever - so it's assumed the caller knows how to deal with the data.
Public Sub copyImageBytesIntoStream(ByRef dstArray() As Byte)

    'Retrieve a pointer to the current DIB bits (in VB format).  Technically we could just use the m_dibBits pointer
    ' and calculate bounds manually, but this provides a foolproof way to do the same thing using existing code.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
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
Public Sub retrieveDIBPointerAndSize(ByRef dibPointer As Long, ByRef dibSize As Long)
    dibPointer = m_dibBits
    dibSize = m_dibArrayWidth * m_dibHeight
End Sub

'The counterpart to copyImageBytesIntoStream, above.  This function will blindly copy an array of bytes over the current
' image's data.  For it to work, the DIB must have been correctly created at the right size and color depth BEFORE CALLING
' THIS FUNCTION.  Otherwise, you will get errors.
Public Sub copyStreamOverImageArray(ByRef srcArray() As Byte)
    CopyMemory ByVal m_dibBits, srcArray(0), UBound(srcArray) + 1
End Sub

'Sometimes this class needs to access its own DIB bits. Here's how.
Private Sub prepInternalSafeArray(ByRef dstSafeArray As SAFEARRAY2D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_dibArrayWidth
        .pvData = m_dibBits
    End With
End Sub

'2D array access is slow, so PD tries to use 1D arrays when possible.  This function initializes a base SafeArray, but the caller
' will need to manually update it to point at a desired line.
Private Sub prepInternalSafeArray1D(ByRef dstSafeArray As SAFEARRAY1D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = m_dibArrayWidth
        .pvData = m_dibBits
    End With
End Sub

'Sometimes this class needs to access the DIB bits of other DIBs. Here's how.
Private Sub PrepExternalSafeArray(ByRef srcDIB As pdDIB, ByRef dstSafeArray As SAFEARRAY2D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.getDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.getDIBArrayWidth
        .pvData = srcDIB.getActualDIBBits
    End With
End Sub

'Pre-composite an image with an alpha-channel against a background color. Until PhotoDemon is capable of rendering transparent
' images itself, this is necessary to give transparent images a white background.
Public Sub CompositeBackgroundColor(Optional ByVal newR As Byte = 255, Optional ByVal newG As Byte = 255, Optional ByVal newB As Byte = 255)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If m_DibColorDepth <> 32 Then Exit Sub
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Compositing DIB of size " & m_dibWidth & "x" & m_dibHeight & ".  Alpha premultiplication = " & m_IsAlphaPremultiplied & "."
    #End If
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.getDIBArrayWidth
    dibPointer = Me.getActualDIBBits
    
    Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
    prepInternalSafeArray1D dibSA
    
    Dim x As Long, y As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'Because our alpha values are pre-multiplied, we can composite them against the background color via use of a look-up table.
    Dim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    'Populate a unique lookup table for each color, based on each possible alpha value (0 to 255)
    For x = 0 To 255
        tmpAlpha = 1 - (x / 255)
        rLookup(x) = newR * tmpAlpha
        gLookUp(x) = newG * tmpAlpha
        bLookup(x) = newB * tmpAlpha
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
                dibPixels(x + 2) = newR
                dibPixels(x + 1) = newG
                dibPixels(x) = newB
            Else
            
                'Use that alpha value to blend the current colors with the newly requested color
                If m_IsAlphaPremultiplied Then
                    dibPixels(x) = dibPixels(x) + bLookup(checkAlpha)
                    dibPixels(x + 1) = dibPixels(x + 1) + gLookUp(checkAlpha)
                    dibPixels(x + 2) = dibPixels(x + 2) + rLookup(checkAlpha)
                    dibPixels(x + 3) = 255
                Else
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = checkAlpha / 255
                
                    'Use that alpha value to blend the current colors with the newly requested color
                    dibPixels(x) = Blend2Colors(dibPixels(x), newB, tmpAlpha)
                    dibPixels(x + 1) = Blend2Colors(dibPixels(x + 1), newG, tmpAlpha)
                    dibPixels(x + 2) = Blend2Colors(dibPixels(x + 2), newR, tmpAlpha)
                    dibPixels(x + 3) = 255
                    
                End If
                
            End If
            
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

'Copy the alpha values from another pdDIB object to this one. This is useful when the alpha channel for this DIB must be lost
' during a transformation (typically something involving FreeImage or GDI+), and a temporary DIB was made to preserve the alpha
' data. If the image is currently 24bpp, this function will first convert it to 32bpp before copying the alpha data.
'
'IMPORTANT NOTE: this function does not deal with premultiplied alpha whatsoever.  It simply copies alpha data as-is.
' You must manually set the alpha premultiplication property yourself, and you must also make sure that any number of
' un/premultiply steps are applied to both DIBs to keep the alpha values properly in sync.

'IMPORTANT NOTE: to keep this function fast, NO ARRAY BOUNDS CHECKING IS DONE. Make sure that the passed pdDIB object is
' THE SAME SIZE (or larger) than this DIB, or you will experience critical errors.
Public Sub copyAlphaFromExistingDIB(ByRef srcDIB As pdDIB)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If m_DibColorDepth <> 32 Then convertTo32bpp

    'Prepare a SafeArray that points to our own DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Now make a second array that points to the source DIB data
    Dim aData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepExternalSafeArray srcDIB, srcSA
    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    'Loop through the image, copying alpha values as we go
    For x = 0 To m_dibWidth - 1
        QuickX = x * 4
    For y = 0 To m_dibHeight - 1
        
        'Access the alpha data for this pixel
        iData(QuickX + 3, y) = aData(QuickX + 3, y)
        
    Next y
    Next x
    
    'With our alpha transfer complete, point both arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    CopyMemory ByVal VarPtrArray(aData), 0&, 4
    Erase aData

End Sub

'DIBs created from GDI+ contain pre-multiplied alpha values. These are great for high-performance blting, but terrible for
' photo editing. This routine can be used to either apply or remove premultiplied alpha from an image.
Public Sub SetAlphaPremultiplication(Optional ByVal applyPremultiplication As Boolean = False, Optional ByVal ignoreEmbeddedValue As Boolean = False)

    'This function doesn't matter if the image isn't 32bpp
    If m_DibColorDepth <> 32 Then Exit Sub
    
    'If alpha premultiplication already matches the requested state, exit now
    If (m_IsAlphaPremultiplied = applyPremultiplication) And (Not ignoreEmbeddedValue) Then
        Debug.Print "WARNING! Alpha premultiplication = " & applyPremultiplication & " is pointless, as image is already in that state.  (" & m_dibWidth & "x" & m_dibHeight & "x" & m_DibColorDepth & ")  Abandoning request."
        Exit Sub
    End If
    
    'Prepare a SAFEARRAY, which allows us to access the DIB data directly
    'Dim iData() As Byte
    'Dim tmpSA As SAFEARRAY2D
    'prepInternalSafeArray tmpSA
    'CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
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
    scanlineSize = Me.getDIBArrayWidth
    dibPointer = Me.getActualDIBBits
    
    Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
    prepInternalSafeArray1D dibSA
    
    Dim x As Long, y As Long, QuickX As Long
    Dim r As Long, g As Long, b As Long
    Dim tmpAlpha As Byte, tmpAlphaModifier As Double
    
    'Loop through the image, converting alpha as we go
    For y = 0 To finalY
        
        'Point our 1D pixel array at the proper scanline
        dibSA.pvData = dibPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(dibPixels()), VarPtr(dibSA), 4
        
    For x = 0 To finalX Step 4
        
        'Retrieve alpha for the current pixel
        tmpAlpha = dibPixels(x + 3)
        
        'Branch according to applying or removing premultiplication
        If applyPremultiplication Then
        
            'When applying premultiplication, we can ignore fully opaque pixels
            If tmpAlpha <> 255 Then
            
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
                    If r > 255 Then r = 255
                    If g > 255 Then g = 255
                    If b > 255 Then b = 255
                                    
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
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub alphaBlendToDC(ByVal dstDC As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    Dim bfParams As Long
    
    If m_DibColorDepth = 32 Then
        
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Also, raise a warning if premultiplication is not set
        If Not m_IsAlphaPremultiplied Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDC being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
        
    AlphaBlend dstDC, dstX, dstY, IIf(newWidth = 0, m_dibWidth, newWidth), IIf(newHeight = 0, m_dibHeight, newHeight), m_dibDC, 0, 0, m_dibWidth, m_dibHeight, bfParams
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub alphaBlendToDCEx(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customAlpha As Long = 255)
    
    Dim bfParams As Long
    
    If m_DibColorDepth = 32 Then
    
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Raise a warning if premultiplication is not set
        If Not m_IsAlphaPremultiplied Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDCEx being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
    
    AlphaBlend dstDC, dstX, dstY, dstWidth, dstHeight, m_dibDC, srcX, srcY, srcWidth, srcHeight, bfParams
    
End Sub
