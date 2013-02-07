VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Layer class
'Copyright ©2011-2013 by Tanner Helland
'Created: 29/August/12
'Last updated: 25/November/12
'Last update: added a function for copying just the alpha values from another layer object
'
'This marvelous class is PhotoDemon's replacement for picture boxes (which it used to store images in previous versions).
' pdLayer is a powerful DIB class responsible for managing all image data in memory.  Some of the benefits provided by
' this approach include:
'
'- Ability to load large images without problems
'- Alpha channels
'- High bit depths
'- Much faster than picture boxes
'- Allows for the eventual implementation of layers and adjustment layers
'
'These are serious benefits.
'
'Note that anything you can do with an hDC property you can do with this class - simply use the getLayerDC function to
' return the DIB's hDC, then do with it what you please.  All functions are heavily commented and should be self-explanatory.
'
'Note also that this class is treated as a subset of pdImage().  Right now each pdImage object only contains one layer
' (the image itself), but in the future a pdImage object could theoretically store many layers.
'
'SPECIAL THANKS
'In building this class, I utilized a number of other DIB classes for reference and testing.  Special thanks to:
' Herman Liu (personal correspondence)
' Carles PV's iBMP project: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Steve McMahon's DIB/SafeArray analysis: http://www.vbaccelerator.com/home/VB/Code/vbMedia/DIB_Sections/True_Colour_DIBSection/article.asp
' Many thanks to these three individuals for their outstanding work on graphics in VB.
'
'***************************************************************************

Option Explicit


'DIB Types
Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type Bitmap
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
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    Colorused As Long
    ColorImportant As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    Type      As Integer
    Size      As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits   As Long
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
    lBound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type


'Drawing API functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCHBLT_COLORONCOLOR As Long = 3
Private Const STRETCHBLT_HALFTONE As Long = 4

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DIB API functions
' (Note that these are currently declared in FastDrawing as well)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'DC API functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

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

'Clipboard interaction
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Const CLIPBOARD_FORMAT_BMP As Long = 2
Private Const CLIPBOARD_FORMAT_DIB As Long = 8

'Variables related to the DIB
'hDC for this layer
Private layerDC As Long
'DIB handle for this layer
Private layerDIB As Long
'Original handle when this layer is first created (we must store this so we can properly clean up the DIB when we're finished)
Private layerDIBOriginal As Long
'Pointer to the actual DIB bits
Private layerDIBits As Long
'Persistent DIB header; this will be used to create the DIB associated with this layer
Private layerDIBHeader As BITMAPINFOHEADER

'The layer's width and height
Private layerWidth As Long, layerHeight As Long

'The layer's array width (layerWidth * 4 for 32-bit, varies for 24-bit due to DWORD-alignment)
Private layerArrayWidth As Long

'The layer's color depth (should only ever be 24 or 32)
Private layerColorDepth As Long

'Used when writing/reading the layer data to/from a file
Private Const LAYER_IDENTIFIER As String * 4 = "PDlr"
Private Const LAYER_FILE_VERSION_2012 As Long = &H1000

'These variables are a temporary addition to ensure that print previewing still works.  They will be removed in a future
' update as part of a print preview revamp, so DO NOT add features that rely on them.
Public PreviewX As Long
Public PreviewY As Long
Public PreviewWidth As Long
Public PreviewHeight As Long

'When converting a 32bpp layer to 8bpp with transparency, we need to know the location of a transparent pixel.
' This location is passed to FreeImage, and it designates the palette index at that location as the transparent index.
Private trnsX As Long, trnsY As Long

'Remember the original color of the pixels made transparent, then restore it after the "magic magenta" has served its purpose
Private originalTransparentColor As Long

'Return the original transparent color (set as part of the MarkAlpha routine).
Public Function getOriginalTransparentColor() As Long
    getOriginalTransparentColor = originalTransparentColor
End Function

'Return the location of a transparent pixel.  This will be set as part of the MarkAlpha routine.
Public Sub getTransparentLocation(ByRef srcX As Long, ByRef srcY As Long)
    srcX = trnsX
    srcY = trnsY
End Sub

'This function will apply an "alpha-cutoff" to a 32bpp image.  Any alpha values above the cutoff will be replaced with
' magic magenta (253,0,253) and the location of a transparent pixel (trnsX, trnsY) will be marked.
Public Sub applyAlphaCutoff(Optional ByVal cutOff As Long = 127, Optional ByVal forceColor As Boolean = True)

    'Make sure this layer is 32bpp.  If it isn't, running this function is pointless.
    If layerColorDepth = 32 Then

        'Make sure this layer isn't empty
        If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
            'Loop through the image and force each alpha value to 0 and 255 based on the supplied cutoff.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
            
            Dim chkAlpha As Byte
            Dim tmpAlpha As Double
            
            trnsX = -1
                
            'Loop through the image, checking alphas as we go
            For x = 0 To layerWidth - 1
                QuickX = x * 4
            For y = 0 To layerHeight - 1
                
                chkAlpha = iData(QuickX + 3, y)
                
                'If the alpha value is less than the cutoff, mark this pixel
                If chkAlpha < cutOff Then
                
                    'Remember this location if we haven't already
                    If trnsX = -1 Then
                        trnsX = x
                        trnsY = y
                        originalTransparentColor = RGB(iData(QuickX + 2, y), iData(QuickX + 1, y), iData(QuickX, y))
                    End If
                    
                    'When previewing this effect, we don't actually want to change pixel colors.  Otherwise,
                    ' HALFTONE-mode resampling may bleed the magenta into the preview and give poor results.
                    If forceColor Then
                        'Color this pixel with magic magenta (253,0,253)
                        iData(QuickX + 2, y) = 253
                        iData(QuickX + 1, y) = 0
                        iData(QuickX, y) = 253
                    End If
                    
                    'Blank the alpha channel
                    iData(QuickX + 3, y) = 0
                    
                'If the pixel is not beneath the cut-off, and not full opaque, composite it against white
                ElseIf chkAlpha <> 255 Then
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = CSng(chkAlpha) / 255
            
                    'Use that alpha value to blend the current colors with the newly requested color, in a checkerboard pattern
                    iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), 255, tmpAlpha)
                    iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), 255, tmpAlpha)
                    iData(QuickX, y) = Blend2Colors(iData(QuickX, y), 255, tmpAlpha)
                    
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

'Translate an OLE color to an RGB Long
Private Function TranslateColor(ByVal colorRef As Long) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, TranslateColor) Then
        TranslateColor = RGB(255, 255, 255)
    End If
End Function

'Return this layer's color depth
Public Function getLayerColorDepth() As Long
    getLayerColorDepth = layerColorDepth
End Function

'Return this layer's array width
Public Function getLayerArrayWidth() As Long
    getLayerArrayWidth = layerArrayWidth
End Function

'Return this layer's width
Public Function getLayerWidth() As Long
    getLayerWidth = layerWidth
End Function

'Return this layer's height
Public Function getLayerHeight() As Long
    getLayerHeight = layerHeight
End Function

'Return whether or not this layer has image data associated with it
Public Function hasImage() As Boolean
    hasImage = (layerDIB <> 0)
End Function

'Return this layer's hDC
Public Function getLayerDC() As Long
    getLayerDC = layerDC
End Function

'Return a pointer to this layer's DIB
Public Function getLayerDIB() As Long
    getLayerDIB = layerDIB
End Function

'Return a pointer to this layer's pixel data
Public Function getLayerDIBits() As Long
    getLayerDIBits = layerDIBits
End Function

'Copy the current layer's contents to the clipboard
Public Sub copyLayerToClipboard()
    
    'Make sure the current layer actually contains an image before copying it to the clipboard
    If (layerDIB <> 0) Then
    
        'Make sure the clipboard is available and working
        If (OpenClipboard(0) <> 0) Then
               
            'Get a handle to the current desktop, and create a compatible clipboard device context in it
            Dim desktophWnd As Long
            desktophWnd = GetDesktopWindow
            
            Dim desktopDC As Long, clipboardDC As Long
            desktopDC = GetDC(desktophWnd)
            clipboardDC = CreateCompatibleDC(desktopDC)
            
            'Make sure the clipboard was able to receive a compatible device context
            If (clipboardDC <> 0) Then
            
                'Create a bitmap compatible with the current desktop.  This will receive the actual pixel data of the current layer
                Dim clipboardBMP As Long
                clipboardBMP = CreateCompatibleBitmap(desktopDC, layerWidth, layerHeight)
                
                If (clipboardBMP <> 0) Then
                    
                    'Place the compatible bitmap within the clipboard device context
                    Dim clipboardOldBMP As Long
                    clipboardOldBMP = SelectObject(clipboardDC, clipboardBMP)
                    
                    'Use BitBlt to paint the current layer to that bitmap
                    BitBlt clipboardDC, 0, 0, layerWidth, layerHeight, layerDC, 0, 0, vbSrcCopy
                    
                    'Remove that bitmap from the clipboard device context to leave room for the copy
                    SelectObject clipboardDC, clipboardOldBMP
        
                    'Empty the current clipboard, copy this layer into it, then close the clipboard and delete the temporary
                    ' DC we created.
                    EmptyClipboard
                    SetClipboardData CLIPBOARD_FORMAT_BMP, clipboardBMP
                    
                    'Also, add a PNG-format copy of the image if GDI+ is available
                    
                    'If g_ImageFormats.GDIPlusEnabled Then
                  '
                  '      Dim tmpStream() As Byte
                  '      Dim tmpIIStream As IUnknown
                  '
                  '      Dim gdiCheck As Boolean
                  '      gdiCheck = GDIPlusSavePNGStream(CurrentImage, tmpStream, tmpIIStream)
                  '
                  '      Dim lPNG As Long
                  '      lPNG = RegisterClipboardFormat("PNG")
                  '
                  '      SetClipboardData lPNG, VarPtr(tmpIIStream)
                  '
                  '  End If
                                        
                    CloseClipboard
        
                    DeleteDC clipboardDC
                    
                End If
                
            End If
            
            'Release (DON'T DELETE!) our control of the current desktop device context
            ReleaseDC desktophWnd, desktopDC
            
        End If
        
    End If

End Sub

'Make a copy of an existing layer
Public Function createFromExistingLayer(ByRef srcLayer As pdLayer, Optional ByVal newWidth As Long = -1, Optional ByVal newHeight As Long = -1, Optional useHalftoning As Boolean = True) As Boolean

    'Make sure the layer we're passed isn't empty
    If srcLayer.getLayerDC <> 0 Then
        
        'Prepare new width and height values as requested by the user
        If newWidth = -1 Then newWidth = srcLayer.getLayerWidth
        If newHeight = -1 Then newHeight = srcLayer.getLayerHeight
        
        'If the width and height values are not being changed, the transfer is simple
        If (newWidth = srcLayer.getLayerWidth) And (newHeight = srcLayer.getLayerHeight) Then
        
            'Create a new, blank DIB the same size as the source layer
            If createBlank(srcLayer.getLayerWidth, srcLayer.getLayerHeight, srcLayer.getLayerColorDepth) Then
                'Copy the image data without modification
                BitBlt layerDC, 0, 0, layerWidth, layerHeight, srcLayer.getLayerDC, 0, 0, vbSrcCopy
                createFromExistingLayer = True
                Exit Function
            End If
        
        'If new width and height values are being specified, the transfer is a bit more complex
        Else
        
            'Create a new, blank DIB at the requested size
            If createBlank(newWidth, newHeight, srcLayer.getLayerColorDepth) Then

                'If either dimension of the new image will be smaller than the source, request halftoning
                If useHalftoning Then
                    SetStretchBltMode layerDC, STRETCHBLT_HALFTONE
                Else
                    SetStretchBltMode layerDC, STRETCHBLT_COLORONCOLOR
                End If

                'Resize and copy the image data
                StretchBlt layerDC, 0, 0, newWidth, newHeight, srcLayer.getLayerDC, 0, 0, srcLayer.getLayerWidth, srcLayer.getLayerHeight, vbSrcCopy
                
                'Now comes a nasty hack; HALFTONE stretching does not preserve the alpha channel, but COLORONCOLOR does.  So for 32bpp
                ' images, we now make a second copy of the original image using COLORONCOLOR - which means it contains valid alpha values.
                ' Then we copy those into our current image, effectively taking the best of both HALFTONE and COLORONCOLOR stretching.
                If srcLayer.getLayerColorDepth = 32 Then
                    
                    Dim hackLayer As pdLayer
                    Set hackLayer = New pdLayer
                    hackLayer.createBlank newWidth, newHeight, 32
            
                    SetStretchBltMode hackLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
                    StretchBlt hackLayer.getLayerDC, 0, 0, newWidth, newHeight, srcLayer.getLayerDC, 0, 0, srcLayer.getLayerWidth, srcLayer.getLayerHeight, vbSrcCopy
            
                    'At this point, hacklayer contains transparency data roughly equivalent to our half-toned images.  (Not perfect, but close.)
                    ' We need to copy that data into our image, then release the hack layer.
                    
                    'Start, as always, with a SafeArray
                    Dim iData() As Byte
                    Dim tmpSA As SAFEARRAY2D
                    prepInternalSafeArray tmpSA
                    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
                    'Now make a second array that points to the hacked layer alpha data
                    Dim aData() As Byte
                    Dim srcSA As SAFEARRAY2D
                    prepExternalSafeArray hackLayer, srcSA
                    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
                    Dim x As Long, y As Long, QuickX As Long
    
                    'Now loop through the image, copying alpha values as we go
                    For x = 0 To layerWidth - 1
                        QuickX = x * 4
                    For y = 0 To layerHeight - 1
                        iData(QuickX + 3, y) = aData(QuickX + 3, y)
                    Next y
                    Next x
                    
                    'Decommission both arrays
                    CopyMemory ByVal VarPtrArray(iData), 0&, 4
                    Erase iData
                    CopyMemory ByVal VarPtrArray(aData), 0&, 4
                    Erase aData
                    
                    'Release the temporary layer
                    hackLayer.eraseLayer
                    Set hackLayer = Nothing
    
                End If
                
                createFromExistingLayer = True
                Exit Function
                        
            End If
        
        End If
            
        createFromExistingLayer = False
    
    End If
    
    createFromExistingLayer = False

End Function

'Convert this layer to 24bpp mode
Public Function convertTo24bpp() As Boolean

    'Make sure this layer is 32bpp.  If it isn't, running this function pointless.
    If layerColorDepth = 32 Then

        'Make sure this layer isn't empty
        If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
            'Create a temporary layer to hold a copy of this layer's data (because it's about to get deleted)
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer Me
                        
            'Composite the temporary layer against a white background, per convention.
            tmpLayer.compositeBackgroundColor 255, 255, 255
                        
            'Now erase our own layer
            eraseLayer
    
            'Create a new layer that's exactly the same size as the old one
            If createBlank(tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, 24) Then
                
                'Copy the image data from the temporary layer without modification
                BitBlt layerDC, 0, 0, layerWidth, layerHeight, tmpLayer.getLayerDC, 0, 0, vbSrcCopy
                
                'Erase the temporary layer
                tmpLayer.eraseLayer
                Set tmpLayer = Nothing
                
                'Exit
                convertTo24bpp = True
                Exit Function
                
            End If
    
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong
    convertTo24bpp = False

End Function

'Convert this layer to 32bpp mode
Public Function convertTo32bpp() As Boolean

    'Make sure this layer is 24bpp.  If it isn't, running this function is pointless.
    If layerColorDepth = 24 Then

        'Make sure this layer isn't empty
        If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
            'Create a temporary layer to hold a copy of this layer's data (because it's about to get deleted)
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer Me
            
            'Now erase our own layer
            eraseLayer
    
            'Create a new layer that's exactly the same size as the old one
            If createBlank(tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, 32) Then
                
                'Copy the image data from the temporary layer without modification
                BitBlt layerDC, 0, 0, layerWidth, layerHeight, tmpLayer.getLayerDC, 0, 0, vbSrcCopy
                
                'Erase the temporary layer
                tmpLayer.eraseLayer
                Set tmpLayer = Nothing
                                
                'Finally, we need to loop through the image and set all alpha values to 255.  Otherwise the image will
                ' be completely transparent (and we don't want that!)
                Dim iData() As Byte
                Dim tmpSA As SAFEARRAY2D
                prepInternalSafeArray tmpSA
                CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
                Dim x As Long, y As Long, QuickX As Long
                
                'Loop through the image, setting the alpha of each pixel to 255 (opaque)
                For x = 0 To layerWidth - 1
                    QuickX = x * 4
                For y = 0 To layerHeight - 1
                    iData(QuickX + 3, y) = 255
                Next y
                Next x
    
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                
                'Exit
                convertTo32bpp = True
                Exit Function
                
            End If
    
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong
    convertTo32bpp = False

End Function

'Convert the current layer to 32bpp CMYK encoding (for export as CMYK via FreeImage)
Public Sub convertToCMYK32()
    
    'Make sure this layer isn't empty
    If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
        'If is layer isn't 32bpp, make it 32bpp
        If layerColorDepth <> 32 Then Me.convertTo32bpp
    
        'Prepare direct access to the DIB data
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                                
        Dim cyan As Long, magenta As Long, yellow As Long, K As Long
        
        Dim qvDepth As Long
        qvDepth = layerColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To layerWidth - 1
            QuickX = x * qvDepth
        For y = 0 To layerHeight - 1
                        
            'Cyan
            cyan = 255 - iData(QuickX + 2, y)
            
            'Magenta
            magenta = 255 - iData(QuickX + 1, y)
            
            'Yellow
            yellow = 255 - iData(QuickX, y)
            
            'Key
            K = CByte(MinimumInt(cyan, magenta, yellow))
            iData(QuickX + 3, y) = K
            
            If K = 255 Then
                iData(QuickX, y) = 0
                iData(QuickX + 1, y) = 0
                iData(QuickX + 2, y) = 0
            Else
                iData(QuickX, y) = cyan - K
                iData(QuickX + 1, y) = magenta - K
                iData(QuickX + 2, y) = yellow - K
            End If
                            
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        
        Erase iData
                        
    End If
    
End Sub

'Is this layer grayscale?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
Public Function isLayerGrayscale() As Boolean
    
    'Make sure this layer isn't empty
    If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
        'Loop through the image and compare RGB values to determine grayscale or not.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                        
        Dim r As Long, g As Long, b As Long
        
        Dim qvDepth As Long
        qvDepth = layerColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To layerWidth - 1
            QuickX = x * qvDepth
        For y = 0 To layerHeight - 1
            
            r = iData(QuickX + 2, y)
            g = iData(QuickX + 1, y)
            b = iData(QuickX, y)
            
            'For optimization reasons, this is stated as multiple IFs instead of an OR.  I'm not sure if VB
            ' short-circuits OR statements, but I know multiple IF statements work.  :)
            If r <> g Then
                
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
        
                isLayerGrayscale = False
                Exit Function
                
            ElseIf g <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isLayerGrayscale = False
                Exit Function
                
            ElseIf r <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isLayerGrayscale = False
                Exit Function
                
            End If
                
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        
        Erase iData
                        
        'Exit
        isLayerGrayscale = True
        Exit Function
        
    End If
    
    'If we made it to this line, the layer is blank, so it doesn't matter what value we return
    isLayerGrayscale = False

End Function

'Does this layer have "binary" transparency, e.g. does it have alpha values of only 0 or 255?
' (This is used to determine how transparency is handled when exporting to lower color depths.)
Public Function isAlphaBinary() As Boolean

    'Make sure this layer is 32bpp.  If it isn't, running this function is pointless.
    If layerColorDepth = 32 Then

        'Make sure this layer isn't empty
        If (layerDC <> 0) And (layerWidth <> 0) And (layerHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255.  If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
                
            'By default, assume that the image does not have a binary alpha channel.  (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For x = 0 To layerWidth - 1
                QuickX = x * 4
            For y = 0 To layerHeight - 1
                chkAlpha = iData(QuickX + 3, y)
                
                'For optimization reasons, this is stated as two IFs instead of an OR.  I'm not sure if VB
                ' short-circuits OR statements, but I know nested IF statements work.  :)
                If chkAlpha <> 255 Then
                
                    If chkAlpha <> 0 Then
                        notBinary = True
                        Exit For
                    End If
                
                End If
                
            Next y
                If notBinary = True Then Exit For
            Next x
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            Erase iData
                
            'Exit
            isAlphaBinary = Not notBinary
            Exit Function
                
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong
    isAlphaBinary = False

End Function

'Give this layer a picture from a standard VB picture object
Public Function CreateFromPicture(ByRef srcPicture As StdPicture, Optional forceWhiteBackground As Boolean = False) As Boolean

    'Make sure the picture we're passed isn't empty
    If (Not srcPicture Is Nothing) Then
    
        'Make sure the picture is actually a picture
        If GetObjectType(srcPicture) = OBJ_BITMAP Then
        
            'Select the picture's attributes into a bitmap object
            Dim tmpBitmap As Bitmap
            GetObject srcPicture.Handle, Len(tmpBitmap), tmpBitmap
            
            'Use that bitmap object to create a new, blank DIB of the same size
            If createBlank(tmpBitmap.Width, tmpBitmap.Height, tmpBitmap.BitsPerPixel) Then
            
                'Create a new DC
                Dim TmpDC As Long
                TmpDC = CreateCompatibleDC(0)
                
                'If successful, select the object into that DC
                If TmpDC <> 0 Then
                
                    'Temporary holder for the object selection
                    Dim oldBitmap As Long
                    oldBitmap = SelectObject(TmpDC, srcPicture.Handle)
                    
                    'Use BitBlt to copy the pixel data to this layer
                    BitBlt layerDC, 0, 0, layerWidth, layerHeight, TmpDC, 0, 0, vbSrcCopy
                    
                    'Now that we have the pixel data, erase all temporary objects
                    SelectObject TmpDC, oldBitmap
                    DeleteDC TmpDC
                    
                    'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                    ' black background.  We want the background to be white, so perform our own premultiplication.
                    ' (In the future, this could be a checkerboard pattern or a custom color instead of white.)
                    'NOTE! This is now handled at draw-time, and the instruction to re-render has been made optional.
                    If forceWhiteBackground And (layerColorDepth = 32) Then compositeBackgroundColor
                    
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

'Create a blank layer.  If no colorDepth is specified, it will default to 24bpp (16 million colors, no alpha-channel).
Public Function createBlank(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal colorDepth As Long = 24, Optional ByVal backColor As Long = vbWhite) As Boolean
    
    'Erase any existing layer data
    eraseLayer
    
    'PhotoDemon only supports 24 and 32 BPP at present
    If colorDepth <> 32 And colorDepth <> 24 Then
        colorDepth = 24
    End If
        
    'The back color may or may not be a system color, so translate it just in case
    backColor = TranslateColor(backColor)
        
    'Remember this color depth, width, and height
    layerColorDepth = colorDepth
    layerWidth = iWidth
    layerHeight = iHeight
    
    'Prepare the required header
    With layerDIBHeader
        .Size = Len(layerDIBHeader)
        .Planes = 1
        .BitCount = colorDepth
        .Width = iWidth
        .Height = -iHeight
        'As always, this value needs to be a multiple of four; with 32bpp that's automatic, with 24bpp it is not
        If colorDepth = 32 Then
            layerArrayWidth = 4 * iWidth
        Else
            layerArrayWidth = (iWidth * 3 + 3) And &HFFFFFFFC
        End If
        .ImageSize = layerArrayWidth * iHeight
    End With
    
    'Create a new DC for use with this layer
    layerDC = CreateCompatibleDC(0)
    
    If layerDC <> 0 Then
        
        'Create a DIB
        layerDIB = CreateDIBSection(layerDC, layerDIBHeader, 0, layerDIBits, 0, 0)
        
        'If successful, select the newly created dib into our DC
        If layerDIB <> 0 Then
        
            'We will later use layerDIBOriginal to clear up the memory associated with this layer
            layerDIBOriginal = SelectObject(layerDC, layerDIB)
            
            'Finally, set the backColor
            Dim layerRect As RECT
            SetRect layerRect, 0, 0, iWidth, iHeight
            
            Dim hBrush As Long
            hBrush = CreateSolidBrush(backColor)
            FillRect layerDC, layerRect, hBrush
            DeleteObject hBrush
        
        'If DIB creation failed, clear out the work we've done so far
        Else
            eraseLayer
        End If
        
    End If
    
    'Return success contingent on whether we have a DIB pointer or not
    createBlank = (layerDIB <> 0)
    
End Function

'This will effectively reset everything related to this layer, including image data.  Use cautiously!
Public Function eraseLayer()

    'If we have image data, clear it out
    If layerDC <> 0 Then
        If layerDIB <> 0 Then
            SelectObject layerDC, layerDIBOriginal
            DeleteObject layerDIB
        End If
        DeleteDC layerDC
    End If

    'Reset all associated DIB section variables
    layerDC = 0
    layerDIB = 0
    layerDIBOriginal = 0
    layerDIBits = 0

    'Reset layer size
    layerWidth = 0
    layerHeight = 0

End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Reset all associated DIB section variables
    layerDC = 0
    layerDIB = 0
    layerDIBOriginal = 0
    layerDIBits = 0
    
    'Reset layer size
    layerWidth = 0
    layerHeight = 0
    
End Sub

'TERMINATE class
Private Sub Class_Terminate()
    eraseLayer
End Sub

'Draw this layer to a picture box.  The image will be automatically resized and centered.
Public Sub renderToPictureBox(ByRef dstPicture As PictureBox)

    'Erase any existing picture
    dstPicture.Picture = LoadPicture("")

    'If the target picture box is smaller than this layer, request halftoning
    If (dstPicture.ScaleWidth < layerWidth) Or (dstPicture.ScaleHeight < layerHeight) Then
        SetStretchBltMode dstPicture.hDC, STRETCHBLT_HALFTONE
    Else
        SetStretchBltMode dstPicture.hDC, STRETCHBLT_COLORONCOLOR
    End If
    
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = layerWidth
    srcHeight = layerHeight
    
    'Calculate the aspect ratio of this layer and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = srcWidth / srcHeight
    dstAspect = dstWidth / dstHeight
    
    Dim dWidth As Long, dHeight As Long
    
    If srcAspect > dstAspect Then
        dWidth = dstWidth
        dHeight = CSng(srcHeight / srcWidth) * dWidth + 0.5
        PreviewY = CInt((dstHeight - dHeight) / 2)
        PreviewX = 0
        'PreviewWidth and PreviewHeight are only generated so that the "Print Preview" function works (it relies on those values
        ' to render its picture box).  They will be removed in a future update so DO NOT rely on them elsewhere.
        PreviewWidth = dWidth
        PreviewHeight = dHeight
        StretchBlt dstPicture.hDC, 0, PreviewY, dWidth, dHeight, layerDC, 0, 0, layerWidth, layerHeight, vbSrcCopy
    Else
        dHeight = dstHeight
        dWidth = CSng(srcWidth / srcHeight) * dHeight + 0.5
        PreviewX = CInt((dstWidth - dWidth) / 2)
        PreviewY = 0
        PreviewWidth = dWidth
        PreviewHeight = dHeight
        StretchBlt dstPicture.hDC, PreviewX, 0, dWidth, dHeight, layerDC, 0, 0, layerWidth, layerHeight, vbSrcCopy
    End If
    
    dstPicture.Picture = dstPicture.Image
    dstPicture.Refresh

End Sub

'Load a layer's DIB information from file.
Public Sub createFromFile(ByRef srcFilename As String)
                
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Open the file and dump out only the essential information
    Open srcFilename For Binary As #fileNum
    
        'Check to make sure this is actually a layer file
        Dim LayerIDCheck As String * 4
        Get #fileNum, 1, LayerIDCheck
        If (LayerIDCheck <> LAYER_IDENTIFIER) Then
            Close #fileNum
            Message "Failed to load layer from disk: invalid layer file specified."
            Exit Sub
        End If
    
        'Now check to make sure that the version number is supported (not implemented right now, because there's only one version)
        Dim LayerVersionCheck As Long
        Get #fileNum, , LayerVersionCheck
                        
        'Get color depth
        Dim fColorDepth As Long
        Get #fileNum, , fColorDepth
        
        'Get size
        Dim fWidth As Long, fHeight As Long, fArrayWidth As Long
        Get #fileNum, , fWidth
        Get #fileNum, , fHeight
        Get #fileNum, , fArrayWidth
        
        'Is it compressed?  (Not implemented now; might be in the future)
        Dim toUncompress As Boolean
        Get #fileNum, , toUncompress
        
        'If we've made it this far, attempt to load pixel data.  Start by erasing and re-initializing this object to the proper size.
        createBlank fWidth, fHeight, fColorDepth
        
        'Now pull the pixel data from the file
        Dim fileData() As Byte
        ReDim fileData(0 To fArrayWidth - 1, 0 To fHeight - 1) As Byte
        Get #fileNum, , fileData
        
    'We're done with the file at this point, so close it
    Close #fileNum
    
    'Now it's time to overwrite our current DIB with the pixel data we pulled from the file.
    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Bulk copy the contents of fileData() over to iData() - which is pointing at this layer's DIB, remember
    CopyMemory iData(0, 0), fileData(0, 0), fArrayWidth * fHeight
    
    'With our work complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    
End Sub

'Write this layer's DIB information to file.
Public Sub writeToFile(ByRef dstFilename As String)

    'Delete any existing file (overwrite)
    If FileExist(dstFilename) = True Then Kill dstFilename
        
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Open the file and dump out only the essential information
    Open dstFilename For Binary As #fileNum
    
        'Identifiers
        Put #fileNum, 1, LAYER_IDENTIFIER
        Put #fileNum, , LAYER_FILE_VERSION_2012
        
        'Color depth
        Put #fileNum, , layerColorDepth
        
        'Size
        Put #fileNum, , layerWidth
        Put #fileNum, , layerHeight
        Put #fileNum, , layerArrayWidth
        
        'Is it compressed?  (Not implemented now; might be in the future)
        Dim toCompress As Boolean
        toCompress = False
        Put #fileNum, , toCompress
        
        'And finally, the pixel data, which is presently uncompressed
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        'Now dump the entire array into the file.  To my knowledge, this is the fastest way to do this in VB.
        Put #fileNum, , iData
        
    Close #fileNum
    
    'With our work complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

End Sub

'Write this layer's DIB information to a valid BMP file.
Public Sub writeToBitmapFile(ByRef dstFilename As String)

    'First, make sure this layer actually contains image data
    If Not Me.hasImage() Then Exit Sub
        
    'Calculate the size of a scanline
    Dim slWidth As Long
    slWidth = ((layerWidth * layerColorDepth + 31) \ 32) * 4
    
    'A valid bitmap header consists of two parts: a file header, and an image header.
    Dim bmpFileHeader As BITMAPFILEHEADER
    Dim bmpInfoHeader As BITMAPINFO
    
    'Build the file header first
    With bmpFileHeader
        .Type = &H4D42      'BMP identifier
        .Size = Len(bmpFileHeader) + (slWidth * layerHeight)      'Length of the file
        .OffBits = Len(bmpFileHeader) + Len(bmpInfoHeader)          'Length of the header area
    End With
    
    '...then the image header
    With bmpInfoHeader.Header
        .Size = 40
        .Planes = 1
        .BitCount = layerColorDepth
        .Width = layerWidth
        .Height = layerHeight
    End With
    
    'Finally, the image bytes
    Dim iData() As Byte
    ReDim iData(0 To slWidth - 1, 0 To layerHeight - 1) As Byte
    GetDIBits layerDC, layerDIB, 0, layerHeight, iData(0, 0), bmpInfoHeader, 0
    
    'With all our information in place, check to see if that file already exists.  Delete it if necessary.
    If FileExist(dstFilename) = True Then Kill dstFilename
        
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Open the file and dump the two headers and image data into it
    Open dstFilename For Binary As #fileNum
    
        'Headers
        Put #fileNum, 1, bmpFileHeader
        Put #fileNum, , bmpInfoHeader
        
        'Image data
        Put #fileNum, , iData()
        
    Close #fileNum
    
    Erase iData
    
End Sub

'Sometimes a layer needs to access its own bits.  Here's how.
Private Sub prepInternalSafeArray(ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = layerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = layerArrayWidth
        .pvData = layerDIBits
    End With
End Sub

'Sometimes a layer needs to access the bits of other layers.  Here's how.
Private Sub prepExternalSafeArray(ByRef srcLayer As pdLayer, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
End Sub

'Pre-composite an image with an alpha-channel against a background color.  Until PhotoDemon is capable of rendering transparent
' images itself, this is necessary to give transparent images a white background.
Public Sub compositeBackgroundColor(Optional ByVal newR As Long = -1, Optional ByVal newG As Long = -1, Optional ByVal newB As Long = -1)

    'This is only useful for images with alpha channels.  Exit if no alpha channel is present.
    If layerColorDepth <> 32 Then Exit Sub

    'DEBUG: These variables can be used to run timing tests on the composite routine.
    'Dim TotalTime As Long
    'Dim StartTime As Long
    'Dim numOfRuns As Long
    'StartTime = GetTickCount

    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'If no color is specified, apply a checkerboard pattern
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
        
    If newR <> -1 Then
        r1 = newR
        r2 = newR
    Else
        r1 = ExtractR(g_AlphaCheckOne)
        r2 = ExtractR(g_AlphaCheckTwo)
    End If
    
    If newG <> -1 Then
        g1 = newG
        g2 = newG
    Else
        g1 = ExtractG(g_AlphaCheckOne)
        g2 = ExtractG(g_AlphaCheckTwo)
    End If
    
    If newB <> -1 Then
        b1 = newB
        b2 = newB
    Else
        b1 = ExtractB(g_AlphaCheckOne)
        b2 = ExtractB(g_AlphaCheckTwo)
    End If
    
    'Make sure the user didn't supply us with bad values
    If r1 < 0 Then r1 = 0
    If r2 < 0 Then r2 = 0
    If g1 < 0 Then g1 = 0
    If g2 < 0 Then g2 = 0
    If b1 < 0 Then b1 = 0
    If b2 < 0 Then b2 = 0
    
    If r1 > 255 Then r1 = 255
    If r2 > 255 Then r2 = 255
    If g1 > 255 Then g1 = 255
    If g2 > 255 Then g2 = 255
    If b1 > 255 Then b1 = 255
    If b2 > 255 Then b2 = 255
    
    'Finally, determine a checkerboard size based on the current user preference
    Dim checkerSize As Byte
    
    Select Case g_AlphaCheckSize
    
        'Small (4x4 checks)
        Case 0
            checkerSize = 4
            
        'Medium (8x8 checks)
        Case 1
            checkerSize = 8
        
        'Large (16x16 checks)
        Case Else
            checkerSize = 16
        
    End Select
    
    'Loop through the image, blending colors as we go
    For x = 0 To layerWidth - 1
        QuickX = x * 4
    For y = 0 To layerHeight - 1
        
        'Access the alpha data for this pixel
        checkAlpha = iData(QuickX + 3, y)
        
        'Ignore fully opaque pixels.  (This makes the routine much faster.)
        If checkAlpha <> 255 Then
            
            'Convert the alpha value to a floating-point variable
            tmpAlpha = CSng(checkAlpha) / 255
            
            'Use that alpha value to blend the current colors with the newly requested color, in a checkerboard pattern
            If (((x \ checkerSize) + (y \ checkerSize)) And 1) = 0 Then
                iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), r1, tmpAlpha)
                iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), g1, tmpAlpha)
                iData(QuickX, y) = Blend2Colors(iData(QuickX, y), b1, tmpAlpha)
            Else
                iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), r2, tmpAlpha)
                iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), g2, tmpAlpha)
                iData(QuickX, y) = Blend2Colors(iData(QuickX, y), b2, tmpAlpha)
            End If
            
        End If
        
    Next y
    Next x
    
    'With our compositing complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    
    'TotalTime = TotalTime + (GetTickCount - StartTime)
    'numOfRuns = numOfRuns + 1
    'Message "Composite average: " & CDbl(TotalTime) / numOfRuns

End Sub

'Blend byte1 w/ byte2 based on mixRatio.  mixRatio is expected to be a value between 0 and 1.
Private Function Blend2Colors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByRef mixRatio As Double) As Byte
    Blend2Colors = (mixRatio * Color1) + ((1 - mixRatio) * Color2)
    'Alternate blend formula from Robert Rayment - my own testing shows no measurable speed difference between the two methods
    'Blend2Colors = mixRatio * (CInt(Color1) - Color2) + Color2
End Function


'Pre-composite an image with an alpha-channel against a background color.  Until PhotoDemon is capable of rendering transparent
' images itself, this is necessary to give transparent images a white background.
Public Sub compositeBackgroundColorSpecial(ByRef srcAlpha As pdLayer, Optional ByVal newR As Long = -1, Optional ByVal newG As Long = -1, Optional ByVal newB As Long = -1)

    'This is only useful for images with alpha channels.  Exit if no alpha channel is present.
    If layerColorDepth <> 32 Then Exit Sub

    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Now make a second array that points to the source alpha data
    Dim aData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepExternalSafeArray srcAlpha, srcSA
    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'If no color is specified, apply a checkerboard pattern
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    If newR <> -1 Then
        r1 = newR
        r2 = newR
    Else
        r1 = ExtractR(g_AlphaCheckOne)
        r2 = ExtractR(g_AlphaCheckTwo)
    End If
    
    If newG <> -1 Then
        g1 = newG
        g2 = newG
    Else
        g1 = ExtractG(g_AlphaCheckOne)
        g2 = ExtractG(g_AlphaCheckTwo)
    End If
    
    If newB <> -1 Then
        b1 = newB
        b2 = newB
    Else
        b1 = ExtractB(g_AlphaCheckOne)
        b2 = ExtractB(g_AlphaCheckTwo)
    End If
    
    'Make sure the user didn't supply us with bad values
    If r1 < 0 Then r1 = 0
    If r2 < 0 Then r2 = 0
    If g1 < 0 Then g1 = 0
    If g2 < 0 Then g2 = 0
    If b1 < 0 Then b1 = 0
    If b2 < 0 Then b2 = 0
    
    If r1 > 255 Then r1 = 255
    If r2 > 255 Then r2 = 255
    If g1 > 255 Then g1 = 255
    If g2 > 255 Then g2 = 255
    If b1 > 255 Then b1 = 255
    If b2 > 255 Then b2 = 255
    
    'Finally, determine a checkerboard size based on the current user preference
    Dim checkerSize As Byte
    
    Select Case g_AlphaCheckSize
    
        'Small (4x4 checks)
        Case 0
            checkerSize = 4
            
        'Medium (8x8 checks)
        Case 1
            checkerSize = 8
        
        'Large (16x16 checks)
        Case Else
            checkerSize = 16
        
    End Select
    
    'Loop through the image, blending colors as we go
    For x = 0 To layerWidth - 1
        QuickX = x * 4
    For y = 0 To layerHeight - 1
        
        'Access the alpha data for this pixel
        checkAlpha = aData(QuickX + 3, y)
        
        'Ignore fully opaque pixels.  (This makes the routine much faster.)
        If checkAlpha <> 255 Then
            
            'Convert the alpha value to a floating-point variable
            tmpAlpha = CSng(checkAlpha) / 255
            
            'Use that alpha value to blend the current colors with the newly requested color, in a checkerboard pattern
            If (((x \ checkerSize) + (y \ checkerSize)) And 1) = 0 Then
                iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), r1, tmpAlpha)
                iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), g1, tmpAlpha)
                iData(QuickX, y) = Blend2Colors(iData(QuickX, y), b1, tmpAlpha)
            Else
                iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), r2, tmpAlpha)
                iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), g2, tmpAlpha)
                iData(QuickX, y) = Blend2Colors(iData(QuickX, y), b2, tmpAlpha)
            End If
            
        End If
        
    Next y
    Next x
    
    'With our compositing complete, point both arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    CopyMemory ByVal VarPtrArray(aData), 0&, 4
    Erase aData

End Sub

'Check to see if a 32bpp image is really 32bpp. (Basically, scan all pixels in the alpha channel.  If all values are set to
' 255 or all values are set to 0, remove the alpha channel and rebuild this layer in 24bpp mode.)
Public Function verifyAlphaChannel(Optional ByVal newR As Long = -1, Optional ByVal newG As Long = -1, Optional ByVal newB As Long = -1) As Boolean

    'This is only useful for images with alpha channels.  Exit if no alpha channel is present.
    If layerColorDepth <> 32 Then
        verifyAlphaChannel = True
        Exit Function
    End If
    
    'This routine will fail if the width or height of this layer is 0
    If layerWidth = 0 Or layerHeight = 0 Then
        verifyAlphaChannel = True
        Exit Function
    End If

    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    Dim checkAlpha As Boolean, initAlpha As Double
    checkAlpha = False
    
    'Determine the alpha value of the top-left pixel.  This will be used as our baseline value.
    initAlpha = iData(3, 0)
    
    'If initAlpha is something other than 255 or 0, we don't need to check the image
    If (initAlpha <> 0) And (initAlpha <> 255) Then
        verifyAlphaChannel = True
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        Exit Function
    End If
        
    'Loop through the image, comparing colors as we go
    For x = 0 To layerWidth - 1
        QuickX = x * 4
    For y = 0 To layerHeight - 1
        
        'Compare the alpha data for this pixel to the initial pixel.  If they DO NOT match, this is a valid alpha channel.
        If initAlpha <> iData(QuickX + 3, y) Then
            checkAlpha = True
            Exit For
        End If
        
    Next y
    
        'If the alpha channel has been verified, exit this loop
        If checkAlpha Then Exit For
        
    Next x
    
    'With our check complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

    'Return checkAlpha.  If varying alpha values were found, this function returns TRUE.  If all values were the same,
    ' this function returns FALSE.
    verifyAlphaChannel = checkAlpha

End Function

'Copy the alpha values from another pdLayer object to this one.  This is useful when the alpha channel for this layer must be lost
' during a transformation (typically something involving FreeImage or GDI+), and a temporary layer was made to preserve the alpha
' data.  If the image is currently 24bpp, this function will first convert it to 32bpp before copying the alpha data.

'IMPORTANT NOTE: to keep this function fast, NO ARRAY BOUNDS CHECKING IS DONE.  Make sure that the passed pdLayer object is
'                 THE SAME SIZE (or larger) than this layer, or you will experience critical errors.
Public Sub copyAlphaFromExistingLayer(ByRef srcLayer As pdLayer)

    'This is only useful for images with alpha channels.  Exit if no alpha channel is present.
    If layerColorDepth <> 32 Then convertTo32bpp

    'Prepare a SafeArray that points to our own DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Now make a second array that points to the source DIB data
    Dim aData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepExternalSafeArray srcLayer, srcSA
    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    'Loop through the image, copying alpha values as we go
    For x = 0 To layerWidth - 1
        QuickX = x * 4
    For y = 0 To layerHeight - 1
        
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

'Layers created from GDI+ contain pre-multiplied alpha values.  These are great for high-performance blting, but terrible for
' photo editing.  This routine can be used to restore (well, estimate) the original color values of the image.
Public Sub fixPremultipliedAlpha()

    'Prepare a SAFEARRAY, which allows us to access the DIB data directly
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tmpAlpha As Byte, tmpAlphaModifier As Double

    'Loop through the image, removing the effects of pre-multiplication as we go
    For x = 0 To layerWidth - 1
        QuickX = x * 4
    For y = 0 To layerHeight - 1
        
        tmpAlpha = iData(QuickX + 3, y)
        
        'Ignore completely opaque and completely transparent pixels
        If tmpAlpha <> 255 Then
            If tmpAlpha <> 0 Then
            
                r = iData(QuickX + 2, y)
                g = iData(QuickX + 1, y)
                b = iData(QuickX, y)
            
                tmpAlphaModifier = 255 / tmpAlpha
            
                'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                r = r * tmpAlphaModifier
                g = g * tmpAlphaModifier
                b = b * tmpAlphaModifier
                
                iData(QuickX + 2, y) = CByte(r)
                iData(QuickX + 1, y) = CByte(g)
                iData(QuickX, y) = CByte(b)
            
            End If
        End If
    Next y
    Next x
    
    'With our alpha channel complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

End Sub
