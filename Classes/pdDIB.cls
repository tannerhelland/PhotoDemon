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
'Copyright ©2012-2014 by Tanner Helland
'Created: 29/August/12
'Last updated: 14/April/14
'Last update: greatly improved the .renderToPictureBox function for 32bpp images; GDI+ is now used for high-quality resampling.
'
'This marvelous class is PhotoDemon's replacement for picture boxes (which it used to store images in previous versions).
' pdDIB is a powerful DIB class responsible for managing all image data in memory. Some of the benefits provided by
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
'Note that anything you can do with an hDC property you can do with this class - simply use the getDIBDC function to
' return the DIB's hDC, then do with it what you please. All functions are heavily commented and should be self-explanatory.
'
'Note also that this class is treated as a subset of pdImage(). Right now each pdImage object only contains one layer
' (the image itself), but in the future a pdImage object could theoretically store many layers.
'
'SPECIAL THANKS
'In building this class, I utilized a number of other DIB classes for reference and testing. Special thanks to:
' Carles PV's iBMP project: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Steve McMahon's DIB/SafeArray analysis: http://www.vbaccelerator.com/home/VB/Code/vbMedia/DIB_Sections/True_Colour_DIBSection/article.asp
' Many thanks to these three individuals for their outstanding work on graphics in VB.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
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
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DIB API functions
' (Note that these are currently declared in FastDrawing as well)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

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
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const CLIPBOARD_FORMAT_BMP As Long = 2

'AlphaBlend API call
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'Variables related to the DIB
'hDC for this DIB
Private dibDC As Long
'Actual DIB handle for this DIB
Private dibHandle As Long
'Original handle when this DIB is first created (we must store this so we can properly clean up the DIB when we're finished)
Private dibHandleOriginal As Long
'Pointer to the actual DIB bits
Private dibBits As Long
'Persistent DIB header; this will be used with WAPI to create the DIB initially
Private dibHeader As BITMAPINFOHEADER

'The DIB's width and height
Private dibWidth As Long, dibHeight As Long

'The DIB's array width (dibWidth * 4 for 32-bit, varies for 24-bit due to DWORD-alignment)
Private dibArrayWidth As Long

'The DIB's color depth (should only ever be 24 or 32)
Private dibColorDepth As Long

'Used when writing/reading DIB data to/from a file
Private Const DIB_IDENTIFIER As String * 4 = "PDlr"
Private Const DIB_FILE_VERSION_2012 As Long = &H1000

'When converting a 32bpp DIB to 8bpp with transparency, we need to know the location of a transparent pixel.
' This location is passed to FreeImage, and it designates the palette index at that location as the transparent index.
Private trnsX As Long, trnsY As Long

'Remember the original color of the pixels made transparent, then restore it after the "magic magenta" has served its purpose
Private originalTransparentColor As Long

'If this DIB was loaded from file, its original file format will be set here.  (By default, this value will be -1).
Private originalFIF As Long

'DIBs may require ICC profile handling.  All ICC Profile interactions are handled through a publicly available ICCProfile class,
' on a per-DIB basis.
Public ICCProfile As pdICCProfile

'DIBs created from loaded image files contain resolution data.  This data isn't important on a per-DIB basis, but the parent
' pdImage object will copy the resolution data from its first-loaded child DIB (if it has a resolution)
Private xResolution As Double, yResolution As Double, dibDPI As Double

'Some image formats (e.g. PNG) will specify a preferred background color.  If available at load-time, that color value
' will be stored here.
Private originalBackgroundColor As Long

'DIBs created from loaded image files will have their original color depth stored here.  Note that PD only works in 24/32 bpp
' mode at present, but this value may contain other bit-depths (e.g. 8bpp if the source data was a GIF).
Private originalColorDepth As Long

'Get/set original color depth
Public Function getOriginalColorDepth() As Long
    getOriginalColorDepth = originalColorDepth
End Function

Public Sub setOriginalColorDepth(ByVal origColorDepth As Long)
    originalColorDepth = origColorDepth
End Sub

'If this DIB was created from an image file, and that file specified a preferred background color, use this
' function to copy the value into this DIB.
Public Sub setBackgroundColor(ByVal newBackColor As Long)
    originalBackgroundColor = newBackColor
End Sub

Public Function getBackgroundColor() As Long
    getBackgroundColor = originalBackgroundColor
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
    
    xResolution = xRes
    yResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    dibDPI = (xRes + yRes) \ 2

End Sub

'Even though we store separate x and y DPI, PD only deals in a single per-DIB DPI value
Public Function getDPI() As Double
    getDPI = dibDPI
End Function

'Get/Set original DIB format
Public Function getOriginalFormat() As Long
    getOriginalFormat = originalFIF
End Function

Public Sub setOriginalFormat(ByVal origFIF As Long)
    originalFIF = origFIF
End Sub

'Return the original transparent color (set as part of the MarkAlpha routine).
Public Function getOriginalTransparentColor() As Long
    getOriginalTransparentColor = originalTransparentColor
End Function

'Return the location of a transparent pixel. This will be set as part of the MarkAlpha routine.
Public Sub getTransparentLocation(ByRef srcX As Long, ByRef srcY As Long)
    srcX = trnsX
    srcY = trnsY
End Sub

'This function will apply an "alpha-cutoff" to a 32bpp image. Any alpha values above the cutoff will be replaced with
' magic magenta (253,0,253) and the location of a transparent pixel (trnsX, trnsY) will be marked.
Public Sub applyAlphaCutoff(Optional ByVal cutOff As Long = 127, Optional ByVal forceColor As Boolean = True, Optional ByVal newBackgroundColor As Long = 16777215)

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If dibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then
            
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
            
            trnsX = -1
            
            'Retrieve RGB values from the new background color, which we'll use to composite semi-transparent pixels
            Dim backR As Byte, backG As Byte, backB As Byte
            backR = CByte(ExtractR(newBackgroundColor))
            backG = CByte(ExtractG(newBackgroundColor))
            backB = CByte(ExtractB(newBackgroundColor))
                
            'Loop through the image, checking alphas as we go
            For x = 0 To dibWidth - 1
                QuickX = x * 4
            For y = 0 To dibHeight - 1
                
                chkAlpha = iData(QuickX + 3, y)
                
                'If the alpha value is less than the cutoff, mark this pixel
                If chkAlpha < cutOff Then
                
                    'Remember this location if we haven't already
                    If trnsX = -1 Then
                        trnsX = x
                        trnsY = y
                        originalTransparentColor = RGB(iData(QuickX + 2, y), iData(QuickX + 1, y), iData(QuickX, y))
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

'Translate an OLE color to an RGB Long
Private Function TranslateColor(ByVal colorRef As Long) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, TranslateColor) Then
        TranslateColor = RGB(255, 255, 255)
    End If
End Function

'Return this DIB's color depth
Public Function getDIBColorDepth() As Long
    getDIBColorDepth = dibColorDepth
End Function

'Return this DIB's array width
Public Function getDIBArrayWidth() As Long
    getDIBArrayWidth = dibArrayWidth
End Function

'Return this DIB's width
Public Function getDIBWidth() As Long
    getDIBWidth = dibWidth
End Function

'Return this DIB's height
Public Function getDIBHeight() As Long
    getDIBHeight = dibHeight
End Function

'Return whether or not this DIB has image data associated with it
Public Function hasImage() As Boolean
    hasImage = (dibHandle <> 0)
End Function

'Return this DIB's hDC
Public Function getDIBDC() As Long
    getDIBDC = dibDC
End Function

'Return a pointer to this DIB's actual DIB
Public Function getDIBHandle() As Long
    getDIBHandle = dibHandle
End Function

'Return a pointer to this DIB's pixel data (no header)
Public Function getActualDIBBits() As Long
    getActualDIBBits = dibBits
End Function

'Return a pointer to this DIB's header
Public Function getDIBHeader() As Long
    getDIBHeader = VarPtr(dibHeader)
End Function

'Copy the current DIB's contents to the clipboard
Public Sub copyDIBToClipboard()
    
    'Make sure the DIB actually contains an image before copying it to the clipboard
    If (dibHandle <> 0) Then
    
        'We are going to copy the image data to the clipboard twice - once in PNG format, then again in standard BMP format.
        ' This maxmimizes operability between major software packages.
        
        'Start by using the wonderful vbAccelerator clipboard class, which will handle the messiest parts for us.
        Dim clpObject As cCustomClipboard
        Set clpObject = New cCustomClipboard
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            clpObject.ClearClipboard
        
            'FreeImage is required to perform the PNG transformation.  We could use GDI+, but FreeImage is easier to interface with.
            If g_ImageFormats.FreeImageEnabled And dibColorDepth = 32 Then
            
                Dim PNGID As Long
                PNGID = clpObject.AddFormat("PNG")
                
                'Convert our current DIB to a FreeImage-type DIB
                Dim fi_DIB As Long
                fi_DIB = FreeImage_CreateFromDC(dibDC)
                
                'Convert the bitmap to PNG format, save it to an array, and release the original bitmap from memory
                Dim pngArray() As Byte
                Dim fi_Check As Long
                fi_Check = FreeImage_SaveToMemoryEx(FIF_PNG, fi_DIB, pngArray, FISO_PNG_Z_DEFAULT_COMPRESSION, True)
                
                'Now, hand over the filled PNG array to the clipboard
                clpObject.SetBinaryData PNGID, pngArray
                
            End If
            
            'With a PNG copy successfully saved, proceed to copy a standard 24bpp bitmap to the clipboard
            
            'Get a handle to the current desktop, and create a compatible clipboard device context in it
            Dim desktophWnd As Long
            desktophWnd = GetDesktopWindow
            
            Dim desktopDC As Long, clipboardDC As Long
            desktopDC = GetDC(desktophWnd)
            clipboardDC = CreateCompatibleDC(desktopDC)
            
            'If our temporary DC was created successfully, use it to create a temporary bitmap for the clipboard
            If (clipboardDC <> 0) Then
            
                'Create a bitmap compatible with the current desktop. This will receive the actual pixel data of the current DIB.
                Dim clipboardBMP As Long
                clipboardBMP = CreateCompatibleBitmap(desktopDC, dibWidth, dibHeight)
                
                If (clipboardBMP <> 0) Then
                    
                    'Place the compatible bitmap within the clipboard device context
                    Dim clipboardOldBMP As Long
                    clipboardOldBMP = SelectObject(clipboardDC, clipboardBMP)
                    
                    'Composite the DIB against a white background, then use BitBlt to paint it to the clipboard bitmap
                    Me.convertTo24bpp
                    BitBlt clipboardDC, 0, 0, dibWidth, dibHeight, dibDC, 0, 0, vbSrcCopy
                    
                    'Remove that bitmap from the clipboard device context to leave room for the copy
                    SelectObject clipboardDC, clipboardOldBMP
        
                    'Copy the bitmap to the clipboard, then close and exit
                    clpObject.SetClipboardMemoryHandle CLIPBOARD_FORMAT_BMP, clipboardBMP
                    
                    DeleteDC clipboardDC
                    
                End If
                
            End If
            
            'Release (DON'T DELETE!) our control of the current desktop device context
            ReleaseDC desktophWnd, desktopDC
            
            'Release our hold on the clipboard
            clpObject.ClipboardClose
            
        End If
        
    End If

End Sub

'Make a copy of an existing DIB
Public Function createFromExistingDIB(ByRef srcDIB As pdDIB, Optional ByVal newWidth As Long = -1, Optional ByVal newHeight As Long = -1, Optional useHalftoning As Boolean = True) As Boolean

    'Make sure the DIB we're passed isn't empty
    If srcDIB.getDIBDC <> 0 Then
        
        'Prepare new width and height values as requested by the user
        If newWidth = -1 Then newWidth = srcDIB.getDIBWidth
        If newHeight = -1 Then newHeight = srcDIB.getDIBHeight
        
        'If the width and height values are not being changed, the transfer is simple
        If (newWidth = srcDIB.getDIBWidth) And (newHeight = srcDIB.getDIBHeight) Then
        
            'Create a new, blank DIB the same size as the source DIB
            If createBlank(srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBColorDepth) Then
                'Copy the image data without modification
                BitBlt dibDC, 0, 0, dibWidth, dibHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
                createFromExistingDIB = True
                Exit Function
            End If
        
        'If new width and height values are being specified, the transfer is a bit more complex
        Else
        
            'Create a new, blank DIB at the requested size
            If createBlank(newWidth, newHeight, srcDIB.getDIBColorDepth) Then

                'We now break up the resize request into two branches: 24bpp data, and 32bpp data.
                If srcDIB.getDIBColorDepth = 24 Then
                
                    'If either dimension of the new image will be smaller than the source, request halftoning
                    If useHalftoning Then
                        SetStretchBltMode dibDC, STRETCHBLT_HALFTONE
                    Else
                        SetStretchBltMode dibDC, STRETCHBLT_COLORONCOLOR
                    End If
    
                    'Resize and copy the image data
                    StretchBlt dibDC, 0, 0, newWidth, newHeight, srcDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, vbSrcCopy
                    
                '32bpp data requires special attention, because StretchBlt won't halftone alpha data.
                Else
                
                    'Instead of halftoning, switch between interpolation and nearest-neighbor
                    If useHalftoning Then
                        GDIPlusResizeDIB Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, InterpolationModeHighQualityBicubic
                    Else
                        GDIPlusResizeDIB Me, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, InterpolationModeNearestNeighbor
                    End If
                
                End If
                
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
    If dibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then
    
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB Me
                        
            'Composite the temporary DIB against a white background, per convention.
            tmpDIB.compositeBackgroundColor ExtractR(newBackColor), ExtractG(newBackColor), ExtractB(newBackColor)
            
            'Now erase our own DIB
            eraseDIB
            
            'Create a new DIB that's exactly the same size as the old one
            If createBlank(tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, 24) Then
                
                'Copy the image data from the temporary DIB without modification
                BitBlt dibDC, 0, 0, dibWidth, dibHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
                
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
    If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then

        'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createFromExistingDIB Me
        
        'Now erase our own DIB
        eraseDIB

        'Create a new DIB that's exactly the same size as the old one
        If createBlank(tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, 32) Then
            
            'Copy the image data from the temporary DIB without modification
            BitBlt dibDC, 0, 0, dibWidth, dibHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
                            
            'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
            ' be completely transparent (and we don't want that!)
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4

            Dim x As Long, y As Long, QuickX As Long
            
            Dim newTransFloat As Double
            newTransFloat = newTransparency / 255
            
            'Loop through the image, setting the alpha of each pixel to 255 (opaque)
            For x = 0 To dibWidth - 1
                QuickX = x * 4
            For y = 0 To dibHeight - 1
                
                If tmpDIB.getDIBColorDepth = 24 Then
                    iData(QuickX + 3, y) = newTransparency
                Else
                    iData(QuickX + 3, y) = iData(QuickX + 3, y) * newTransFloat
                End If
                
            Next y
            Next x

            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            'Exit
            convertTo32bpp = True
            Exit Function
            
        End If

    End If
    
    
    'If we made it to this line, something went horribly wrong
    convertTo32bpp = False

End Function

'Convert the current DIB to 32bpp CMYK encoding (for export as CMYK via FreeImage)
Public Sub convertToCMYK32()
    
    'Make sure this DIB isn't empty
    If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then
    
        'If is DIB isn't 32bpp, make it 32bpp
        If dibColorDepth <> 32 Then Me.convertTo32bpp
    
        'Prepare direct access to the DIB data
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                                
        Dim cyan As Long, magenta As Long, yellow As Long, k As Long
        
        Dim qvDepth As Long
        qvDepth = dibColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To dibWidth - 1
            QuickX = x * qvDepth
        For y = 0 To dibHeight - 1
                        
            'Cyan
            cyan = 255 - iData(QuickX + 2, y)
            
            'Magenta
            magenta = 255 - iData(QuickX + 1, y)
            
            'Yellow
            yellow = 255 - iData(QuickX, y)
            
            'Key
            k = CByte(Min3Int(cyan, magenta, yellow))
            iData(QuickX + 3, y) = k
            
            If k = 255 Then
                iData(QuickX, y) = 0
                iData(QuickX + 1, y) = 0
                iData(QuickX + 2, y) = 0
            Else
                iData(QuickX, y) = cyan - k
                iData(QuickX + 1, y) = magenta - k
                iData(QuickX + 2, y) = yellow - k
            End If
                            
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        
        Erase iData
                        
    End If
    
End Sub

'Is this DIB grayscale? Determination is made by scanning each pixel and comparing RGB values to see if they match.
Public Function isDIBGrayscale() As Boolean
    
    'Make sure this DIB isn't empty
    If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then
    
        'Loop through the image and compare RGB values to determine grayscale or not.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                        
        Dim r As Long, g As Long, b As Long
        
        Dim qvDepth As Long
        qvDepth = dibColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To dibWidth - 1
            QuickX = x * qvDepth
        For y = 0 To dibHeight - 1
            
            r = iData(QuickX + 2, y)
            g = iData(QuickX + 1, y)
            b = iData(QuickX, y)
            
            'For optimization reasons, this is stated as multiple IFs instead of an OR. I'm not sure if VB
            ' short-circuits OR statements, but I know multiple IF statements work. :)
            If r <> g Then
                
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
        
                isDIBGrayscale = False
                Exit Function
                
            ElseIf g <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            ElseIf r <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            End If
                
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        
        Erase iData
                        
        'Exit
        isDIBGrayscale = True
        Exit Function
        
    End If
    
    'If we made it to this line, the DIB is blank, so it doesn't matter what value we return
    isDIBGrayscale = False

End Function

'Does this DIB have "binary" transparency, e.g. does it have alpha values of only 0 or 255?
' (This is used to determine how transparency is handled when exporting to lower color depths.)
Public Function isAlphaBinary() As Boolean

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If dibColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (dibDC <> 0) And (dibWidth <> 0) And (dibHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
                
            'By default, assume that the image does not have a binary alpha channel. (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For x = 0 To dibWidth - 1
                QuickX = x * 4
            For y = 0 To dibHeight - 1
                chkAlpha = iData(QuickX + 3, y)
                
                'For optimization reasons, this is stated as two IFs instead of an OR. I'm not sure if VB
                ' short-circuits OR statements, but I know nested IF statements work. :)
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
                Dim TmpDC As Long
                TmpDC = CreateCompatibleDC(0)
                
                'If successful, select the object into that DC
                If TmpDC <> 0 Then
                
                    'Temporary holder for the object selection
                    Dim oldBitmap As Long
                    oldBitmap = SelectObject(TmpDC, srcPicture.Handle)
                    
                    'Use BitBlt to copy the pixel data to this DIB
                    BitBlt dibDC, 0, 0, dibWidth, dibHeight, TmpDC, 0, 0, vbSrcCopy
                    
                    'Now that we have the pixel data, erase all temporary objects
                    SelectObject TmpDC, oldBitmap
                    DeleteDC TmpDC
                    
                    'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                    ' black background. We typically want the background to be white, so perform a composite if requested.
                    If forceWhiteBackground And (dibColorDepth = 32) Then compositeBackgroundColor
                    
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
    If colorDepth <> 32 And colorDepth <> 24 Then
        colorDepth = 24
    End If
        
    'The back color may or may not be a system color, so translate it just in case
    BackColor = TranslateColor(BackColor)
        
    'Remember this color depth, width, and height
    dibColorDepth = colorDepth
    dibWidth = iWidth
    dibHeight = iHeight
    
    'Prepare the required header
    With dibHeader
        .Size = Len(dibHeader)
        .Planes = 1
        .BitCount = colorDepth
        .Width = iWidth
        .Height = -iHeight
        'As always, this value needs to be a multiple of four; with 32bpp that's automatic, with 24bpp it is not
        If colorDepth = 32 Then
            dibArrayWidth = 4 * iWidth
        Else
            dibArrayWidth = (iWidth * 3 + 3) And &HFFFFFFFC
        End If
        .ImageSize = dibArrayWidth * iHeight
    End With
    
    'Create a new DC for use with this DIB
    dibDC = CreateCompatibleDC(0)
    
    If dibDC <> 0 Then
    
        'Activate color management for this DC.
        ' NOTE!  Due to rejiggering of PD's color management engine, this is no longer necessary.  Color management is only applied
        '         at image load-time (to convert into sRGB), and when drawing to the screen - but when drawing to the screen, all
        '         management is handled at the destination.  For performance reasons, we no longer color-manage DIBs themselves,
        '         which allows us to move RGB/A data between them much more efficiently.
        'turnOnColorManagementForDC dibDC
        
        'Create a DIB
        dibHandle = CreateDIBSection(dibDC, dibHeader, 0, dibBits, 0, 0)
        
        'If successful, select the newly created dib into our DC
        If (dibHandle <> 0) Then
        
            'We will later use dibHandleOriginal to clear up the memory associated with this DIB
            dibHandleOriginal = SelectObject(dibDC, dibHandle)
        
            'If the DIB is 24bpp, apply the background color now
            If (colorDepth = 24) Then
                
                Dim dibRect As RECT
                    SetRect dibRect, 0, 0, iWidth, iHeight
                
                    Dim hBrush As Long
                    hBrush = CreateSolidBrush(BackColor)
                    FillRect dibDC, dibRect, hBrush
                    DeleteObject hBrush
                
            Else
            
                'Finally, set the backColor and/or transparency
                GDIPlusFillDIBRect Me, 0, 0, iWidth, iHeight, BackColor, dibAlpha
                
            End If
            
                
                
                'If colorDepth = 24 Then
                '    Dim dibRect As RECT
                '    SetRect dibRect, 0, 0, iWidth, iHeight
                '
                '    Dim hBrush As Long
                '    hBrush = CreateSolidBrush(BackColor)
                '    FillRect dibDC, dibRect, hBrush
                '    DeleteObject hBrush
                'End If
                
            'If the DIB is 32bpp, apply any transparency now
            'Else
                '(NOT IMPLEMENTED YET; may add in the future if it proves useful)
            'End If
        
        'If DIB creation failed, clear out the work we've done so far
        Else
            eraseDIB
        End If
        
    End If
    
    'Return success contingent on whether we have a DIB pointer or not
    createBlank = (dibHandle <> 0)
    
End Function

'This will effectively reset everything related to this DIB, including all image data. Use cautiously!
Public Function eraseDIB()

    'If we have image data, clear it out
    If dibDC <> 0 Then
        If dibHandle <> 0 Then
            SelectObject dibDC, dibHandleOriginal
            DeleteObject dibHandle
        End If
        DeleteDC dibDC
    End If

    'Reset all associated DIB section variables
    dibDC = 0
    dibHandle = 0
    dibHandleOriginal = 0
    dibBits = 0

    'Reset DIB size
    dibWidth = 0
    dibHeight = 0

End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Reset all associated DIB section variables
    dibDC = 0
    dibHandle = 0
    dibHandleOriginal = 0
    dibBits = 0
    
    'Reset DIB size
    dibWidth = 0
    dibHeight = 0
    
    'Initialize the ICC profile manager
    Set ICCProfile = New pdICCProfile
    
End Sub

'TERMINATE class
Private Sub Class_Terminate()
    eraseDIB
    
    'Release the ICC profile manager
    Set ICCProfile = Nothing
    
End Sub

'Draw this DIB to a picture box. The image will be automatically resized and centered.
Public Sub renderToPictureBox(ByRef dstPicture As PictureBox)

    'Erase any existing picture
    If dstPicture.AutoRedraw Then dstPicture.Picture = LoadPicture("")
    
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = dibWidth
    srcHeight = dibHeight
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = srcWidth / srcHeight
    dstAspect = dstWidth / dstHeight
        
    Dim dWidth As Long, dHeight As Long, previewX As Long, previewY As Long
    
    If srcAspect > dstAspect Then
        dWidth = dstWidth
        dHeight = CDbl(srcHeight / srcWidth) * dWidth
        previewY = CLng((dstHeight - dHeight) / 2)
        previewX = 0
        
    Else
        dHeight = dstHeight
        dWidth = (srcAspect * dHeight)
        previewX = CLng((dstWidth - dWidth) / 2)
        previewY = 0
        
    End If
    
    'Create a temporary DIB to house the stretched image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank dWidth, dHeight
    
    'If the target area is smaller than the DIB, request halftoning
    If (dstPicture.ScaleWidth < dibWidth) Or (dstPicture.ScaleHeight < dibHeight) Then
        SetStretchBltMode tmpDIB.getDIBDC, STRETCHBLT_HALFTONE
    Else
        SetStretchBltMode tmpDIB.getDIBDC, STRETCHBLT_COLORONCOLOR
    End If
    
    'For 24bpp images, simply copy over the current image data
    If dibColorDepth = 24 Then
        StretchBlt tmpDIB.getDIBDC, 0, 0, dWidth, dHeight, dibDC, 0, 0, dibWidth, dibHeight, vbSrcCopy
    
    'For 32bpp images, we must first apply a background checkerboard pattern, then alphablend the shrunken
    ' image onto it.  (Note that GDI+ is used to perform the resize, as StretchBlt can't handle RGBA data properly.)
    Else
    
        Drawing.fillDIBWithAlphaCheckerboard tmpDIB, 0, 0, dWidth, dHeight
        
        Dim gdipResizeDIB As pdDIB
        Set gdipResizeDIB = New pdDIB
        gdipResizeDIB.createBlank dWidth, dHeight, 32
        GDIPlusResizeDIB gdipResizeDIB, 0, 0, dWidth, dHeight, Me, 0, 0, dibWidth, dibHeight, InterpolationModeHighQualityBicubic
        gdipResizeDIB.alphaBlendToDC tmpDIB.getDIBDC
        
    End If
    
    'Activate color management for the destination picture box
    assignDefaultColorProfileToObject dstPicture.hWnd, dstPicture.hDC
    turnOnColorManagementForDC dstPicture.hDC
    
    'Render the image to the picture box
    BitBlt dstPicture.hDC, previewX, previewY, dWidth, dHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
    
    If dstPicture.AutoRedraw Then
        dstPicture.Picture = dstPicture.Image
        dstPicture.Refresh
    End If
    
    Set tmpDIB = Nothing

End Sub

'Load a DIB's header and image information from file.  If the file is compressed, it will automatically be decompressed.
' Optionally, if the DIB data is embedded inside another file, that handle can be passed to the this function, and it will attempt to
' load DIB data from the current file pointer location.  (In this case, just supply a blank filename.)
Public Function createFromFile(ByRef srcFilename As String, Optional ByVal useEmbeddedLocation As Boolean = False, Optional ByRef fileNum As Integer = 0) As Boolean
    
    'If this is not an embedded file, create and open it now
    If Not useEmbeddedLocation Then
        fileNum = FreeFile
        Open srcFilename For Binary As #fileNum
    End If
    
    'Check to make sure this file actually contains DIB data
    Dim DIBIDCheck As String * 4
    Get #fileNum, , DIBIDCheck
    If (DIBIDCheck <> DIB_IDENTIFIER) Then
        Close #fileNum
        Message "Failed to load embedded image; invalid header found."
        createFromFile = False
        Exit Function
    End If

    'Now check to make sure that the version number is supported (not implemented right now, because there's only one version)
    Dim DIBVersionCheck As Long
    Get #fileNum, , DIBVersionCheck
                    
    'Get color depth
    Dim fColorDepth As Long
    Get #fileNum, , fColorDepth
    
    'Get size
    Dim fWidth As Long, fHeight As Long, fArrayWidth As Long
    Get #fileNum, , fWidth
    Get #fileNum, , fHeight
    Get #fileNum, , fArrayWidth
    
    'Is the DIB data compressed?
    Dim toDecompress As Boolean
    Get #fileNum, , toDecompress
    
    Dim fileData() As Byte
    
    'If the data is compressed, we need to retrieve two additional pieces of information: the original and compressed size
    Dim origDataSize As Long, compressedDataSize As Long
    
    If toDecompress Then
    
        'If zLib is not enabled for some reason, abandon the load
        If Not g_ZLibEnabled Then
            Message "DIB data must be decompressed, but zLib is disabled or missing.  Please install the zLib plugin to continue."
            createFromFile = False
            Exit Function
        End If
        
        Get #fileNum, , origDataSize
        Get #fileNum, , compressedDataSize
        
        'Pull the raw pixel data straight from the file
        ReDim fileData(0 To compressedDataSize) As Byte
        Get #fileNum, , fileData
        
    Else
        
        'Pull the raw pixel data straight from the file
        ReDim fileData(0 To fArrayWidth - 1, 0 To fHeight - 1) As Byte
        Get #fileNum, , fileData
        
    End If
    
    'If we've made it this far, attempt to load pixel data. Start by erasing and re-initializing this object to the proper size.
    If createBlank(fWidth, fHeight, fColorDepth) Then
        
        'Now it's time to overwrite our current DIB with the pixel data we pulled from the file.
        ' Start, as always, with a SafeArray pointed at the DIB data.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepInternalSafeArray tmpSA
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        'If the DIB data is compressed, decompress it now
        If toDecompress Then
        
            Dim uncompressedImageData() As Byte
            
            'If the data is decompressed succesfully, convert it to a 2D array and copy it into this class's DIB memory chunk
            If decompressArray(fileData, uncompressedImageData, origDataSize) Then
            
                Dim iData2D() As Byte
                ReDim iData2D(0 To fArrayWidth - 1, 0 To fHeight - 1) As Byte
                CopyMemory iData2D(0, 0), uncompressedImageData(0), UBound(uncompressedImageData) + 1
                CopyMemory iData(0, 0), iData2D(0, 0), fArrayWidth * fHeight
                
                Erase iData2D
                Erase uncompressedImageData
                Erase fileData
            
            'If decompression failed, abandon further processing
            Else
                
                Message "Decompression was not successful (unspecified zLib plugin error).  DIB processing abandoned."
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
                createFromFile = False
                Exit Function
                
            End If
        
        'If the DIB data is not compressed, copy it straight into the DIB
        Else
        
            'Bulk copy the contents of fileData() over to iData() - which is pointing at this object's DIB, remember
            CopyMemory iData(0, 0), fileData(0, 0), fArrayWidth * fHeight
        
        End If
        
        'With our work complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        
        createFromFile = True
    Else
        createFromFile = False
    End If
    
    'If this is an embedded file, leave the file open for further processing.  Otherwise, close it.
    If Not useEmbeddedLocation Then Close #fileNum
    
End Function

'Write this class's DIB information to file.  Compression can optionally be applied to the DIB data.
' Also, if the calling function wants to embed the DIB's data inside an existing file, specify the embedInFile parameter and pass a valid
' file handle, with the pointer already moved to the location desired for embedding.
Public Function writeToFile(ByRef dstFilename As String, Optional ByVal toCompress As Boolean = True, Optional ByVal embedInFile As Boolean = False, Optional ByRef fileNum As Integer) As Boolean
    
    'If we are not embedding the data inside an existing file, create a new file now.
    If Not embedInFile Then
        
        'Delete any existing file (overwrite)
        If FileExist(dstFilename) Then Kill dstFilename
        
        'Open the file
        fileNum = FreeFile
        Open dstFilename For Binary As #fileNum
        
    End If
    
    'Identifiers
    Put #fileNum, , DIB_IDENTIFIER
    Put #fileNum, , DIB_FILE_VERSION_2012
    
    'Color depth
    Put #fileNum, , dibColorDepth
    
    'Size
    Put #fileNum, , dibWidth
    Put #fileNum, , dibHeight
    Put #fileNum, , dibArrayWidth
        
    'And finally, the pixel data.  Start by pointing a VB array at the DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    If toCompress And Not g_ZLibEnabled Then Message "Compression requested, but zLib plugin is disabled or missing.  Data will not be compressed."
    
    'If compression has been requested, and zLib is available, apply compression now.
    If toCompress And g_ZLibEnabled Then
        
        'Convert the DIB to a 1D array
        Dim src1DImage() As Byte
        Dim aSize As Long
        aSize = (UBound(iData, 1) + 1) * (UBound(iData, 2) + 1) - 1
        ReDim src1DImage(0 To aSize) As Byte
        CopyMemory src1DImage(0), iData(0, 0), aSize
        
        'Compress that 1D array
        Dim origSize As Long, compressSize As Long
        Dim compressedImageData() As Byte
        
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        
        'If compression was successful, write out the compressed data to file
        If compressArray(src1DImage, compressedImageData, origSize, compressSize) Then
        
            'Mark the data as compressed
            Put #fileNum, , toCompress
            
            'Write the original size of the array, followed by the compressed size, followed by the image data itself.  (The array sizes will
            ' help us automatically create intelligently-sized buffers at decompression time.)
            Put #fileNum, , origSize
            Put #fileNum, , compressSize
            Put #fileNum, , compressedImageData
        
        'If compression failed, write uncompressed data instead.
        Else
            Message "Compression failed (unspecified zLib plugin error).  Uncompressed data written instead."
            toCompress = False
            Put #fileNum, , toCompress
            Put #fileNum, , iData
        End If
        
        Erase compressedImageData
        
    'If compression has not been requested, or zLib is unavailable, dump the raw image data to file.
    Else
        
        'Now dump the entire array into the file. To my knowledge, this is the fastest way to do this in VB.
        toCompress = False
        Put #fileNum, , toCompress
        Put #fileNum, , iData
        
    End If
    
    'Release the VB array that points at the DIB
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    
    'If this is not an embedding, close the file now
    If Not embedInFile Then Close #fileNum

End Function

'Write this class's DIB information to a valid BMP file.
Public Sub writeToBitmapFile(ByRef dstFilename As String)

    'First, make sure this DIB actually contains image data
    If Not Me.hasImage() Then Exit Sub
        
    'Calculate the size of a scanline
    Dim slWidth As Long
    slWidth = ((dibWidth * dibColorDepth + 31) \ 32) * 4
    
    'A valid bitmap header consists of two parts: a file header, and an image header.
    Dim bmpFileHeader As BITMAPFILEHEADER
    Dim bmpInfoHeader As BITMAPINFO
    
    'Build the file header first
    With bmpFileHeader
        .Type = &H4D42 'BMP identifier
        .Size = Len(bmpFileHeader) + (slWidth * dibHeight) 'Length of the file
        .OffBits = Len(bmpFileHeader) + Len(bmpInfoHeader) 'Length of the header area
    End With
    
    '...then the image header
    With bmpInfoHeader.Header
        .Size = 40
        .Planes = 1
        .BitCount = dibColorDepth
        .Width = dibWidth
        .Height = dibHeight
    End With
    
    'Finally, the image bytes
    Dim iData() As Byte
    ReDim iData(0 To slWidth - 1, 0 To dibHeight - 1) As Byte
    GetDIBits dibDC, dibHandle, 0, dibHeight, iData(0, 0), bmpInfoHeader, 0
    
    'With all our information in place, check to see if that file already exists. Delete it if necessary.
    If FileExist(dstFilename) Then Kill dstFilename
        
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

'As a helper to external functions, this sub can be used to fill a destination array with a copy the image's current data.
' Note that this is raw image bytes - no header whatsoever - so it's assumed the caller knows how to deal with the data.
Public Sub copyImageBytesIntoStream(ByRef dstArray() As Byte)

    'Retrieve a pointer to the current DIB bits (in VB format).  Technically we could just use the dibBits pointer
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

'The counterpart to copyImageBytesIntoStream, above.  This function will blindly copy an array of bytes over the current
' image's data.  For it to work, the DIB must have been correctly created at the right size and color depth BEFORE CALLING
' THIS FUNCTION.  Otherwise, you will get errors.
Public Sub copyStreamOverImageArray(ByRef srcArray() As Byte)
    CopyMemory ByVal dibBits, srcArray(0), UBound(srcArray) + 1
End Sub

'Sometimes this class needs to access its own DIB bits. Here's how.
Private Sub prepInternalSafeArray(ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = dibArrayWidth
        .pvData = dibBits
    End With
End Sub

'Sometimes this class needs to access the DIB bits of other DIBs. Here's how.
Private Sub prepExternalSafeArray(ByRef srcDIB As pdDIB, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
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
Public Sub compositeBackgroundColor(Optional ByVal newR As Byte = 255, Optional ByVal newG As Byte = 255, Optional ByVal newB As Byte = 255, Optional ByVal parentIsPremultiplied As Boolean = False)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If dibColorDepth <> 32 Then Exit Sub

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
    
    'Because our alpha values are pre-multiplied, we can composite them against the background color via use of a look-up table.
    Dim rLookup(0 To 255) As Byte, gLookup(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    'Populate a unique lookup table for each color, based on each possible alpha value (0 to 255)
    For x = 0 To 255
        
        tmpAlpha = 1 - (x / 255)
        rLookup(x) = newR * tmpAlpha
        gLookup(x) = newG * tmpAlpha
        bLookup(x) = newB * tmpAlpha
    
    Next x
    
    'Loop through the image, blending colors as we go
    For x = 0 To dibWidth - 1
        QuickX = x * 4
    For y = 0 To dibHeight - 1
        
        'Access the alpha data for this pixel
        checkAlpha = iData(QuickX + 3, y)
        
        'Ignore fully opaque pixels. (This makes the routine much faster.)
        If checkAlpha <> 255 Then
            
            'Use that alpha value to blend the current colors with the newly requested color
            If parentIsPremultiplied Then
            
                iData(QuickX + 2, y) = iData(QuickX + 2, y) + rLookup(checkAlpha)
                iData(QuickX + 1, y) = iData(QuickX + 1, y) + gLookup(checkAlpha)
                iData(QuickX, y) = iData(QuickX, y) + bLookup(checkAlpha)
                
            Else
            
                'Convert the alpha value to a floating-point variable
                tmpAlpha = checkAlpha / 255
            
                'Use that alpha value to blend the current colors with the newly requested color
                iData(QuickX + 2, y) = Blend2Colors(iData(QuickX + 2, y), newR, tmpAlpha)
                iData(QuickX + 1, y) = Blend2Colors(iData(QuickX + 1, y), newG, tmpAlpha)
                iData(QuickX, y) = Blend2Colors(iData(QuickX, y), newB, tmpAlpha)
                
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

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function Blend2Colors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByRef mixRatio As Double) As Byte
    Blend2Colors = (mixRatio * Color1) + ((1 - mixRatio) * Color2)
    'Alternate blend formula from Robert Rayment - my own testing shows no measurable speed difference between the two methods
    'Blend2Colors = mixRatio * (CInt(Color1) - Color2) + Color2
End Function

'Check to see if a 32bpp image is really 32bpp. (Basically, scan all pixels in the alpha channel. If all values are set to
' 255 or all values are set to 0, remove the alpha channel and rebuild this DIB in 24bpp mode.)
Public Function verifyAlphaChannel() As Boolean

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If dibColorDepth <> 32 Then
        verifyAlphaChannel = True
        Exit Function
    End If
    
    'This routine will fail if the width or height of this DIB is 0
    If dibWidth = 0 Or dibHeight = 0 Then
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
    
    'Determine the alpha value of the top-left pixel. This will be used as our baseline value.
    initAlpha = iData(3, 0)
    
    'If initAlpha is something other than 255 or 0, we don't need to check the image
    If (initAlpha <> 0) And (initAlpha <> 255) Then
        verifyAlphaChannel = True
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        Exit Function
    End If
        
    'Loop through the image, comparing colors as we go
    For x = 0 To dibWidth - 1
        QuickX = x * 4
    For y = 0 To dibHeight - 1
        
        'Compare the alpha data for this pixel to the initial pixel. If they DO NOT match, this is a valid alpha channel.
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

    'Return checkAlpha. If varying alpha values were found, this function returns TRUE. If all values were the same,
    ' this function returns FALSE.
    verifyAlphaChannel = checkAlpha

End Function

'Copy the alpha values from another pdDIB object to this one. This is useful when the alpha channel for this DIB must be lost
' during a transformation (typically something involving FreeImage or GDI+), and a temporary DIB was made to preserve the alpha
' data. If the image is currently 24bpp, this function will first convert it to 32bpp before copying the alpha data.

'IMPORTANT NOTE: to keep this function fast, NO ARRAY BOUNDS CHECKING IS DONE. Make sure that the passed pdDIB object is
' THE SAME SIZE (or larger) than this DIB, or you will experience critical errors.
Public Sub copyAlphaFromExistingDIB(ByRef srcDIB As pdDIB)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If dibColorDepth <> 32 Then convertTo32bpp

    'Prepare a SafeArray that points to our own DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Now make a second array that points to the source DIB data
    Dim aData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepExternalSafeArray srcDIB, srcSA
    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    'Loop through the image, copying alpha values as we go
    For x = 0 To dibWidth - 1
        QuickX = x * 4
    For y = 0 To dibHeight - 1
        
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
Public Sub fixPremultipliedAlpha(Optional ByVal applyPremultiplication As Boolean = False)

    'This function doesn't matter if the image isn't 32bpp
    If dibColorDepth <> 32 Then Exit Sub

    'Prepare a SAFEARRAY, which allows us to access the DIB data directly
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tmpAlpha As Byte, tmpAlphaModifier As Double

    'Loop through the image, removing the effects of pre-multiplication as we go
    For x = 0 To dibWidth - 1
        QuickX = x * 4
    For y = 0 To dibHeight - 1
        
        tmpAlpha = iData(QuickX + 3, y)
        
        
        If applyPremultiplication Then
        
            'When applying premultiplication, we can ignore fully opaque pixels
            If tmpAlpha <> 255 Then
            
                'We can shortcut the calculation of full transparent pixels (they are made black)
                If tmpAlpha = 0 Then
                    iData(QuickX + 2, y) = 0
                    iData(QuickX + 1, y) = 0
                    iData(QuickX, y) = 0
                Else
            
                    r = iData(QuickX + 2, y)
                    g = iData(QuickX + 1, y)
                    b = iData(QuickX, y)
                    
                    tmpAlphaModifier = tmpAlpha / 255
                    
                    'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                    r = r * tmpAlphaModifier
                    g = g * tmpAlphaModifier
                    b = b * tmpAlphaModifier
                    
                    iData(QuickX + 2, y) = CByte(r)
                    iData(QuickX + 1, y) = CByte(g)
                    iData(QuickX, y) = CByte(b)
                    
                End If
            
            End If
        
        Else
            
            'When removing premultiplication, we can ignore fully opaque and fully transparent pixels
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
                
                If r > 255 Then r = 255
                If g > 255 Then g = 255
                If b > 255 Then b = 255
                                
                iData(QuickX + 2, y) = CByte(r)
                iData(QuickX + 1, y) = CByte(g)
                iData(QuickX, y) = CByte(b)
            
            End If
            End If
        
        End If
        
    Next y
    Next x
    
    'With our alpha channel complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub alphaBlendToDC(ByVal dstDC As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    Dim bfParams As Long
    
    If dibColorDepth = 32 Then
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
    
    AlphaBlend dstDC, dstX, dstY, IIf(newWidth = 0, dibWidth, newWidth), IIf(newHeight = 0, dibHeight, newHeight), dibDC, 0, 0, dibWidth, dibHeight, bfParams
    
End Sub
