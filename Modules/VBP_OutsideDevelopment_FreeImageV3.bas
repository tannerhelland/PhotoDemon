Attribute VB_Name = "Outside_FreeImageV3"
'Note: this file has been heavily modified for use within PhotoDemon.  The vast majority of the code is copied directly from the official
' VB6 wrapper by Carsten Klein, but I have stripped out unused functions, retyped certain enums (to work more nicely with PD's custom
' systems), and directly modified a few functions to handle data more easily for PD's purposes.

'So basically: IF YOU WANT TO USE THIS CODE IN YOUR OWN PROJECT, PLEASE DOWNLOAD AN ORIGINAL COPY FROM THIS LINK:
'http://freeimage.sourceforge.net/download.html

'Many thanks to Carsten Klein and the FreeImage team for this excellent library (and associated DLL).


'// ==========================================================
'// Visual Basic Wrapper for FreeImage 3
'// Original FreeImage 3 functions and VB compatible derived functions
'// Design and implementation by
'// - Carsten Klein (cklein05@users.sourceforge.net)
'//
'// Main reference : Curland, Matthew., Advanced Visual Basic 6, Addison Wesley, ISBN 0201707128, (c) 2000
'//                  Steve McMahon, creator of the excellent site vbAccelerator at http://www.vbaccelerator.com/
'//                  MSDN Knowledge Base
'//
'// This file is part of FreeImage 3
'//
'// COVERED CODE IS PROVIDED UNDER THIS LICENSE ON AN "AS IS" BASIS, WITHOUT WARRANTY
'// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, WITHOUT LIMITATION, WARRANTIES
'// THAT THE COVERED CODE IS FREE OF DEFECTS, MERCHANTABLE, FIT FOR A PARTICULAR PURPOSE
'// OR NON-INFRINGING. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE COVERED
'// CODE IS WITH YOU. SHOULD ANY COVERED CODE PROVE DEFECTIVE IN ANY RESPECT, YOU (NOT
'// THE INITIAL DEVELOPER OR ANY OTHER CONTRIBUTOR) ASSUME THE COST OF ANY NECESSARY
'// SERVICING, REPAIR OR CORRECTION. THIS DISCLAIMER OF WARRANTY CONSTITUTES AN ESSENTIAL
'// PART OF THIS LICENSE. NO USE OF ANY COVERED CODE IS AUTHORIZED HEREUNDER EXCEPT UNDER
'// THIS DISCLAIMER.
'//
'// Use at your own risk!
'// ==========================================================

'// ==========================================================
'// CVS
'// $Revision: 2.20 $
'// $Date: 2013/10/10 11:11:15 $
'// $Id: MFreeImage.bas,v 2.20 2013/10/10 11:11:15 cklein05 Exp $
'// ==========================================================


Option Explicit

'--------------------------------------------------------------------------------
' General notes on implementation and design
'--------------------------------------------------------------------------------

' General:

' Most of the pointer type parameters used in the FreeImage API are actually
' declared as Long in VB. That is also true for return values. 'Out' parameters
' are declared ByRef, so they can receive the provided address of the pointer.
' 'In' parameters are declared ByVal since in VB the Long variable is not a
' pointer type but contains the address of the pointer.


' Functions returning a special type:

' Some of the following external function declarations of the FreeImage 3 functions
' are declared Private. Additionally the token 'Int' is appended to the VB function
' name, what means 'Internal' to avoid naming confusion. All of these return a value
' of a certain type that can't be used with a declared function in VB directly but
' would need the function to be declared in a type library. Since this wrapper module
' should not depend on a compile time type library, these functions require some extra
' work to be done and also a VB wrapper function to make them look like the C/C++
' function.


' Functions returning Strings:

' Some of the declared FreeImage functions are defined as 'const char *' in C/C++
' and so actually return a string pointer. Without using a type library for declaring
' these functions, in VB it is impossible to declare these functions to return a
' VB String type. So each of these functions is wrapped by a VB implemented function
' named correctly according to the FreeImage API, actually returning a 'real' VB String.


' Functions returning Booleans:

' A Boolean is a numeric 32 bit value in both C/C++ and VB. In C/C++ TRUE is defined
' as 1 whereas in VB True is -1 (all bits set). When a function is declared as 'Boolean'
' in VB, the return value (all 32 bits) of the called function is just used "as is" and
' maybe assigned to a VB boolean variable. A Boolean in VB is 'False' when the numeric
' value is NULL (0) and 'True' in any other case. So, at a first glance, everything
' would be great since both numeric values -1 (VB True) and 1 (C/C++ TRUE) are actually
' 'True' in VB.
' But, if you have a VB variable (or a function returning a Boolean) with just some bits
' set and use the VB 'Not' operator, the result is not what you would expect. In this
' case, if bTest is True, (Not bTest) is also True. The 'Not' operator just toggles all
' bits by XOR-ing the value with -1. So, the result is not so surprisingly any more:
' The C/C++ TRUE value is 0...0001. When all bits are XORed with 1, the result is
' 1...1110 what is also not NULL (0) so this is still 'True' in VB.
' The resolution is to convert these return values into real VB Booleans in a wrapper
' function, one for each declared FreeImage function. Therefore each C/C++ BOOL
' function is declared Private as xxxInt(...). A Public Boolean wrapper function
' xxx(...) returns a real Boolean with 'xxx = (xxxInt(...) = 1)'.


' Extended and derived functions:

' Some of the functions are additionally provided in an extended, call it a more VB
' friendly version, named '...Ex'. For example look at the 'FreeImage_GetPaletteEx'
' function. Most of them are dealing with arrays and so actually return a VB style
' array of correct type.

' The wrapper also includes some derived functions that should make life easier for
' not only a VB programmer.

' Better VB interoperability is given by offering conversion between DIBs and
' VB Picture objects. See the FreeImage_CreateFromOlePicture and
' FreeImage_GetOlePicture functions.

' Both known VB functions LoadPicture() and SavePicture() are provided in extended
' versions calles LoadPictureEx() and SavePictureEx() offering the FreeImage 3´s
' image file types.

' The FreeImage 3 error handling is provided in VB after calling the VB specific
' function FreeImage_InitErrorHandler()


' Enumerations:

' All of the enumaration members are additionally 'declared' as constants in a
' conditional compiler directive '#If...#Then' block that is actually unreachable.
' For example see:
'
' Public Enum FREE_IMAGE_QUANTIZE
'    FIQ_WUQUANT = 0           ' Xiaolin Wu color quantization algorithm
'    FIQ_NNQUANT = 1           ' NeuQuant neural-net quantization algorithm by Anthony Dekker
' End Enum
' #If False Then
'    Const FIQ_WUQUANT = 0
'    Const FIQ_NNQUANT = 1
' #End If
'
' Since this module is supposed to be used directly in VB projects rather than in
' compiled form (mybe through an ActiveX-DLL), this is for tweaking some ugly VB
' behaviour regarding enumerations. Enum members are automatically adjusted in case
' by the VB IDE whenever you type these members in wrong case. Since these are also
' constants now, they are no longer adjusted to wrong case but always corrected
' according to the definition of the constant. As the expression '#If False Then'
' actually never comes true, these constants are not really defined either when running
' in the VB IDE nor in compiled form.


'--------------------------------------------------------------------------------
' ToDo and known issues (unordered and with no priority)
'--------------------------------------------------------------------------------

' ToDo: more inline documentation for mask image creation and icon functions
'       needed

'--------------------------------------------------------------------------------
' Change Log
'--------------------------------------------------------------------------------

'* : fixed
'- : removed
'! : changed
'+ : added
'
'October 1, 2012 - 2.17
'- [Carsten Klein] removed temporary workaround for 16-bit standard type bitmaps introduced in version 2.15, which temporarily stored RGB masks directly after the BITMAPINFO structure, when creating a HBITMAP.
'* [Carsten Klein] fixed a potential overflow bug in both pNormalizeRational and pNormalizeSRational: these now do nothing if any of numerator and denominator is either 1 or 0 (zero).
'+ [Carsten Klein] added load flag JPEG_GREYSCALE as well as the enum constant FILO_JPEG_GREYSCALE.
'! [Carsten Klein] changed constant FREEIMAGE_RELEASE_SERIAL to 4 to match current version 3.15.4
'
'! now FreeImage version 3.15.4
'
' NOTE FROM TANNER: a very detailed changelog follows this line in the original, but it has been removed for brevity's sake

'--------------------------------------------------------------------------------
' Win32 API function, struct and constant declarations
'--------------------------------------------------------------------------------

Private Const ERROR_SUCCESS As Long = 0

'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

'OLEAUT32
Public Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
    ByRef lpPictDesc As PictDesc, _
    ByRef riid As Guid, _
    ByVal fOwn As Long, _
    ByRef lplpvObj As IPicture) As Long
    
Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32.dll" ( _
    ByVal cDims As Long, _
    ByRef ppsaOut As Long) As Long
    
Private Declare Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" ( _
    ByVal psa As Long) As Long
    
Private Declare Sub SafeArrayDestroyData Lib "oleaut32.dll" ( _
    ByVal psa As Long)
    
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
    ByVal clr As OLE_COLOR, _
    ByVal hPal As Long, _
    ByRef lpcolorref As Long) As Long
    
'Private Const CLR_INVALID As Long = &HFFFF&
    

'SAFEARRAY
Private Const FADF_AUTO As Long = (&H1)
Private Const FADF_FIXEDSIZE As Long = (&H10)

Private Type SAVEARRAY1D
   cDims As Integer
   fFeatures As Integer
   cbElements As Long
   cLocks As Long
   pvData As Long
   cElements As Long
   lLbound As Long
End Type


'MSVBVM60
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef Ptr() As Any) As Long


'USER32
Private Declare Function ReleaseDC Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal hDC As Long) As Long

Private Declare Function GetDC Lib "user32.dll" ( _
    ByVal hWnd As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function DestroyIcon Lib "user32.dll" ( _
    ByVal hIcon As Long) As Long

Private Declare Function CreateIconIndirect Lib "user32.dll" ( _
    ByRef piconinfo As ICONINFO) As Long

Private Type PictDesc
   cbSizeofStruct As Long
   picType As Long
   hImage As Long
   xExt As Long
   yExt As Long
End Type

Private Type BITMAP_API
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Type ICONINFO
   fIcon As Long
   xHotspot As Long
   yHotspot As Long
   hbmMask As Long
   hbmColor As Long
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
    
'GDI32

Private Declare Function GetStretchBltMode Lib "gdi32.dll" ( _
    ByVal hDC As Long) As Long

Private Declare Function SetStretchBltMode Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal nStretchMode As Long) As Long
    
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal dx As Long, _
    ByVal dy As Long, _
    ByVal srcX As Long, _
    ByVal srcY As Long, _
    ByVal Scan As Long, _
    ByVal NumScans As Long, _
    ByVal Bits As Long, _
    ByVal BitsInfo As Long, _
    ByVal wUsage As Long) As Long
    
Private Declare Function StretchDIBits Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal dx As Long, _
    ByVal dy As Long, _
    ByVal srcX As Long, _
    ByVal srcY As Long, _
    ByVal wSrcWidth As Long, _
    ByVal wSrcHeight As Long, _
    ByVal lpBits As Long, _
    ByVal lpBitsInfo As Long, _
    ByVal wUsage As Long, _
    ByVal dwRop As Long) As Long
    
Private Declare Function CreateDIBitmap Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal lpInfoHeader As Long, _
    ByVal dwUsage As Long, _
    ByVal lpInitBits As Long, _
    ByVal lpInitInfo As Long, _
    ByVal wUsage As Long) As Long
    
Private Declare Function CreateDIBSection Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal pbmi As Long, _
    ByVal iUsage As Long, _
    ByRef ppvBits As Long, _
    ByVal hSection As Long, _
    ByVal dwOffset As Long) As Long

Private Const CBM_INIT As Long = &H4
    
Private Declare Function GetDIBits Lib "gdi32.dll" ( _
    ByVal aHDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByVal lpBits As Long, _
    ByVal lpBI As Long, _
    ByVal wUsage As Long) As Long
    
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    ByRef lpObject As Any) As Long
    
Private Declare Function SelectObject Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function GetCurrentObject Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal uObjectType As Long) As Long

Private Const OBJ_BITMAP As Long = 7
    
Private Const COLORONCOLOR As Long = 3

'MSIMG32
Private Declare Function AlphaBlend Lib "MSIMG32.dll" ( _
    ByVal hdcDest As Long, _
    ByVal nXOriginDest As Long, _
    ByVal nYOriginDest As Long, _
    ByVal nWidthDest As Long, _
    ByVal nHeightDest As Long, _
    ByVal hdcSrc As Long, _
    ByVal nXOriginSrc As Long, _
    ByVal nYOriginSrc As Long, _
    ByVal nWidthSrc As Long, _
    ByVal nHeightSrc As Long, _
    ByVal lBlendFunction As Long) As Long

Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1

Private Const BLACKONWHITE As Long = 1
Private Const WHITEONBLACK As Long = 2
'Private Const COLORONCOLOR As Long = 3

Public Enum STRETCH_MODE
   SM_BLACKONWHITE = 1
   SM_WHITEONBLACK = 2
   SM_COLORONCOLOR = 3
End Enum
#If False Then
   Const SM_BLACKONWHITE = 1
   Const SM_WHITEONBLACK = 2
   Const SM_COLORONCOLOR = 3
#End If


Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCERASE As Long = &H440328
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086

Public Enum RASTER_OPERATOR
   ROP_SRCAND = SRCAND
   ROP_SRCCOPY = SRCCOPY
   ROP_SRCERASE = SRCERASE
   ROP_SRCINVERT = SRCINVERT
   ROP_SRCPAINT = SRCPAINT
End Enum
#If False Then
   Const ROP_SRCAND = SRCAND
   Const ROP_SRCCOPY = SRCCOPY
   Const ROP_SRCERASE = SRCERASE
   Const ROP_SRCINVERT = SRCINVERT
   Const ROP_SRCPAINT = SRCPAINT
#End If

Private Const DIB_RGB_COLORS As Long = 0

Public Enum DRAW_MODE
   DM_DRAW_DEFAULT = &H0
   DM_MIRROR_NONE = DM_DRAW_DEFAULT
   DM_MIRROR_VERTICAL = &H1
   DM_MIRROR_HORIZONTAL = &H2
   DM_MIRROR_BOTH = DM_MIRROR_VERTICAL Or DM_MIRROR_HORIZONTAL
End Enum
#If False Then
   Const DM_DRAW_DEFAULT = &H0
   Const DM_MIRROR_NONE = DM_DRAW_DEFAULT
   Const DM_MIRROR_VERTICAL = &H1
   Const DM_MIRROR_HORIZONTAL = &H2
   Const DM_MIRROR_BOTH = DM_MIRROR_VERTICAL Or DM_MIRROR_HORIZONTAL
#End If

Public Enum HISTOGRAM_ORIENTATION
   HOR_TOP_DOWN = &H0
   HOR_BOTTOM_UP = &H1
End Enum
#If False Then
   Const HOR_TOP_DOWN = &H0
   Const HOR_BOTTOM_UP = &H1
#End If


'--------------------------------------------------------------------------------
' FreeImage 3 types, constants and enumerations
'--------------------------------------------------------------------------------

'FREEIMAGE

' Version information
Public Const FREEIMAGE_MAJOR_VERSION As Long = 3
Public Const FREEIMAGE_MINOR_VERSION As Long = 16
Public Const FREEIMAGE_RELEASE_SERIAL As Long = 0

' Memory stream pointer operation flags
Public Const SEEK_SET As Long = 0
Public Const SEEK_CUR As Long = 1
Public Const SEEK_END As Long = 2

' Indexes for byte arrays, masks and shifts for treating pixels as words
' These coincide with the order of RGBQUAD and RGBTRIPLE
' Little Endian (x86 / MS Windows, Linux) : BGR(A) order
Public Const FI_RGBA_RED As Long = 2
Public Const FI_RGBA_GREEN As Long = 1
Public Const FI_RGBA_BLUE As Long = 0
Public Const FI_RGBA_ALPHA As Long = 3
Public Const FI_RGBA_RED_MASK As Long = &HFF0000
Public Const FI_RGBA_GREEN_MASK As Long = &HFF00
Public Const FI_RGBA_BLUE_MASK As Long = &HFF
Public Const FI_RGBA_ALPHA_MASK As Long = &HFF000000
Public Const FI_RGBA_RED_SHIFT As Long = 16
Public Const FI_RGBA_GREEN_SHIFT As Long = 8
Public Const FI_RGBA_BLUE_SHIFT As Long = 0
Public Const FI_RGBA_ALPHA_SHIFT As Long = 24

' The 16 bit macros only include masks and shifts, since each color element is not byte aligned
Public Const FI16_555_RED_MASK As Long = &H7C00
Public Const FI16_555_GREEN_MASK As Long = &H3E0
Public Const FI16_555_BLUE_MASK As Long = &H1F
Public Const FI16_555_RED_SHIFT As Long = 10
Public Const FI16_555_GREEN_SHIFT As Long = 5
Public Const FI16_555_BLUE_SHIFT As Long = 0
Public Const FI16_565_RED_MASK As Long = &HF800
Public Const FI16_565_GREEN_MASK As Long = &H7E0
Public Const FI16_565_BLUE_MASK As Long = &H1F
Public Const FI16_565_RED_SHIFT As Long = 11
Public Const FI16_565_GREEN_SHIFT As Long = 5
Public Const FI16_565_BLUE_SHIFT As Long = 0

' ICC profile support
Public Const FIICC_DEFAULT As Long = &H0
Public Const FIICC_COLOR_IS_CMYK As Long = &H1

Private Const FREE_IMAGE_ICC_COLOR_MODEL_MASK As Long = &H1
Public Enum FREE_IMAGE_ICC_COLOR_MODEL
   FIICC_COLOR_MODEL_RGB = &H0
   FIICC_COLOR_MODEL_CMYK = &H1
End Enum

' Load / Save flag constants
Public Const FIF_LOAD_NOPIXELS = &H8000              ' load the image header only (not supported by all plugins)

Public Const BMP_DEFAULT As Long = 0
Public Const BMP_SAVE_RLE As Long = 1
Public Const CUT_DEFAULT As Long = 0
Public Const DDS_DEFAULT As Long = 0
Public Const EXR_DEFAULT As Long = 0                 ' save data as half with piz-based wavelet compression
Public Const EXR_FLOAT As Long = &H1                 ' save data as float instead of as half (not recommended)
Public Const EXR_NONE As Long = &H2                  ' save with no compression
Public Const EXR_ZIP As Long = &H4                   ' save with zlib compression, in blocks of 16 scan lines
Public Const EXR_PIZ As Long = &H8                   ' save with piz-based wavelet compression
Public Const EXR_PXR24 As Long = &H10                ' save with lossy 24-bit float compression
Public Const EXR_B44 As Long = &H20                  ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
Public Const EXR_LC As Long = &H40                   ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
Public Const FAXG3_DEFAULT As Long = 0
Public Const GIF_DEFAULT As Long = 0
Public Const GIF_LOAD256 As Long = 1                 ' Load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
Public Const GIF_PLAYBACK As Long = 2                ''Play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
Public Const HDR_DEFAULT As Long = 0
Public Const ICO_DEFAULT As Long = 0
Public Const ICO_MAKEALPHA As Long = 1               ' convert to 32bpp and create an alpha channel from the AND-mask when loading
Public Const IFF_DEFAULT As Long = 0
Public Const J2K_DEFAULT  As Long = 0                ' save with a 16:1 rate
Public Const JP2_DEFAULT As Long = 0                 ' save with a 16:1 rate
Public Const JPEG_DEFAULT As Long = 0                ' loading (see JPEG_FAST); saving (see JPEG_QUALITYGOOD|JPEG_SUBSAMPLING_420)
Public Const JPEG_FAST As Long = &H1                 ' load the file as fast as possible, sacrificing some quality
Public Const JPEG_ACCURATE As Long = &H2             ' load the file with the best quality, sacrificing some speed
Public Const JPEG_CMYK As Long = &H4                 ' load separated CMYK "as is" (use 'OR' to combine with other flags)
Public Const JPEG_EXIFROTATE As Long = &H8           ' load and rotate according to Exif 'Orientation' tag if available
Public Const JPEG_GREYSCALE As Long = &H10           ' load and convert to a 8-bit greyscale image
Public Const JPEG_QUALITYSUPERB As Long = &H80       ' save with superb quality (100:1)
Public Const JPEG_QUALITYGOOD As Long = &H100        ' save with good quality (75:1)
Public Const JPEG_QUALITYNORMAL As Long = &H200      ' save with normal quality (50:1)
Public Const JPEG_QUALITYAVERAGE As Long = &H400     ' save with average quality (25:1)
Public Const JPEG_QUALITYBAD As Long = &H800         ' save with bad quality (10:1)
Public Const JPEG_PROGRESSIVE As Long = &H2000       ' save as a progressive-JPEG (use 'OR' to combine with other save flags)
Public Const JPEG_SUBSAMPLING_411 As Long = &H1000   ' save with high 4x1 chroma subsampling (4:1:1)
Public Const JPEG_SUBSAMPLING_420 As Long = &H4000   ' save with medium 2x2 medium chroma subsampling (4:2:0) - default value
Public Const JPEG_SUBSAMPLING_422 As Long = &H8000   ' save with low 2x1 chroma subsampling (4:2:2)
Public Const JPEG_SUBSAMPLING_444 As Long = &H10000  ' save with no chroma subsampling (4:4:4)
Public Const JPEG_OPTIMIZE As Long = &H20000         ' on saving, compute optimal Huffman coding tables (can reduce a few percent of file size)
Public Const JPEG_BASELINE As Long = &H40000         ' save basic JPEG, without metadata or any markers
Public Const KOALA_DEFAULT As Long = 0
Public Const LBM_DEFAULT As Long = 0
Public Const MNG_DEFAULT As Long = 0
Public Const PCD_DEFAULT As Long = 0
Public Const PCD_BASE As Long = 1                    ' load the bitmap sized 768 x 512
Public Const PCD_BASEDIV4 As Long = 2                ' load the bitmap sized 384 x 256
Public Const PCD_BASEDIV16 As Long = 3               ' load the bitmap sized 192 x 128
Public Const PCX_DEFAULT As Long = 0
Public Const PFM_DEFAULT As Long = 0
Public Const PICT_DEFAULT As Long = 0
Public Const PNG_DEFAULT As Long = 0
Public Const PNG_IGNOREGAMMA As Long = 1             ' avoid gamma correction
Public Const PNG_Z_BEST_SPEED As Long = &H1          ' save using ZLib level 1 compression flag (default value is 6)
Public Const PNG_Z_DEFAULT_COMPRESSION As Long = &H6 ' save using ZLib level 6 compression flag (default recommended value)
Public Const PNG_Z_BEST_COMPRESSION As Long = &H9    ' save using ZLib level 9 compression flag (default value is 6)
Public Const PNG_Z_NO_COMPRESSION As Long = &H100    ' save without ZLib compression
Public Const PNG_INTERLACED As Long = &H200          ' save using Adam7 interlacing (use | to combine with other save flags)
Public Const PNM_DEFAULT As Long = 0
Public Const PNM_SAVE_RAW As Long = 0                ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
Public Const PNM_SAVE_ASCII As Long = 1              ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
Public Const PSD_DEFAULT As Long = 0
Public Const PSD_CMYK As Long = 1                    ' reads tags for separated CMYK (default is conversion to RGB)
Public Const PSD_LAB As Long = 2                     ' reads tags for CIELab (default is conversion to RGB)
Public Const RAS_DEFAULT As Long = 0
Public Const RAW_DEFAULT As Long = 0                 ' load the file as linear RGB 48-bit
Public Const RAW_PREVIEW As Long = 1                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
Public Const RAW_DISPLAY As Long = 2                 ' load the file as RGB 24-bit
Public Const RAW_HALFSIZE As Long = 4                ' load the file as half-size color image
Public Const SGI_DEFAULT As Long = 0
Public Const TARGA_DEFAULT As Long = 0
Public Const TARGA_LOAD_RGB888 As Long = 1           ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
Public Const TARGA_SAVE_RLE As Long = 2              ' if set, the writer saves with RLE compression
Public Const TIFF_DEFAULT As Long = 0
Public Const TIFF_CMYK As Long = 1                   ' reads/stores tags for separated CMYK (use 'OR' to combine with compression flags)
Public Const TIFF_PACKBITS As Long = &H100           ' save using PACKBITS compression
Public Const TIFF_DEFLATE As Long = &H200            ' save using DEFLATE compression (a.k.a. ZLIB compression)
Public Const TIFF_ADOBE_DEFLATE As Long = &H400      ' save using ADOBE DEFLATE compression
Public Const TIFF_NONE As Long = &H800               ' save without any compression
Public Const TIFF_CCITTFAX3 As Long = &H1000         ' save using CCITT Group 3 fax encoding
Public Const TIFF_CCITTFAX4 As Long = &H2000         ' save using CCITT Group 4 fax encoding
Public Const TIFF_LZW As Long = &H4000               ' save using LZW compression
Public Const TIFF_JPEG As Long = &H8000              ' save using JPEG compression
Public Const TIFF_LOGLUV As Long = &H10000           ' save using LogLuv compression
Public Const WBMP_DEFAULT As Long = 0
Public Const XBM_DEFAULT As Long = 0
Public Const XPM_DEFAULT As Long = 0
Public Const WEBP_DEFAULT As Long = 0                ' save with good quality (75:1)
Public Const WEBP_LOSSLESS As Long = &H100           ' save in lossless mode
Public Const JXR_DEFAULT As Long = 0                 ' save with quality 80 and no chroma subsampling (4:4:4)
Public Const JXR_LOSSLESS As Long = &H64             ' save lossless
Public Const JXR_PROGRESSIVE As Long = &H2000        ' save as a progressive-JXR (use | to combine with other save flags)

Public Enum FREE_IMAGE_FORMAT
   FIF_UNKNOWN = -1
   FIF_BMP = 0
   FIF_ICO = 1
   FIF_JPEG = 2
   FIF_JNG = 3
   FIF_KOALA = 4
   FIF_LBM = 5
   FIF_IFF = FIF_LBM
   FIF_MNG = 6
   FIF_PBM = 7
   FIF_PBMRAW = 8
   FIF_PCD = 9
   FIF_PCX = 10
   FIF_PGM = 11
   FIF_PGMRAW = 12
   FIF_PNG = 13
   FIF_PPM = 14
   FIF_PPMRAW = 15
   FIF_RAS = 16
   FIF_TARGA = 17
   FIF_TIFF = 18
   FIF_WBMP = 19
   FIF_PSD = 20
   FIF_CUT = 21
   FIF_XBM = 22
   FIF_XPM = 23
   FIF_DDS = 24
   FIF_GIF = 25
   FIF_HDR = 26
   FIF_FAXG3 = 27
   FIF_SGI = 28
   FIF_EXR = 29
   FIF_J2K = 30
   FIF_JP2 = 31
   FIF_PFM = 32
   FIF_PICT = 33
   FIF_RAW = 34
   FIF_WEBP = 35
   FIF_JXR = 36
End Enum
#If False Then
   Const FIF_UNKNOWN = -1
   Const FIF_BMP = 0
   Const FIF_ICO = 1
   Const FIF_JPEG = 2
   Const FIF_JNG = 3
   Const FIF_KOALA = 4
   Const FIF_LBM = 5
   Const FIF_IFF = FIF_LBM
   Const FIF_MNG = 6
   Const FIF_PBM = 7
   Const FIF_PBMRAW = 8
   Const FIF_PCD = 9
   Const FIF_PCX = 10
   Const FIF_PGM = 11
   Const FIF_PGMRAW = 12
   Const FIF_PNG = 13
   Const FIF_PPM = 14
   Const FIF_PPMRAW = 15
   Const FIF_RAS = 16
   Const FIF_TARGA = 17
   Const FIF_TIFF = 18
   Const FIF_WBMP = 19
   Const FIF_PSD = 20
   Const FIF_CUT = 21
   Const FIF_XBM = 22
   Const FIF_XPM = 23
   Const FIF_DDS = 24
   Const FIF_GIF = 25
   Const FIF_HDR = 26
   Const FIF_FAXG3 = 27
   Const FIF_SGI = 28
   Const FIF_EXR = 29
   Const FIF_J2K = 30
   Const FIF_JP2 = 31
   Const FIF_PFM = 32
   Const FIF_PICT = 33
   Const FIF_RAW = 34
   Const FIF_WEBP = 35
   Const FIF_JXR = 36
#End If

Public Enum FREE_IMAGE_LOAD_OPTIONS
   FILO_LOAD_NOPIXELS = FIF_LOAD_NOPIXELS         ' load the image header only (not supported by all plugins)
   FILO_LOAD_DEFAULT = 0
   FILO_GIF_DEFAULT = GIF_DEFAULT
   FILO_GIF_LOAD256 = GIF_LOAD256                 ' load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
   FILO_GIF_PLAYBACK = GIF_PLAYBACK               ' 'play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
   FILO_ICO_DEFAULT = ICO_DEFAULT
   FILO_ICO_MAKEALPHA = ICO_MAKEALPHA             ' convert to 32bpp and create an alpha channel from the AND-mask when loading
   FILO_JPEG_DEFAULT = JPEG_DEFAULT               ' for loading this is a synonym for FILO_JPEG_FAST
   FILO_JPEG_FAST = JPEG_FAST                     ' load the file as fast as possible, sacrificing some quality
   FILO_JPEG_ACCURATE = JPEG_ACCURATE             ' load the file with the best quality, sacrificing some speed
   FILO_JPEG_CMYK = JPEG_CMYK                     ' load separated CMYK "as is" (use 'OR' to combine with other load flags)
   FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE         ' load and rotate according to Exif 'Orientation' tag if available
   FILO_PCD_DEFAULT = PCD_DEFAULT
   FILO_JPEG_GREYSCALE = JPEG_GREYSCALE           ' load and convert to a 8-bit greyscale image
   FILO_PCD_BASE = PCD_BASE                       ' load the bitmap sized 768 x 512
   FILO_PCD_BASEDIV4 = PCD_BASEDIV4               ' load the bitmap sized 384 x 256
   FILO_PCD_BASEDIV16 = PCD_BASEDIV16             ' load the bitmap sized 192 x 128
   FILO_PNG_DEFAULT = PNG_DEFAULT
   FILO_PNG_IGNOREGAMMA = PNG_IGNOREGAMMA         ' avoid gamma correction
   FILO_PSD_CMYK = PSD_CMYK                       ' reads tags for separated CMYK (default is conversion to RGB)
   FILO_PSD_LAB = PSD_LAB                         ' reads tags for CIELab (default is conversion to RGB)
   FILO_RAW_DEFAULT = RAW_DEFAULT                 ' load the file as linear RGB 48-bit
   FILO_RAW_PREVIEW = RAW_PREVIEW                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
   FILO_RAW_DISPLAY = RAW_DISPLAY                 ' load the file as RGB 24-bit
   FILO_RAW_HALFSIZE = RAW_HALFSIZE               ' load the file as half-size color image
   FILO_TARGA_DEFAULT = TARGA_LOAD_RGB888
   FILO_TARGA_LOAD_RGB888 = TARGA_LOAD_RGB888     ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
End Enum
#If False Then
   Const FILO_LOAD_NOPIXELS = &H8000
   Const FILO_LOAD_DEFAULT = 0
   Const FILO_GIF_DEFAULT = GIF_DEFAULT
   Const FILO_GIF_LOAD256 = GIF_LOAD256
   Const FILO_GIF_PLAYBACK = GIF_PLAYBACK
   Const FILO_ICO_DEFAULT = ICO_DEFAULT
   Const FILO_ICO_MAKEALPHA = ICO_MAKEALPHA
   Const FILO_JPEG_DEFAULT = JPEG_DEFAULT
   Const FILO_JPEG_FAST = JPEG_FAST
   Const FILO_JPEG_ACCURATE = JPEG_ACCURATE
   Const FILO_JPEG_CMYK = JPEG_CMYK
   Const FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE
   Const FILO_PCD_DEFAULT = PCD_DEFAULT
   Const FILO_PCD_BASE = PCD_BASE
   Const FILO_PCD_BASEDIV4 = PCD_BASEDIV4
   Const FILO_PCD_BASEDIV16 = PCD_BASEDIV16
   Const FILO_PNG_DEFAULT = PNG_DEFAULT
   Const FILO_PNG_IGNOREGAMMA = PNG_IGNOREGAMMA
   Const FILO_TARGA_DEFAULT = TARGA_LOAD_RGB888
   Const FILO_TARGA_LOAD_RGB888 = TARGA_LOAD_RGB888
#End If

Public Enum FREE_IMAGE_SAVE_OPTIONS
   FISO_SAVE_DEFAULT = 0
   FISO_BMP_DEFAULT = BMP_DEFAULT
   FISO_BMP_SAVE_RLE = BMP_SAVE_RLE
   FISO_EXR_DEFAULT = EXR_DEFAULT                 ' save data as half with piz-based wavelet compression
   FISO_EXR_FLOAT = EXR_FLOAT                     ' save data as float instead of as half (not recommended)
   FISO_EXR_NONE = EXR_NONE                       ' save with no compression
   FISO_EXR_ZIP = EXR_ZIP                         ' save with zlib compression, in blocks of 16 scan lines
   FISO_EXR_PIZ = EXR_PIZ                         ' save with piz-based wavelet compression
   FISO_EXR_PXR24 = EXR_PXR24                     ' save with lossy 24-bit float compression
   FISO_EXR_B44 = EXR_B44                         ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
   FISO_EXR_LC = EXR_LC                           ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
   FISO_JPEG_DEFAULT = JPEG_DEFAULT               ' for saving this is a synonym for FISO_JPEG_QUALITYGOOD
   FISO_JPEG_QUALITYSUPERB = JPEG_QUALITYSUPERB   ' save with superb quality (100:1)
   FISO_JPEG_QUALITYGOOD = JPEG_QUALITYGOOD       ' save with good quality (75:1)
   FISO_JPEG_QUALITYNORMAL = JPEG_QUALITYNORMAL   ' save with normal quality (50:1)
   FISO_JPEG_QUALITYAVERAGE = JPEG_QUALITYAVERAGE ' save with average quality (25:1)
   FISO_JPEG_QUALITYBAD = JPEG_QUALITYBAD         ' save with bad quality (10:1)
   FISO_JPEG_PROGRESSIVE = JPEG_PROGRESSIVE       ' save as a progressive-JPEG (use 'OR' to combine with other save flags)
   FISO_JPEG_SUBSAMPLING_411 = JPEG_SUBSAMPLING_411      ' save with high 4x1 chroma subsampling (4:1:1)
   FISO_JPEG_SUBSAMPLING_420 = JPEG_SUBSAMPLING_420      ' save with medium 2x2 medium chroma subsampling (4:2:0) - default value
   FISO_JPEG_SUBSAMPLING_422 = JPEG_SUBSAMPLING_422      ' save with low 2x1 chroma subsampling (4:2:2)
   FISO_JPEG_SUBSAMPLING_444 = JPEG_SUBSAMPLING_444      ' save with no chroma subsampling (4:4:4)
   FISO_JPEG_OPTIMIZE = JPEG_OPTIMIZE                    ' compute optimal Huffman coding tables (can reduce a few percent of file size)
   FISO_JPEG_BASELINE = JPEG_BASELINE                    ' save basic JPEG, without metadata or any markers
   FISO_PNG_Z_BEST_SPEED = PNG_Z_BEST_SPEED              ' save using ZLib level 1 compression flag (default value is 6)
   FISO_PNG_Z_DEFAULT_COMPRESSION = PNG_Z_DEFAULT_COMPRESSION ' save using ZLib level 6 compression flag (default recommended value)
   FISO_PNG_Z_BEST_COMPRESSION = PNG_Z_BEST_COMPRESSION  ' save using ZLib level 9 compression flag (default value is 6)
   FISO_PNG_Z_NO_COMPRESSION = PNG_Z_NO_COMPRESSION      ' save without ZLib compression
   FISO_PNG_INTERLACED = PNG_INTERLACED           ' save using Adam7 interlacing (use | to combine with other save flags)
   FISO_PNM_DEFAULT = PNM_DEFAULT
   FISO_PNM_SAVE_RAW = PNM_SAVE_RAW               ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
   FISO_PNM_SAVE_ASCII = PNM_SAVE_ASCII           ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
   FISO_TARGA_SAVE_RLE = TARGA_SAVE_RLE           ' if set, the writer saves with RLE compression
   FISO_TIFF_DEFAULT = TIFF_DEFAULT
   FISO_TIFF_CMYK = TIFF_CMYK                     ' stores tags for separated CMYK (use 'OR' to combine with compression flags)
   FISO_TIFF_PACKBITS = TIFF_PACKBITS             ' save using PACKBITS compression
   FISO_TIFF_DEFLATE = TIFF_DEFLATE               ' save using DEFLATE compression (a.k.a. ZLIB compression)
   FISO_TIFF_ADOBE_DEFLATE = TIFF_ADOBE_DEFLATE   ' save using ADOBE DEFLATE compression
   FISO_TIFF_NONE = TIFF_NONE                     ' save without any compression
   FISO_TIFF_CCITTFAX3 = TIFF_CCITTFAX3           ' save using CCITT Group 3 fax encoding
   FISO_TIFF_CCITTFAX4 = TIFF_CCITTFAX4           ' save using CCITT Group 4 fax encoding
   FISO_TIFF_LZW = TIFF_LZW                       ' save using LZW compression
   FISO_TIFF_JPEG = TIFF_JPEG                     ' save using JPEG compression
   FISO_TIFF_LOGLUV = TIFF_LOGLUV                 ' save using LogLuv compression
   FISO_WEBP_LOSSLESS = WEBP_LOSSLESS             ' save in lossless mode
   FISO_JXR_LOSSLESS = JXR_LOSSLESS               ' save lossless
   FISP_JXR_PROGRESSIVE = JXR_PROGRESSIVE         ' save as a progressive-JXR (use | to combine with other save flags)
End Enum
#If False Then
   Const FISO_SAVE_DEFAULT = 0
   Const FISO_BMP_DEFAULT = BMP_DEFAULT
   Const FISO_BMP_SAVE_RLE = BMP_SAVE_RLE
   Const FISO_JPEG_DEFAULT = JPEG_DEFAULT
   Const FISO_JPEG_QUALITYSUPERB = JPEG_QUALITYSUPERB
   Const FISO_JPEG_QUALITYGOOD = JPEG_QUALITYGOOD
   Const FISO_JPEG_QUALITYNORMAL = JPEG_QUALITYNORMAL
   Const FISO_JPEG_QUALITYAVERAGE = JPEG_QUALITYAVERAGE
   Const FISO_JPEG_QUALITYBAD = JPEG_QUALITYBAD
   Const FISO_JPEG_PROGRESSIVE = JPEG_PROGRESSIVE
   Const FISO_JPEG_SUBSAMPLING_411 = JPEG_SUBSAMPLING_411
   Const FISO_JPEG_SUBSAMPLING_420 = JPEG_SUBSAMPLING_420
   Const FISO_JPEG_SUBSAMPLING_422 = JPEG_SUBSAMPLING_422
   Const FISO_JPEG_SUBSAMPLING_444 = JPEG_SUBSAMPLING_444
   Const FISO_PNM_DEFAULT = PNM_DEFAULT
   Const FISO_PNM_SAVE_RAW = PNM_SAVE_RAW
   Const FISO_PNM_SAVE_ASCII = PNM_SAVE_ASCII
   Const FISO_TARGA_SAVE_RLE = TARGA_SAVE_RLE
   Const FISO_TIFF_DEFAULT = TIFF_DEFAULT
   Const FISO_TIFF_CMYK = TIFF_CMYK
   Const FISO_TIFF_PACKBITS = TIFF_PACKBITS
   Const FISO_TIFF_DEFLATE = TIFF_DEFLATE
   Const FISO_TIFF_ADOBE_DEFLATE = TIFF_ADOBE_DEFLATE
   Const FISO_TIFF_NONE = TIFF_NONE
   Const FISO_TIFF_CCITTFAX3 = TIFF_CCITTFAX3
   Const FISO_TIFF_CCITTFAX4 = TIFF_CCITTFAX4
   Const FISO_TIFF_LZW = TIFF_LZW
   Const FISO_TIFF_JPEG = TIFF_JPEG
   Const FISO_WEBP_LOSSLESS = WEBP_LOSSLESS
   Const FISO_JXR_LOSSLESS = JXR_LOSSLESS
   Const FISP_JXR_PROGRESSIVE = JXR_PROGRESSIVE
#End If

Public Enum FREE_IMAGE_TYPE
   FIT_UNKNOWN = 0           ' unknown type
   FIT_BITMAP = 1            ' standard image           : 1-, 4-, 8-, 16-, 24-, 32-bit
   FIT_UINT16 = 2            ' array of unsigned short  : unsigned 16-bit
   FIT_INT16 = 3             ' array of short           : signed 16-bit
   FIT_UINT32 = 4            ' array of unsigned long   : unsigned 32-bit
   FIT_INT32 = 5             ' array of long            : signed 32-bit
   FIT_FLOAT = 6             ' array of float           : 32-bit IEEE floating point
   FIT_DOUBLE = 7            ' array of double          : 64-bit IEEE floating point
   FIT_COMPLEX = 8           ' array of FICOMPLEX       : 2 x 64-bit IEEE floating point
   FIT_RGB16 = 9             ' 48-bit RGB image         : 3 x 16-bit
   FIT_RGBA16 = 10           ' 64-bit RGBA image        : 4 x 16-bit
   FIT_RGBF = 11             ' 96-bit RGB float image   : 3 x 32-bit IEEE floating point
   FIT_RGBAF = 12            ' 128-bit RGBA float image : 4 x 32-bit IEEE floating point
End Enum
#If False Then
   Const FIT_UNKNOWN = 0
   Const FIT_BITMAP = 1
   Const FIT_UINT16 = 2
   Const FIT_INT16 = 3
   Const FIT_UINT32 = 4
   Const FIT_INT32 = 5
   Const FIT_FLOAT = 6
   Const FIT_DOUBLE = 7
   Const FIT_COMPLEX = 8
   Const FIT_RGB16 = 9
   Const FIT_RGBA16 = 10
   Const FIT_RGBF = 11
   Const FIT_RGBAF = 12
#End If

Public Enum FREE_IMAGE_COLOR_TYPE
   FIC_MINISWHITE = 0        ' min value is white
   FIC_MINISBLACK = 1        ' min value is black
   FIC_RGB = 2               ' RGB color model
   FIC_PALETTE = 3           ' color map indexed
   FIC_RGBALPHA = 4          ' RGB color model with alpha channel
   FIC_CMYK = 5              ' CMYK color model
End Enum
#If False Then
   Const FIC_MINISWHITE = 0
   Const FIC_MINISBLACK = 1
   Const FIC_RGB = 2
   Const FIC_PALETTE = 3
   Const FIC_RGBALPHA = 4
   Const FIC_CMYK = 5
#End If

Public Enum FREE_IMAGE_QUANTIZE
   FIQ_WUQUANT = 0           ' Xiaolin Wu color quantization algorithm
   FIQ_NNQUANT = 1           ' NeuQuant neural-net quantization algorithm by Anthony Dekker
End Enum
#If False Then
   Const FIQ_WUQUANT = 0
   Const FIQ_NNQUANT = 1
#End If

Public Enum FREE_IMAGE_DITHER
   FID_FS = 0                ' Floyd & Steinberg error diffusion
   FID_BAYER4x4 = 1          ' Bayer ordered dispersed dot dithering (order 2 dithering matrix)
   FID_BAYER8x8 = 2          ' Bayer ordered dispersed dot dithering (order 3 dithering matrix)
   FID_CLUSTER6x6 = 3        ' Ordered clustered dot dithering (order 3 - 6x6 matrix)
   FID_CLUSTER8x8 = 4        ' Ordered clustered dot dithering (order 4 - 8x8 matrix)
   FID_CLUSTER16x16 = 5      ' Ordered clustered dot dithering (order 8 - 16x16 matrix)
   FID_BAYER16x16 = 6        ' Bayer ordered dispersed dot dithering (order 4 dithering matrix)
End Enum
#If False Then
   Const FID_FS = 0
   Const FID_BAYER4x4 = 1
   Const FID_BAYER8x8 = 2
   Const FID_CLUSTER6x6 = 3
   Const FID_CLUSTER8x8 = 4
   Const FID_CLUSTER16x16 = 5
   Const FID_BAYER16x16 = 6
#End If

Public Enum FREE_IMAGE_JPEG_OPERATION
   FIJPEG_OP_NONE = 0        ' no transformation
   FIJPEG_OP_FLIP_H = 1      ' horizontal flip
   FIJPEG_OP_FLIP_V = 2      ' vertical flip
   FIJPEG_OP_TRANSPOSE = 3   ' transpose across UL-to-LR axis
   FIJPEG_OP_TRANSVERSE = 4  ' transpose across UR-to-LL axis
   FIJPEG_OP_ROTATE_90 = 5   ' 90-degree clockwise rotation
   FIJPEG_OP_ROTATE_180 = 6  ' 180-degree rotation
   FIJPEG_OP_ROTATE_270 = 7  ' 270-degree clockwise (or 90 ccw)
End Enum
#If False Then
   Const FIJPEG_OP_NONE = 0
   Const FIJPEG_OP_FLIP_H = 1
   Const FIJPEG_OP_FLIP_V = 2
   Const FIJPEG_OP_TRANSPOSE = 3
   Const FIJPEG_OP_TRANSVERSE = 4
   Const FIJPEG_OP_ROTATE_90 = 5
   Const FIJPEG_OP_ROTATE_180 = 6
   Const FIJPEG_OP_ROTATE_270 = 7
#End If

Public Enum FREE_IMAGE_TMO
   FITMO_DRAGO03 = 0         ' Adaptive logarithmic mapping (F. Drago, 2003)
   FITMO_REINHARD05 = 1      ' Dynamic range reduction inspired by photoreceptor physiology (E. Reinhard, 2005)
   FITMO_FATTAL02 = 2        ' Gradient domain high dynamic range compression (R. Fattal, 2002)
End Enum
#If False Then
   Const FITMO_DRAGO03 = 0
   Const FITMO_REINHARD05 = 1
   Const FITMO_FATTAL02 = 2
#End If

Public Enum FREE_IMAGE_FILTER
   FILTER_BOX = 0            ' Box, pulse, Fourier window, 1st order (constant) b-spline
   FILTER_BICUBIC = 1        ' Mitchell & Netravali's two-param cubic filter
   FILTER_BILINEAR = 2       ' Bilinear filter
   FILTER_BSPLINE = 3        ' 4th order (cubic) b-spline
   FILTER_CATMULLROM = 4     ' Catmull-Rom spline, Overhauser spline
   FILTER_LANCZOS3 = 5       ' Lanczos3 filter
End Enum
#If False Then
   Const FILTER_BOX = 0
   Const FILTER_BICUBIC = 1
   Const FILTER_BILINEAR = 2
   Const FILTER_BSPLINE = 3
   Const FILTER_CATMULLROM = 4
   Const FILTER_LANCZOS3 = 5
#End If

Public Enum FREE_IMAGE_COLOR_CHANNEL
   FICC_RGB = 0              ' Use red, green and blue channels
   FICC_RED = 1              ' Use red channel
   FICC_GREEN = 2            ' Use green channel
   FICC_BLUE = 3             ' Use blue channel
   FICC_ALPHA = 4            ' Use alpha channel
   FICC_BLACK = 5            ' Use black channel
   FICC_REAL = 6             ' Complex images: use real part
   FICC_IMAG = 7             ' Complex images: use imaginary part
   FICC_MAG = 8              ' Complex images: use magnitude
   FICC_PHASE = 9            ' Complex images: use phase
End Enum
#If False Then
   Const FICC_RGB = 0
   Const FICC_RED = 1
   Const FICC_GREEN = 2
   Const FICC_BLUE = 3
   Const FICC_ALPHA = 4
   Const FICC_BLACK = 5
   Const FICC_REAL = 6
   Const FICC_IMAG = 7
   Const FICC_MAG = 8
   Const FICC_PHASE = 9
#End If

Public Enum FREE_IMAGE_MDTYPE
   FIDT_NOTYPE = 0           ' placeholder
   FIDT_BYTE = 1             ' 8-bit unsigned integer
   FIDT_ASCII = 2            ' 8-bit bytes w/ last byte null
   FIDT_SHORT = 3            ' 16-bit unsigned integer
   FIDT_LONG = 4             ' 32-bit unsigned integer
   FIDT_RATIONAL = 5         ' 64-bit unsigned fraction
   FIDT_SBYTE = 6            ' 8-bit signed integer
   FIDT_UNDEFINED = 7        ' 8-bit untyped data
   FIDT_SSHORT = 8           ' 16-bit signed integer
   FIDT_SLONG = 9            ' 32-bit signed integer
   FIDT_SRATIONAL = 10       ' 64-bit signed fraction
   FIDT_FLOAT = 11           ' 32-bit IEEE floating point
   FIDT_DOUBLE = 12          ' 64-bit IEEE floating point
   FIDT_IFD = 13             ' 32-bit unsigned integer (offset)
   FIDT_PALETTE = 14         ' 32-bit RGBQUAD
End Enum
#If False Then
   Const FIDT_NOTYPE = 0
   Const FIDT_BYTE = 1
   Const FIDT_ASCII = 2
   Const FIDT_SHORT = 3
   Const FIDT_LONG = 4
   Const FIDT_RATIONAL = 5
   Const FIDT_SBYTE = 6
   Const FIDT_UNDEFINED = 7
   Const FIDT_SSHORT = 8
   Const FIDT_SLONG = 9
   Const FIDT_SRATIONAL = 10
   Const FIDT_FLOAT = 11
   Const FIDT_DOUBLE = 12
   Const FIDT_IFD = 13
   Const FIDT_PALETTE = 14
#End If

Public Enum FREE_IMAGE_MDMODEL
   FIMD_NODATA = -1          '
   FIMD_COMMENTS = 0         ' single comment or keywords
   FIMD_EXIF_MAIN = 1        ' Exif-TIFF metadata
   FIMD_EXIF_EXIF = 2        ' Exif-specific metadata
   FIMD_EXIF_GPS = 3         ' Exif GPS metadata
   FIMD_EXIF_MAKERNOTE = 4   ' Exif maker note metadata
   FIMD_EXIF_INTEROP = 5     ' Exif interoperability metadata
   FIMD_IPTC = 6             ' IPTC/NAA metadata
   FIMD_XMP = 7              ' Abobe XMP metadata
   FIMD_GEOTIFF = 8          ' GeoTIFF metadata
   FIMD_ANIMATION = 9        ' Animation metadata
   FIMD_CUSTOM = 10          ' Used to attach other metadata types to a dib
   FIMD_EXIF_RAW = 11        ' Exif metadata as a raw buffer
End Enum
#If False Then
   Const FIMD_NODATA = -1
   Const FIMD_COMMENTS = 0
   Const FIMD_EXIF_MAIN = 1
   Const FIMD_EXIF_EXIF = 2
   Const FIMD_EXIF_GPS = 3
   Const FIMD_EXIF_MAKERNOTE = 4
   Const FIMD_EXIF_INTEROP = 5
   Const FIMD_IPTC = 6
   Const FIMD_XMP = 7
   Const FIMD_GEOTIFF = 8
   Const FIMD_ANIMATION = 9
   Const FIMD_CUSTOM = 10
   Const FIMD_EXIF_RAW = 11
#End If

' These are the GIF_DISPOSAL metadata constants
Public Enum FREE_IMAGE_FRAME_DISPOSAL_METHODS
   FIFD_GIF_DISPOSAL_UNSPECIFIED = 0
   FIFD_GIF_DISPOSAL_LEAVE = 1
   FIFD_GIF_DISPOSAL_BACKGROUND = 2
   FIFD_GIF_DISPOSAL_PREVIOUS = 3
End Enum

' Constants used in FreeImage_FillBackground and FreeImage_EnlargeCanvas
Public Enum FREE_IMAGE_COLOR_OPTIONS
   FI_COLOR_IS_RGB_COLOR = &H0          ' RGBQUAD color is a RGB color (contains no valid alpha channel)
   FI_COLOR_IS_RGBA_COLOR = &H1         ' RGBQUAD color is a RGBA color (contains a valid alpha channel)
   FI_COLOR_FIND_EQUAL_COLOR = &H2      ' For palettized images: lookup equal RGB color from palette
   FI_COLOR_ALPHA_IS_INDEX = &H4        ' The color's rgbReserved member (alpha) contains the palette index to be used
End Enum
Public Const FI_COLOR_PALETTE_SEARCH_MASK = _
      (FI_COLOR_FIND_EQUAL_COLOR Or FI_COLOR_ALPHA_IS_INDEX)     ' Flag to test, if any color lookup is performed

' the next enums are only used by derived functions of the
' FreeImage 3 VB wrapper
Public Enum FREE_IMAGE_CONVERSION_FLAGS
   FICF_MONOCHROME = &H1
   FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   FICF_MONOCHROME_DITHER = &H3
   FICF_GREYSCALE_4BPP = &H4
   FICF_PALLETISED_8BPP = &H8
   FICF_GREYSCALE_8BPP = FICF_PALLETISED_8BPP Or FICF_MONOCHROME
   FICF_GREYSCALE = FICF_GREYSCALE_8BPP
   FICF_RGB_15BPP = &HF
   FICF_RGB_16BPP = &H10
   FICF_RGB_24BPP = &H18
   FICF_RGB_32BPP = &H20
   FICF_RGB_ALPHA = FICF_RGB_32BPP
   FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0
   FICF_REORDER_GREYSCALE_PALETTE = &H1000
End Enum
#If False Then
   Const FICF_MONOCHROME = &H1
   Const FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   Const FICF_MONOCHROME_DITHER = &H3
   Const FICF_GREYSCALE_4BPP = &H4
   Const FICF_PALLETISED_8BPP = &H8
   Const FICF_GREYSCALE_8BPP = FICF_PALLETISED_8BPP Or FICF_MONOCHROME
   Const FICF_GREYSCALE = FICF_GREYSCALE_8BPP
   Const FICF_RGB_15BPP = &HF
   Const FICF_RGB_16BPP = &H10
   Const FICF_RGB_24BPP = &H18
   Const FICF_RGB_32BPP = &H20
   Const FICF_RGB_ALPHA = FICF_RGB_32BPP
   Const FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0
   Const FICF_REORDER_GREYSCALE_PALETTE = &H1000
#End If

Public Enum FREE_IMAGE_COLOR_DEPTH
   FICD_AUTO = &H0
   FICD_MONOCHROME = &H1
   FICD_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   FICD_MONOCHROME_DITHER = &H3
   FICD_1BPP = FICD_MONOCHROME
   FICD_4BPP = &H4
   FICD_8BPP = &H8
   FICD_15BPP = &HF
   FICD_16BPP = &H10
   FICD_24BPP = &H18
   FICD_32BPP = &H20
End Enum
#If False Then
   Const FICD_AUTO = &H0
   Const FICD_MONOCHROME = &H1
   Const FICD_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   Const FICD_MONOCHROME_DITHER = &H3
   Const FICD_1BPP = FICD_MONOCHROME
   Const FICD_4BPP = &H4
   Const FICD_8BPP = &H8
   Const FICD_15BPP = &HF
   Const FICD_16BPP = &H10
   Const FICD_24BPP = &H18
   Const FICD_32BPP = &H20
#End If

Public Enum FREE_IMAGE_ADJUST_MODE
   AM_STRECH = &H1
   AM_DEFAULT = AM_STRECH
   AM_ADJUST_BOTH = AM_STRECH
   AM_ADJUST_WIDTH = &H2
   AM_ADJUST_HEIGHT = &H4
   AM_ADJUST_OPTIMAL_SIZE = &H8
End Enum
#If False Then
   Const AM_STRECH = &H1
   Const AM_DEFAULT = AM_STRECH
   Const AM_ADJUST_BOTH = AM_STRECH
   Const AM_ADJUST_WIDTH = &H2
   Const AM_ADJUST_HEIGHT = &H4
   Const AM_ADJUST_OPTIMAL_SIZE = &H8
#End If

Public Enum FREE_IMAGE_MASK_FLAGS
   FIMF_MASK_NONE = &H0
   FIMF_MASK_FULL_TRANSPARENCY = &H1
   FIMF_MASK_ALPHA_TRANSPARENCY = &H2
   FIMF_MASK_COLOR_TRANSPARENCY = &H4
   FIMF_MASK_FORCE_TRANSPARENCY = &H8
   FIMF_MASK_INVERSE_MASK = &H10
End Enum
#If False Then
   Const FIMF_MASK_NONE = &H0
   Const FIMF_MASK_FULL_TRANSPARENCY = &H1
   Const FIMF_MASK_ALPHA_TRANSPARENCY = &H2
   Const FIMF_MASK_COLOR_TRANSPARENCY = &H4
   Const FIMF_MASK_FORCE_TRANSPARENCY = &H8
   Const FIMF_MASK_INVERSE_MASK = &H10
#End If

Public Enum FREE_IMAGE_COLOR_FORMAT_FLAGS
   FICFF_COLOR_RGB = &H1
   FICFF_COLOR_BGR = &H2
   FICFF_COLOR_PALETTE_INDEX = &H4
   
   FICFF_COLOR_HAS_ALPHA = &H100
   
   FICFF_COLOR_ARGB = FICFF_COLOR_RGB Or FICFF_COLOR_HAS_ALPHA
   FICFF_COLOR_ABGR = FICFF_COLOR_BGR Or FICFF_COLOR_HAS_ALPHA
   
   FICFF_COLOR_FORMAT_ORDER_MASK = FICFF_COLOR_RGB Or FICFF_COLOR_BGR
End Enum
#If False Then
   Const FICFF_COLOR_RGB = &H1
   Const FICFF_COLOR_BGR = &H2
   Const FICFF_COLOR_PALETTE_INDEX = &H4
   
   Const FICFF_COLOR_HAS_ALPHA = &H100
   
   Const FICFF_COLOR_ARGB = FICFF_COLOR_RGB Or FICFF_COLOR_HAS_ALPHA
   Const FICFF_COLOR_ABGR = FICFF_COLOR_BGR Or FICFF_COLOR_HAS_ALPHA
   
   Const FICFF_COLOR_FORMAT_ORDER_MASK = FICFF_COLOR_RGB Or FICFF_COLOR_BGR
#End If

Public Enum FREE_IMAGE_MASK_CREATION_OPTION_FLAGS
   MCOF_CREATE_MASK_IMAGE = &H1
   MCOF_MODIFY_SOURCE_IMAGE = &H2
   MCOF_CREATE_AND_MODIFY = MCOF_CREATE_MASK_IMAGE Or MCOF_MODIFY_SOURCE_IMAGE
End Enum
#If False Then
   Const MCOF_CREATE_MASK_IMAGE = &H1
   Const MCOF_MODIFY_SOURCE_IMAGE = &H2
   Const MCOF_CREATE_AND_MODIFY = MCOF_CREATE_MASK_IMAGE Or MCOF_MODIFY_SOURCE_IMAGE
#End If

Public Enum FREE_IMAGE_TRANSPARENCY_STATE_FLAGS
   FITSF_IGNORE_TRANSPARENCY = &H0
   FITSF_NONTRANSPARENT = &H1
   FITSF_TRANSPARENT = &H2
   FITSF_INCLUDE_ALPHA_TRANSPARENCY = &H4
End Enum
#If False Then
   Const FITSF_IGNORE_TRANSPARENCY = &H0
   Const FITSF_NONTRANSPARENT = &H1
   Const FITSF_TRANSPARENT = &H2
   Const FITSF_INCLUDE_ALPHA_TRANSPARENCY = &H4
#End If

Public Enum FREE_IMAGE_ICON_TRANSPARENCY_OPTION_FLAGS
   ITOF_NO_TRANSPARENCY = &H0
   ITOF_USE_TRANSPARENCY_INFO = &H1
   ITOF_USE_TRANSPARENCY_INFO_ONLY = ITOF_USE_TRANSPARENCY_INFO
   ITOF_USE_COLOR_TRANSPARENCY = &H2
   ITOF_USE_COLOR_TRANSPARENCY_ONLY = ITOF_USE_COLOR_TRANSPARENCY
   ITOF_USE_TRANSPARENCY_INFO_OR_COLOR = ITOF_USE_TRANSPARENCY_INFO Or ITOF_USE_COLOR_TRANSPARENCY
   ITOF_USE_DEFAULT_TRANSPARENCY = ITOF_USE_TRANSPARENCY_INFO_OR_COLOR
   ITOF_USE_COLOR_TOP_LEFT_PIXEL = &H0
   ITOF_USE_COLOR_FIRST_PIXEL = ITOF_USE_COLOR_TOP_LEFT_PIXEL
   ITOF_USE_COLOR_TOP_RIGHT_PIXEL = &H20
   ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL = &H40
   ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL = &H80
   ITOF_USE_COLOR_SPECIFIED = &H100
   ITOF_FORCE_TRANSPARENCY_INFO = &H400
End Enum
#If False Then
   Const ITOF_NO_TRANSPARENCY = &H0
   Const ITOF_USE_TRANSPARENCY_INFO = &H1
   Const ITOF_USE_TRANSPARENCY_INFO_ONLY = ITOF_USE_TRANSPARENCY_INFO
   Const ITOF_USE_COLOR_TRANSPARENCY = &H2
   Const ITOF_USE_COLOR_TRANSPARENCY_ONLY = ITOF_USE_COLOR_TRANSPARENCY
   Const ITOF_USE_TRANSPARENCY_INFO_OR_COLOR = ITOF_USE_TRANSPARENCY_INFO Or ITOF_USE_COLOR_TRANSPARENCY
   Const ITOF_USE_DEFAULT_TRANSPARENCY = ITOF_USE_TRANSPARENCY_INFO_OR_COLOR
   Const ITOF_USE_COLOR_TOP_LEFT_PIXEL = &H0
   Const ITOF_USE_COLOR_FIRST_PIXEL = ITOF_USE_COLOR_TOP_LEFT_PIXEL
   Const ITOF_USE_COLOR_TOP_RIGHT_PIXEL = &H20
   Const ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL = &H40
   Const ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL = &H80
   Const ITOF_USE_COLOR_SPECIFIED = &H100
   Const ITOF_FORCE_TRANSPARENCY_INFO = &H400
#End If

Private Const ITOF_USE_COLOR_BITMASK As Long = ITOF_USE_COLOR_TOP_RIGHT_PIXEL Or _
                                               ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL Or _
                                               ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL Or _
                                               ITOF_USE_COLOR_SPECIFIED

'TANNER'S NOTE: this is publicly declared elsewhere in this project, so I've commented it out here
'Public Type RGBTRIPLE
'   rgbtBlue As Byte
'   rgbtGreen As Byte
'   rgbtRed As Byte
'End Type

Public Const BI_RGB As Long = 0
Public Const BI_RLE8 As Long = 1
Public Const BI_RLE4 As Long = 2
Public Const BI_BITFIELDS As Long = 3
Public Const BI_JPEG As Long = 4
Public Const BI_PNG As Long = 5

Public Type FIICCPROFILE
   Flags As Integer
   Size As Long
   Data As Long
End Type

Public Type FIRGB16
   Red As Integer
   Green As Integer
   Blue As Integer
End Type

Public Type FIRGBA16
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Public Type FIRGBF
   Red As Double
   Green As Double
   Blue As Double
End Type

Public Type FIRGBAF
   Red As Double
   Green As Double
   Blue As Double
   Alpha As Double
End Type

Public Type FICOMPLEX
   r As Double           ' real part
   i As Double           ' imaginary part
End Type

Public Type FITAG
   Key As Long
   Description As Long
   Id As Integer
   Type As Integer
   Count As Long
   Length As Long
   Value As Long
End Type

Public Type FIRATIONAL
   Numerator As Variant
   Denominator As Variant
End Type

Public Type FREE_IMAGE_TAG
   Model As FREE_IMAGE_MDMODEL
   TagPtr As Long
   Key As String
   Description As String
   Id As Long
   Type As FREE_IMAGE_MDTYPE
   Count As Long
   Length As Long
   StringValue As String
   Palette() As RGBQUAD
   RationalValue() As FIRATIONAL
   Value As Variant
End Type

Public Type FreeImageIO
   read_proc As Long
   write_proc As Long
   seek_proc As Long
   tell_proc As Long
End Type

Public Type Plugin
   format_proc As Long
   description_proc As Long
   extension_proc As Long
   regexpr_proc As Long
   open_proc As Long
   close_proc As Long
   pagecount_proc As Long
   pagecapability_proc As Long
   load_proc As Long
   save_proc As Long
   validate_proc As Long
   mime_proc As Long
   supports_export_bpp_proc As Long
   supports_export_type_proc As Long
   supports_icc_profiles_proc As Long
End Type

' the next structures are only used by derived functions of the
' FreeImage 3 VB wrapper
Public Type RGBTRIPLE
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type ScanLineRGBTRIBLE
   Data() As RGBTRIPLE
End Type

Public Type ScanLinesRGBTRIBLE
   Scanline() As ScanLineRGBTRIBLE
End Type

Private Type SAVEARRAY2D
   cDims As Integer
   fFeatures As Integer
   cbElements As Long
   cLocks As Long
   pvData As Long
   cElements1 As Long
   lLbound1 As Long
   cElements2 As Long
   lLbound2 As Long
End Type

'--------------------------------------------------------------------------------
' FreeImage 3 function declarations
'--------------------------------------------------------------------------------

' The FreeImage 3 functions are declared in the same order as they are described
' in the FreeImage 3 API documentation. The documentation's outline is included
' as comments.

'--------------------------------------------------------------------------------
' Bitmap functions
'--------------------------------------------------------------------------------

' General functions
Public Declare Sub FreeImage_Initialise Lib "FreeImage.dll" Alias "_FreeImage_Initialise@4" ( _
  Optional ByVal LoadLocalPluginsOnly As Long)

Public Declare Sub FreeImage_DeInitialise Lib "FreeImage.dll" Alias "_FreeImage_DeInitialise@0" ()

Private Declare Function FreeImage_GetVersionInt Lib "FreeImage.dll" Alias "_FreeImage_GetVersion@0" () As Long

Private Declare Function FreeImage_GetCopyrightMessageInt Lib "FreeImage.dll" Alias "_FreeImage_GetCopyrightMessage@0" () As Long

Public Declare Sub FreeImage_SetOutputMessage Lib "FreeImage.dll" Alias "_FreeImage_SetOutputMessageStdCall@4" ( _
           ByVal omf As Long)


' Bitmap management functions
Public Declare Function FreeImage_Allocate Lib "FreeImage.dll" Alias "_FreeImage_Allocate@24" ( _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal BitsPerPixel As Long, _
  Optional ByVal RedMask As Long, _
  Optional ByVal GreenMask As Long, _
  Optional ByVal BlueMask As Long) As Long

Public Declare Function FreeImage_AllocateT Lib "FreeImage.dll" Alias "_FreeImage_AllocateT@28" ( _
           ByVal ImageType As FREE_IMAGE_TYPE, _
           ByVal Width As Long, _
           ByVal Height As Long, _
  Optional ByVal BitsPerPixel As Long = 8, _
  Optional ByVal RedMask As Long, _
  Optional ByVal GreenMask As Long, _
  Optional ByVal BlueMask As Long) As Long
  
Public Declare Function FreeImage_HasPixelsInt Lib "FreeImage.dll" Alias "_FreeImage_HasPixels@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_Load Lib "FreeImage.dll" Alias "_FreeImage_Load@12" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal FileName As String, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Public Declare Function FreeImage_LoadUInt Lib "FreeImage.dll" Alias "_FreeImage_LoadU@12" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal FileName As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Public Declare Function FreeImage_LoadFromHandle Lib "FreeImage.dll" Alias "_FreeImage_LoadFromHandle@16" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal IO As Long, _
           ByVal Handle As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveInt Lib "FreeImage.dll" Alias "_FreeImage_Save@16" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BITMAP As Long, _
           ByVal FileName As String, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Private Declare Function FreeImage_SaveUInt Lib "FreeImage.dll" Alias "_FreeImage_SaveU@16" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BITMAP As Long, _
           ByVal FileName As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Private Declare Function FreeImage_SaveToHandleInt Lib "FreeImage.dll" Alias "_FreeImage_SaveToHandle@20" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BITMAP As Long, _
           ByVal IO As Long, _
           ByVal Handle As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Public Declare Function FreeImage_Clone Lib "FreeImage.dll" Alias "_FreeImage_Clone@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Sub FreeImage_Unload Lib "FreeImage.dll" Alias "_FreeImage_Unload@4" ( _
           ByVal BITMAP As Long)


' Bitmap information functions
Public Declare Function FreeImage_GetImageType Lib "FreeImage.dll" Alias "_FreeImage_GetImageType@4" ( _
           ByVal BITMAP As Long) As FREE_IMAGE_TYPE

Public Declare Function FreeImage_GetColorsUsed Lib "FreeImage.dll" Alias "_FreeImage_GetColorsUsed@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetBPP Lib "FreeImage.dll" Alias "_FreeImage_GetBPP@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetWidth Lib "FreeImage.dll" Alias "_FreeImage_GetWidth@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetHeight Lib "FreeImage.dll" Alias "_FreeImage_GetHeight@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetLine Lib "FreeImage.dll" Alias "_FreeImage_GetLine@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetPitch Lib "FreeImage.dll" Alias "_FreeImage_GetPitch@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetDIBSize Lib "FreeImage.dll" Alias "_FreeImage_GetDIBSize@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterX@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterY@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Sub FreeImage_SetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterX@8" ( _
           ByVal BITMAP As Long, _
           ByVal Resolution As Long)

Public Declare Sub FreeImage_SetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterY@8" ( _
           ByVal BITMAP As Long, _
           ByVal Resolution As Long)

Public Declare Function FreeImage_GetInfoHeader Lib "FreeImage.dll" Alias "_FreeImage_GetInfoHeader@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetInfo Lib "FreeImage.dll" Alias "_FreeImage_GetInfo@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetColorType Lib "FreeImage.dll" Alias "_FreeImage_GetColorType@4" ( _
           ByVal BITMAP As Long) As FREE_IMAGE_COLOR_TYPE

Public Declare Function FreeImage_GetRedMask Lib "FreeImage.dll" Alias "_FreeImage_GetRedMask@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetGreenMask Lib "FreeImage.dll" Alias "_FreeImage_GetGreenMask@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetBlueMask Lib "FreeImage.dll" Alias "_FreeImage_GetBlueMask@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetTransparencyCount Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyCount@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetTransparencyTable Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyTable@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Sub FreeImage_SetTransparencyTable Lib "FreeImage.dll" Alias "_FreeImage_SetTransparencyTable@12" ( _
           ByVal BITMAP As Long, _
           ByVal TransTablePtr As Long, _
           ByVal Count As Long)

Private Declare Function FreeImage_IsTransparentInt Lib "FreeImage.dll" Alias "_FreeImage_IsTransparent@4" ( _
           ByVal BITMAP As Long) As Long
           
Public Declare Function FreeImage_GetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_GetTransparentIndex@4" ( _
           ByVal BITMAP As Long) As Long
           
Public Declare Function FreeImage_SetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_SetTransparentIndex@8" ( _
           ByVal BITMAP As Long, _
           ByVal Index As Long) As Long

Private Declare Function FreeImage_HasBackgroundColorInt Lib "FreeImage.dll" Alias "_FreeImage_HasBackgroundColor@4" ( _
           ByVal BITMAP As Long) As Long
           
Private Declare Function FreeImage_GetBackgroundColorInt Lib "FreeImage.dll" Alias "_FreeImage_GetBackgroundColor@8" ( _
           ByVal BITMAP As Long, _
           ByRef BackColor As RGBQUAD) As Long

Private Declare Function FreeImage_GetBackgroundColorAsLongInt Lib "FreeImage.dll" Alias "_FreeImage_GetBackgroundColor@8" ( _
           ByVal BITMAP As Long, _
           ByRef BackColor As Long) As Long

Private Declare Function FreeImage_SetBackgroundColorInt Lib "FreeImage.dll" Alias "_FreeImage_SetBackgroundColor@8" ( _
           ByVal BITMAP As Long, _
           ByRef BackColor As RGBQUAD) As Long
           
Private Declare Function FreeImage_SetBackgroundColorAsLongInt Lib "FreeImage.dll" Alias "_FreeImage_SetBackgroundColor@8" ( _
           ByVal BITMAP As Long, _
           ByRef BackColor As Long) As Long

Public Declare Function FreeImage_GetThumbnail Lib "FreeImage.dll" Alias "_FreeImage_GetThumbnail@4" ( _
           ByVal BITMAP As Long) As Long
           
Private Declare Function FreeImage_SetThumbnailInt Lib "FreeImage.dll" Alias "_FreeImage_SetThumbnail@8" ( _
           ByVal BITMAP As Long, ByVal Thumbnail As Long) As Long


' Filetype functions
Public Declare Function FreeImage_GetFileType Lib "FreeImage.dll" Alias "_FreeImage_GetFileType@8" ( _
           ByVal FileName As String, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
  
Public Declare Function FreeImage_GetFileTypeU Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeU@8" ( _
           ByVal FileName As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT

Public Declare Function FreeImage_GetFileTypeFromHandle Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromHandle@12" ( _
           ByVal IO As Long, _
           ByVal Handle As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT

Public Declare Function FreeImage_GetFileTypeFromMemory Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromMemory@8" ( _
           ByVal Stream As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT


' Pixel access functions
Public Declare Function FreeImage_GetBits Lib "FreeImage.dll" Alias "_FreeImage_GetBits@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_GetScanline Lib "FreeImage.dll" Alias "_FreeImage_GetScanLine@8" ( _
           ByVal BITMAP As Long, _
           ByVal Scanline As Long) As Long

Private Declare Function FreeImage_GetPixelIndexInt Lib "FreeImage.dll" Alias "_FreeImage_GetPixelIndex@16" ( _
           ByVal BITMAP As Long, _
           ByVal x As Long, _
           ByVal y As Long, _
           ByRef Value As Byte) As Long
        
        
' Conversion functions
Public Declare Function FreeImage_ConvertTo4Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo4Bits@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertTo8Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo8Bits@4" ( _
           ByVal BITMAP As Long) As Long
           
Public Declare Function FreeImage_ConvertToGreyscale Lib "FreeImage.dll" Alias "_FreeImage_ConvertToGreyscale@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits555@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits565@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertTo24Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo24Bits@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertTo32Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo32Bits@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ColorQuantize Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantize@8" ( _
           ByVal BITMAP As Long, _
           ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE) As Long
           
Private Declare Function FreeImage_ColorQuantizeExInt Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantizeEx@20" ( _
           ByVal BITMAP As Long, _
  Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, _
  Optional ByVal PaletteSize As Long = 256, _
  Optional ByVal ReserveSize As Long = 0, _
  Optional ByVal ReservePalettePtr As Long = 0) As Long

Public Declare Function FreeImage_Threshold Lib "FreeImage.dll" Alias "_FreeImage_Threshold@8" ( _
           ByVal BITMAP As Long, _
           ByVal Threshold As Byte) As Long

Public Declare Function FreeImage_Dither Lib "FreeImage.dll" Alias "_FreeImage_Dither@8" ( _
           ByVal BITMAP As Long, _
           ByVal DitherMethod As FREE_IMAGE_DITHER) As Long

Private Declare Function FreeImage_ConvertToStandardTypeInt Lib "FreeImage.dll" Alias "_FreeImage_ConvertToStandardType@8" ( _
           ByVal BITMAP As Long, _
           ByVal ScaleLinear As Long) As Long

Private Declare Function FreeImage_ConvertToTypeInt Lib "FreeImage.dll" Alias "_FreeImage_ConvertToType@12" ( _
           ByVal BITMAP As Long, _
           ByVal DestinationType As FREE_IMAGE_TYPE, _
           ByVal ScaleLinear As Long) As Long

Public Declare Function FreeImage_ConvertToFloat Lib "FreeImage.dll" Alias "_FreeImage_ConvertToFloat@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertToRGBF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBF@4" ( _
           ByVal BITMAP As Long) As Long

'Manually patched by Tanner:
Public Declare Function FreeImage_ConvertToRGBAF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBAF@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertToUINT16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToUINT16@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_ConvertToRGB16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGB16@4" ( _
           ByVal BITMAP As Long) As Long

' Tone mapping operators
Public Declare Function FreeImage_ToneMapping Lib "FreeImage.dll" Alias "_FreeImage_ToneMapping@24" ( _
           ByVal BITMAP As Long, _
           ByVal Operator As FREE_IMAGE_TMO, _
  Optional ByVal FirstArgument As Double, _
  Optional ByVal SecondArgument As Double) As Long
  
Public Declare Function FreeImage_TmoDrago03 Lib "FreeImage.dll" Alias "_FreeImage_TmoDrago03@20" ( _
           ByVal BITMAP As Long, _
  Optional ByVal Gamma As Double = 2.2, _
  Optional ByVal Exposure As Double) As Long
  
Public Declare Function FreeImage_TmoReinhard05 Lib "FreeImage.dll" Alias "_FreeImage_TmoReinhard05@20" ( _
           ByVal BITMAP As Long, _
  Optional ByVal Intensity As Double, _
  Optional ByVal Contrast As Double) As Long

Public Declare Function FreeImage_TmoReinhard05Ex Lib "FreeImage.dll" Alias "_FreeImage_TmoReinhard05Ex@36" ( _
           ByVal BITMAP As Long, _
  Optional ByVal Intensity As Double, _
  Optional ByVal Contrast As Double, _
  Optional ByVal Adaptation As Double = 1, _
  Optional ByVal ColorCorrection As Double) As Long

Public Declare Function FreeImage_TmoFattal02 Lib "FreeImage.dll" Alias "_FreeImage_TmoFattal02@20" ( _
           ByVal BITMAP As Long, _
  Optional ByVal ColorSaturation As Double = 0.5, _
  Optional ByVal Attenuation As Double = 0.85) As Long


' ICC profile functions
Private Declare Function FreeImage_GetICCProfileInt Lib "FreeImage.dll" Alias "_FreeImage_GetICCProfile@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Function FreeImage_CreateICCProfile Lib "FreeImage.dll" Alias "_FreeImage_CreateICCProfile@12" ( _
           ByVal BITMAP As Long, _
           ByRef Data As Long, _
           ByVal Size As Long) As Long

Public Declare Sub FreeImage_DestroyICCProfile Lib "FreeImage.dll" Alias "_FreeImage_DestroyICCProfile@4" ( _
           ByVal BITMAP As Long)


' Plugin functions
Public Declare Function FreeImage_GetFIFCount Lib "FreeImage.dll" Alias "_FreeImage_GetFIFCount@0" () As Long

Public Declare Function FreeImage_SetPluginEnabled Lib "FreeImage.dll" Alias "_FreeImage_SetPluginEnabled@8" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal Value As Long) As Long

Public Declare Function FreeImage_IsPluginEnabled Lib "FreeImage.dll" Alias "_FreeImage_IsPluginEnabled@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Public Declare Function FreeImage_GetFIFFromFormat Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFormat@4" ( _
           ByVal Format As String) As FREE_IMAGE_FORMAT

Public Declare Function FreeImage_GetFIFFromMime Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromMime@4" ( _
           ByVal MimeType As String) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_GetFIFMimeTypeInt Lib "FreeImage.dll" Alias "_FreeImage_GetFIFMimeType@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_GetFormatFromFIFInt Lib "FreeImage.dll" Alias "_FreeImage_GetFormatFromFIF@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_GetFIFExtensionListInt Lib "FreeImage.dll" Alias "_FreeImage_GetFIFExtensionList@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_GetFIFDescriptionInt Lib "FreeImage.dll" Alias "_FreeImage_GetFIFDescription@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Public Declare Function FreeImage_GetFIFFromFilename Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFilename@4" ( _
           ByVal FileName As String) As FREE_IMAGE_FORMAT

Public Declare Function FreeImage_GetFIFFromFilenameU Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFilenameU@4" ( _
           ByVal FileName As Long) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_FIFSupportsReadingInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsReading@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsWritingInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsWriting@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsExportTypeInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportType@8" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal ImageType As FREE_IMAGE_TYPE) As Long

Private Declare Function FreeImage_FIFSupportsExportBPPInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportBPP@8" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BitsPerPixel As Long) As Long

Private Declare Function FreeImage_FIFSupportsICCProfilesInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsICCProfiles@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long
           
Private Declare Function FreeImage_FIFSupportsNoPixelsInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsNoPixels@4" ( _
           ByVal Format As FREE_IMAGE_FORMAT) As Long

Public Declare Function FreeImage_RegisterLocalPlugin Lib "FreeImage.dll" Alias "_FreeImage_RegisterLocalPlugin@20" ( _
           ByVal InitProcAddress As Long, _
  Optional ByVal Format As String, _
  Optional ByVal Description As String, _
  Optional ByVal Extension As String, _
  Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT

Public Declare Function FreeImage_RegisterExternalPlugin Lib "FreeImage.dll" Alias "_FreeImage_RegisterExternalPlugin@20" ( _
           ByVal Path As String, _
  Optional ByVal Format As String, _
  Optional ByVal Description As String, _
  Optional ByVal Extension As String, _
  Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT


' Multipage functions
Private Declare Function FreeImage_OpenMultiBitmapInt Lib "FreeImage.dll" Alias "_FreeImage_OpenMultiBitmap@24" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal FileName As String, _
           ByVal CreateNew As Long, _
           ByVal ReadOnly As Long, _
           ByVal KeepCacheInMemory As Long, _
           ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_CloseMultiBitmapInt Lib "FreeImage.dll" Alias "_FreeImage_CloseMultiBitmap@8" ( _
           ByVal BITMAP As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Public Declare Function FreeImage_GetPageCount Lib "FreeImage.dll" Alias "_FreeImage_GetPageCount@4" ( _
           ByVal BITMAP As Long) As Long

Public Declare Sub FreeImage_AppendPage Lib "FreeImage.dll" Alias "_FreeImage_AppendPage@8" ( _
           ByVal BITMAP As Long, _
           ByVal PageBitmap As Long)

Public Declare Sub FreeImage_InsertPage Lib "FreeImage.dll" Alias "_FreeImage_InsertPage@12" ( _
           ByVal BITMAP As Long, _
           ByVal Page As Long, _
           ByVal PageBitmap As Long)

Public Declare Sub FreeImage_DeletePage Lib "FreeImage.dll" Alias "_FreeImage_DeletePage@8" ( _
           ByVal BITMAP As Long, _
           ByVal Page As Long)

Public Declare Function FreeImage_LockPage Lib "FreeImage.dll" Alias "_FreeImage_LockPage@8" ( _
           ByVal BITMAP As Long, _
           ByVal Page As Long) As Long

Private Declare Sub FreeImage_UnlockPageInt Lib "FreeImage.dll" Alias "_FreeImage_UnlockPage@12" ( _
           ByVal BITMAP As Long, _
           ByVal PageBitmap As Long, _
           ByVal ApplyChanges As Long)

' Memory I/O streams
Public Declare Function FreeImage_OpenMemory Lib "FreeImage.dll" Alias "_FreeImage_OpenMemory@8" ( _
  Optional ByRef Data As Byte, _
  Optional ByVal SizeInBytes As Long) As Long
  
Public Declare Function FreeImage_OpenMemoryByPtr Lib "FreeImage.dll" Alias "_FreeImage_OpenMemory@8" ( _
  Optional ByVal DataPtr As Long, _
  Optional ByVal SizeInBytes As Long) As Long

Public Declare Sub FreeImage_CloseMemory Lib "FreeImage.dll" Alias "_FreeImage_CloseMemory@4" ( _
           ByVal Stream As Long)

Public Declare Function FreeImage_LoadFromMemory Lib "FreeImage.dll" Alias "_FreeImage_LoadFromMemory@12" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveToMemoryInt Lib "FreeImage.dll" Alias "_FreeImage_SaveToMemory@16" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BITMAP As Long, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Private Declare Function FreeImage_AcquireMemoryInt Lib "FreeImage.dll" Alias "_FreeImage_AcquireMemory@12" ( _
           ByVal Stream As Long, _
           ByRef DataPtr As Long, _
           ByRef SizeInBytes As Long) As Long

Public Declare Function FreeImage_TellMemory Lib "FreeImage.dll" Alias "_FreeImage_TellMemory@4" ( _
           ByVal Stream As Long) As Long

Private Declare Function FreeImage_SeekMemoryInt Lib "FreeImage.dll" Alias "_FreeImage_SeekMemory@12" ( _
           ByVal Stream As Long, _
           ByVal Offset As Long, _
           ByVal Origin As Long) As Long
           
Public Declare Function FreeImage_ReadMemory Lib "FreeImage.dll" Alias "_FreeImage_ReadMemory@16" ( _
           ByVal BufferPtr As Long, _
           ByVal Size As Long, _
           ByVal Count As Long, _
           ByVal Stream As Long) As Long
           
Public Declare Function FreeImage_WriteMemory Lib "FreeImage.dll" Alias "_FreeImage_WriteMemory@16" ( _
           ByVal BufferPtr As Long, _
           ByVal Size As Long, _
           ByVal Count As Long, _
           ByVal Stream As Long) As Long
           
Public Declare Function FreeImage_LoadMultiBitmapFromMemory Lib "FreeImage.dll" Alias "_FreeImage_LoadMultiBitmapFromMemory@12" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Public Declare Function FreeImage_SaveMultiBitmapToMemory Lib "FreeImage.dll" Alias "_FreeImage_SaveMultiBitmapToMemory@16" ( _
           ByVal Format As FREE_IMAGE_FORMAT, _
           ByVal BITMAP As Long, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long


' Compression functions
Public Declare Function FreeImage_ZLibCompress Lib "FreeImage.dll" Alias "_FreeImage_ZLibCompress@16" ( _
           ByVal TargetPtr As Long, _
           ByVal TargetSize As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long

Public Declare Function FreeImage_ZLibUncompress Lib "FreeImage.dll" Alias "_FreeImage_ZLibUncompress@16" ( _
           ByVal TargetPtr As Long, _
           ByVal TargetSize As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long

Public Declare Function FreeImage_ZLibGZip Lib "FreeImage.dll" Alias "_FreeImage_ZLibGZip@16" ( _
           ByVal TargetPtr As Long, _
           ByVal TargetSize As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long
           
Public Declare Function FreeImage_ZLibGUnzip Lib "FreeImage.dll" Alias "_FreeImage_ZLibGUnzip@16" ( _
           ByVal TargetPtr As Long, _
           ByVal TargetSize As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long

Public Declare Function FreeImage_ZLibCRC32 Lib "FreeImage.dll" Alias "_FreeImage_ZLibCRC32@12" ( _
           ByVal CRC As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long


' Helper functions
Private Declare Function FreeImage_IsLittleEndianInt Lib "FreeImage.dll" Alias "_FreeImage_IsLittleEndian@0" () As Long

Private Declare Function FreeImage_LookupX11ColorInt Lib "FreeImage.dll" Alias "_FreeImage_LookupX11Color@16" ( _
           ByVal Color As String, _
           ByRef Red As Long, _
           ByRef Green As Long, _
           ByRef Blue As Long) As Long

Private Declare Function FreeImage_LookupSVGColorInt Lib "FreeImage.dll" Alias "_FreeImage_LookupSVGColor@16" ( _
           ByVal Color As String, _
           ByRef Red As Long, _
           ByRef Green As Long, _
           ByRef Blue As Long) As Long


'--------------------------------------------------------------------------------
' Metadata functions
'--------------------------------------------------------------------------------

' Metadata iterator
Public Declare Function FreeImage_FindFirstMetadata Lib "FreeImage.dll" Alias "_FreeImage_FindFirstMetadata@12" ( _
           ByVal Model As FREE_IMAGE_MDMODEL, _
           ByVal BITMAP As Long, _
           ByRef Tag As Long) As Long

Public Declare Function FreeImage_FindNextMetadataInt Lib "FreeImage.dll" Alias "_FreeImage_FindNextMetadata@8" ( _
           ByVal hFind As Long, _
           ByRef Tag As Long) As Long

Public Declare Sub FreeImage_FindCloseMetadata Lib "FreeImage.dll" Alias "_FreeImage_FindCloseMetadata@4" ( _
           ByVal hFind As Long)
           
Public Declare Function FreeImage_CloneMetadataInt Lib "FreeImage.dll" Alias "_FreeImage_CloneMetadata@8" ( _
           ByVal BitmapDst As Long, _
           ByVal BitmapSrc As Long) As Long

' Metadata helper functions
Public Declare Function FreeImage_GetMetadataCount Lib "FreeImage.dll" Alias "_FreeImage_GetMetadataCount@8" ( _
           ByVal Model As Long, _
           ByVal BITMAP As Long) As Long


'--------------------------------------------------------------------------------
' Toolkit functions
'--------------------------------------------------------------------------------

' Rotating and flipping
Public Declare Function FreeImage_RotateClassic Lib "FreeImage.dll" Alias "_FreeImage_RotateClassic@12" ( _
           ByVal BITMAP As Long, _
           ByVal Angle As Double) As Long

Public Declare Function FreeImage_Rotate Lib "FreeImage.dll" Alias "_FreeImage_Rotate@16" ( _
           ByVal BITMAP As Long, _
           ByVal Angle As Double, _
  Optional ByRef Color As Any = 0) As Long

Private Declare Function FreeImage_RotateExInt Lib "FreeImage.dll" Alias "_FreeImage_RotateEx@48" ( _
           ByVal BITMAP As Long, _
           ByVal Angle As Double, _
           ByVal ShiftX As Double, _
           ByVal ShiftY As Double, _
           ByVal OriginX As Double, _
           ByVal OriginY As Double, _
           ByVal UseMask As Long) As Long


' Upsampling and downsampling
Public Declare Function FreeImage_Rescale Lib "FreeImage.dll" Alias "_FreeImage_Rescale@16" ( _
           ByVal BITMAP As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal Filter As FREE_IMAGE_FILTER) As Long
           
Public Declare Function FreeImage_RescaleRect Lib "FreeImage.dll" Alias "_FreeImage_RescaleRect@32" ( _
           ByVal BITMAP As Long, _
           ByVal Left As Long, _
           ByVal Top As Long, _
           ByVal Right As Long, _
           ByVal Bottom As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal Filter As FREE_IMAGE_FILTER) As Long
           
Private Declare Function FreeImage_MakeThumbnailInt Lib "FreeImage.dll" Alias "_FreeImage_MakeThumbnail@12" ( _
           ByVal BITMAP As Long, _
           ByVal MaxPixelSize As Long, _
  Optional ByVal Convert As Long) As Long

Public Declare Function FreeImage_SwapPaletteIndices Lib "FreeImage.dll" Alias "_FreeImage_SwapPaletteIndices@12" ( _
           ByVal BITMAP As Long, _
           ByRef IndexA As Byte, _
           ByRef IndexB As Byte) As Long

' Channel processing
Public Declare Function FreeImage_GetChannel Lib "FreeImage.dll" Alias "_FreeImage_GetChannel@8" ( _
           ByVal BITMAP As Long, _
           ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long

Public Declare Function FreeImage_GetComplexChannel Lib "FreeImage.dll" Alias "_FreeImage_GetComplexChannel@8" ( _
           ByVal BITMAP As Long, _
           ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long

' Copy / Paste / Composite routines
Public Declare Function FreeImage_Copy Lib "FreeImage.dll" Alias "_FreeImage_Copy@20" ( _
           ByVal BITMAP As Long, _
           ByVal Left As Long, _
           ByVal Top As Long, _
           ByVal Right As Long, _
           ByVal Bottom As Long) As Long

Public Declare Function FreeImage_Composite Lib "FreeImage.dll" Alias "_FreeImage_Composite@16" ( _
           ByVal BITMAP As Long, _
  Optional ByVal UseFileBackColor As Long, _
  Optional ByRef AppBackColor As Any, _
  Optional ByVal BackgroundBitmap As Long) As Long

Private Declare Function FreeImage_PreMultiplyWithAlphaInt Lib "FreeImage.dll" Alias "_FreeImage_PreMultiplyWithAlpha@4" ( _
           ByVal BITMAP As Long) As Long
           
Public Declare Function FreeImage_FillBackground Lib "FreeImage.dll" Alias "_FreeImage_FillBackground@12" ( _
           ByVal BITMAP As Long, _
           ByRef Color As Any, _
  Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As Long

Public Declare Function FreeImage_EnlargeCanvas Lib "FreeImage.dll" Alias "_FreeImage_EnlargeCanvas@28" ( _
           ByVal BITMAP As Long, _
           ByVal Left As Long, _
           ByVal Top As Long, _
           ByVal Right As Long, _
           ByVal Bottom As Long, _
           ByRef Color As Any, _
  Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As Long

Public Declare Function FreeImage_AllocateEx Lib "FreeImage.dll" Alias "_FreeImage_AllocateEx@36" ( _
           ByVal Width As Long, _
           ByVal Height As Long, _
  Optional ByVal BitsPerPixel As Long = 8, _
  Optional ByRef Color As Any, _
  Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, _
  Optional ByVal PalettePtr As Long = 0, _
  Optional ByVal RedMask As Long = 0, _
  Optional ByVal GreenMask As Long = 0, _
  Optional ByVal BlueMask As Long = 0) As Long
           
Public Declare Function FreeImage_AllocateExT Lib "FreeImage.dll" Alias "_FreeImage_AllocateExT@36" ( _
           ByVal ImageType As FREE_IMAGE_TYPE, _
           ByVal Width As Long, _
           ByVal Height As Long, _
  Optional ByVal BitsPerPixel As Long = 8, _
  Optional ByRef Color As Any, _
  Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, _
  Optional ByVal PalettePtr As Long, _
  Optional ByVal RedMask As Long, _
  Optional ByVal GreenMask As Long, _
  Optional ByVal BlueMask As Long) As Long

' miscellaneous algorithms
'Public Declare Function FreeImage_MultigridPoissonSolver Lib "FreeImage.dll" Alias "_FreeImage_MultigridPoissonSolver@8" ( _
           ByVal LaplacianBitmap As Long, _
  Optional ByVal Cyles As Long = 3) As Long


'--------------------------------------------------------------------------------
' Line converting functions
'--------------------------------------------------------------------------------

' convert to 4 bpp
Public Declare Sub FreeImage_ConvertLine1To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To4@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine8To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)
           
Public Declare Sub FreeImage_ConvertLine16To4_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
                     
Public Declare Sub FreeImage_ConvertLine16To4_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine24To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine32To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To4@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 8 bpp
Public Declare Sub FreeImage_ConvertLine1To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine4To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To8@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To8_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To8_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To8@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To8@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           

' convert to 16 bpp
Public Declare Sub FreeImage_ConvertLine1To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_555@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_555@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_555@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16_565_To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_565_To16_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine1To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_565@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_565@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_565@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16_555_To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_555_To16_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 24 bpp
Public Declare Sub FreeImage_ConvertLine1To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To24@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To24@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16To24_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To24_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To24@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 32 bpp
Public Declare Sub FreeImage_ConvertLine1To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To32@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To32@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To32@16" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16To32_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_555@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To32_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_565@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To32@12" ( _
           ByVal TargetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
          
          

'--------------------------------------------------------------------------------
' Initialization functions
'--------------------------------------------------------------------------------

Public Function FreeImage_IsAvailable(Optional ByRef Version As String) As Boolean

   On Error Resume Next
   Version = FreeImage_GetVersion()
   FreeImage_IsAvailable = (Err.Number = ERROR_SUCCESS)
   On Error GoTo 0

End Function

'--------------------------------------------------------------------------------
' String returning functions wrappers
'--------------------------------------------------------------------------------

Public Function FreeImage_GetVersion() As String

   ' This function returns the version of the FreeImage 3 library
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   FreeImage_GetVersion = pGetStringFromPointerA(FreeImage_GetVersionInt)

End Function

Public Function FreeImage_GetCopyrightMessage() As String

   ' This function returns the copyright message of the FreeImage 3 library
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   FreeImage_GetCopyrightMessage = pGetStringFromPointerA(FreeImage_GetCopyrightMessageInt)

End Function

Public Function FreeImage_GetFormatFromFIF(ByVal Format As FREE_IMAGE_FORMAT) As String

   ' This function returns the result of the 'FreeImage_GetFormatFromFIF' function
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   ' The parameter 'Format' works according to the FreeImage 3 API documentation.
   
   FreeImage_GetFormatFromFIF = pGetStringFromPointerA(FreeImage_GetFormatFromFIFInt(Format))

End Function

Public Function FreeImage_GetFIFExtensionList(ByVal Format As FREE_IMAGE_FORMAT) As String

   ' This function returns the result of the 'FreeImage_GetFIFExtensionList' function
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   ' The parameter 'Format' works according to the FreeImage 3 API documentation.
   
   FreeImage_GetFIFExtensionList = pGetStringFromPointerA(FreeImage_GetFIFExtensionListInt(Format))

End Function

Public Function FreeImage_GetFIFDescription(ByVal Format As FREE_IMAGE_FORMAT) As String

   ' This function returns the result of the 'FreeImage_GetFIFDescription' function
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   ' The parameter 'Format' works according to the FreeImage 3 API documentation.
   
   FreeImage_GetFIFDescription = pGetStringFromPointerA(FreeImage_GetFIFDescriptionInt(Format))

End Function


Public Function FreeImage_GetFIFMimeType(ByVal Format As FREE_IMAGE_FORMAT) As String
   
   ' This function returns the result of the 'FreeImage_GetFIFMimeType' function
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   ' The parameter 'Format' works according to the FreeImage 3 API documentation.
   
   FreeImage_GetFIFMimeType = pGetStringFromPointerA(FreeImage_GetFIFMimeTypeInt(Format))
   
End Function

Public Function FreeImage_GetPixelIndex(ByVal BITMAP As Long, _
                                        ByVal x As Long, _
                                        ByVal y As Long, _
                                        ByRef Value As Byte) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_GetPixelIndex = (FreeImage_GetPixelIndexInt(BITMAP, x, y, Value) = 1)

End Function


Public Function FreeImage_GetPaletteExLong(ByVal BITMAP As Long) As Long()

Dim tSA As SAVEARRAY1D
Dim lpSA As Long

   ' This function returns a VB style array of type Long, containing
   ' the palette data of the Bitmap. This array provides read and write access
   ' to the actual palette data provided by FreeImage. This is done by
   ' creating a VB array with an own SAFEARRAY descriptor making the
   ' array point to the palette pointer returned by FreeImage_GetPalette().
   
   ' The function actually returns an array of type RGBQUAD with each
   ' element packed into a Long. This is possible, since the RGBQUAD
   ' structure is also four bytes in size. Palette data, stored in an
   ' array of type Long may be passed ByRef to a function through an
   ' optional paremeter. For an example have a look at function
   ' FreeImage_ConvertColorDepth()
   
   ' This makes you use code like you would in C/C++:
   
   ' // this code assumes there is a bitmap loaded and
   ' // present in a variable called ‘dib’
   ' if(FreeImage_GetBPP(Bitmap) == 8) {
   '   // Build a greyscale palette
   '   RGBQUAD *pal = FreeImage_GetPalette(Bitmap);
   '   for (int i = 0; i < 256; i++) {
   '     pal[i].rgbRed = i;
   '     pal[i].rgbGreen = i;
   '     pal[i].rgbBlue = i;
   '   }
   
   ' As in C/C++ the array is only valid while the DIB is loaded and the
   ' palette data remains where the pointer returned by FreeImage_GetPalette()
   ' has pointed to when this function was called. So, a good thing would
   ' be, not to keep the returned array in scope over the lifetime of the
   ' Bitmap. Best practise is, to use this function within another routine and
   ' assign the return value (the array) to a local variable only. As soon
   ' as this local variable goes out of scope (when the calling function
   ' returns to it's caller), the array and the descriptor is automatically
   ' cleaned up by VB.
   
   ' This function does not make a deep copy of the palette data, but only
   ' wraps a VB array around the FreeImage palette data. So, it can be called
   ' frequently "on demand" or somewhat "in place" without a significant
   ' performance loss.
   
   ' To learn more about this technique I recommend reading chapter 2 (Leveraging
   ' Arrays) of Matthew Curland's book "Advanced Visual Basic 6"
   
   ' The parameter 'Bitmap' works according to the FreeImage 3 API documentation.
   
   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the 'FreeImage_DestroyLockedArray' function.

   
   If (BITMAP) Then
      
      ' create a proper SAVEARRAY descriptor
      With tSA
         .cbElements = 4                              ' size in bytes of RGBQUAD structure
         .cDims = 1                                   ' the array has only 1 dimension
         .cElements = FreeImage_GetColorsUsed(BITMAP) ' the number of elements in the array is
                                                      ' the number of used colors in the Bitmap
         .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE     ' need AUTO and FIXEDSIZE for safety issues,
                                                      ' so the array can not be modified in size
                                                      ' or erased; according to Matthew Curland never
                                                      ' use FIXEDSIZE alone
         .pvData = FreeImage_GetPalette(BITMAP)       ' let the array point to the memory block, the
                                                      ' FreeImage palette pointer points to
      End With
      
      ' allocate memory for an array descriptor
      ' we cannot use the memory block used by tSA, since it is
      ' released when tSA goes out of scope, leaving us with an
      ' array with zeroed descriptor
      ' we use nearly the same method that VB uses, so VB is able
      ' to cleanup the array variable and it's descriptor; the
      ' array data is not touched when cleaning up, since both AUTO
      ' and FIXEDSIZE flags are set
      Call SafeArrayAllocDescriptor(1, lpSA)
      
      ' copy our own array descriptor over the descriptor allocated
      ' by SafeArrayAllocDescriptor; lpSA is a pointer to that memory
      ' location
      Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
      
      ' the implicit variable named as the function is an array variable in VB
      ' make it point to the allocated array descriptor
      Call CopyMemory(ByVal VarPtrArray(FreeImage_GetPaletteExLong), lpSA, 4)
   End If

End Function

'--------------------------------------------------------------------------------
' BOOL/Boolean returning functions wrappers
'--------------------------------------------------------------------------------

Public Function FreeImage_HasPixels(ByVal BITMAP As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_HasPixels = (FreeImage_HasPixelsInt(BITMAP) = 1)

End Function

Public Function FreeImage_Save(ByVal Format As FREE_IMAGE_FORMAT, _
                               ByVal BITMAP As Long, _
                               ByVal FileName As String, _
                      Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_Save = (FreeImage_SaveUInt(Format, BITMAP, StrPtr(FileName), Flags) = 1)

End Function

Public Function FreeImage_SaveToHandle(ByVal Format As FREE_IMAGE_FORMAT, _
                                       ByVal BITMAP As Long, _
                                       ByVal IO As Long, _
                                       ByVal Handle As Long, _
                              Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_SaveToHandle = (FreeImage_SaveToHandleInt(Format, BITMAP, IO, Handle, Flags) = 1)

End Function

Public Function FreeImage_IsTransparent(ByVal BITMAP As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_IsTransparent = (FreeImage_IsTransparentInt(BITMAP) = 1)

End Function
           
Public Function FreeImage_HasBackgroundColor(ByVal BITMAP As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_HasBackgroundColor = (FreeImage_HasBackgroundColorInt(BITMAP) = 1)

End Function

Public Function FreeImage_GetBackgroundColor(ByVal BITMAP As Long, _
                                             ByRef BackColor As RGBQUAD) As Boolean
   
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_GetBackgroundColor = (FreeImage_GetBackgroundColorInt(BITMAP, BackColor) = 1)
   
End Function

Public Function FreeImage_GetBackgroundColorAsLong(ByVal BITMAP As Long, _
                                                   ByRef BackColor As Long) As Boolean
   
   ' This function gets the background color of an image as FreeImage_GetBackgroundColor() does but
   ' provides it's result as a Long value.

   FreeImage_GetBackgroundColorAsLong = (FreeImage_GetBackgroundColorAsLongInt(BITMAP, BackColor) = 1)
   
End Function

Public Function FreeImage_GetBackgroundColorEx(ByVal BITMAP As Long, _
                                               ByRef Alpha As Byte, _
                                               ByRef Red As Byte, _
                                               ByRef Green As Byte, _
                                               ByRef Blue As Byte) As Boolean
                                              
Dim bkcolor As RGBQUAD

   ' This function gets the background color of an image as FreeImage_GetBackgroundColor() does but
   ' provides it's result as four different byte values, one for each color component.
                                              
   FreeImage_GetBackgroundColorEx = (FreeImage_GetBackgroundColorInt(BITMAP, bkcolor) = 1)
   With bkcolor
      Alpha = .Alpha
      Red = .Red
      Green = .Green
      Blue = .Blue
   End With

End Function

Public Function FreeImage_SetBackgroundColor(ByVal BITMAP As Long, _
                                             ByRef BackColor As RGBQUAD) As Boolean
                                             
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_SetBackgroundColor = (FreeImage_SetBackgroundColorInt(BITMAP, BackColor) = 1)
                                             
End Function

Public Function FreeImage_SetBackgroundColorAsLong(ByVal BITMAP As Long, _
                                                   ByVal BackColor As Long) As Boolean
                                             
   ' This function sets the background color of an image as FreeImage_SetBackgroundColor() does but
   ' the color value to set must be provided as a Long value.

   FreeImage_SetBackgroundColorAsLong = (FreeImage_SetBackgroundColorAsLongInt(BITMAP, BackColor) = 1)
                                             
End Function

Public Function FreeImage_SetBackgroundColorEx(ByVal BITMAP As Long, _
                                               ByVal Alpha As Byte, _
                                               ByVal Red As Byte, _
                                               ByVal Green As Byte, _
                                               ByVal Blue As Byte) As Boolean
                                              
Dim tColor As RGBQUAD

   ' This function sets the color at position (x|y) as FreeImage_SetPixelColor() does but
   ' the color value to set must be provided four different byte values, one for each
   ' color component.
                                             
   With tColor
      .Alpha = Alpha
      .Red = Red
      .Green = Green
      .Blue = Blue
   End With
   FreeImage_SetBackgroundColorEx = (FreeImage_SetBackgroundColorInt(BITMAP, tColor) = 1)

End Function

Public Function FreeImage_FIFSupportsReading(ByVal Format As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsReading = (FreeImage_FIFSupportsReadingInt(Format) = 1)

End Function

Public Function FreeImage_FIFSupportsWriting(ByVal Format As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsWriting = (FreeImage_FIFSupportsWritingInt(Format) = 1)
   
End Function

Public Function FreeImage_FIFSupportsExportType(ByVal Format As FREE_IMAGE_FORMAT, _
                                                ByVal ImageType As FREE_IMAGE_TYPE) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsExportType = (FreeImage_FIFSupportsExportTypeInt(Format, ImageType) = 1)

End Function

Public Function FreeImage_FIFSupportsExportBPP(ByVal Format As FREE_IMAGE_FORMAT, _
                                               ByVal BitsPerPixel As Long) As Boolean
   
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsExportBPP = (FreeImage_FIFSupportsExportBPPInt(Format, BitsPerPixel) = 1)
                                             
End Function

Public Function FreeImage_FIFSupportsICCProfiles(ByVal Format As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsICCProfiles = (FreeImage_FIFSupportsICCProfilesInt(Format) = 1)

End Function

Public Function FreeImage_FIFSupportsNoPixels(ByVal Format As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsNoPixels = (FreeImage_FIFSupportsNoPixelsInt(Format) = 1)

End Function

Public Function FreeImage_CloseMultiBitmap(ByVal BITMAP As Long, _
                                  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_CloseMultiBitmap = (FreeImage_CloseMultiBitmapInt(BITMAP, Flags) = 1)

End Function

Public Function FreeImage_SaveToMemory(ByVal Format As FREE_IMAGE_FORMAT, _
                                       ByVal BITMAP As Long, _
                                       ByVal Stream As Long, _
                              Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean
                              
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_SaveToMemory = (FreeImage_SaveToMemoryInt(Format, BITMAP, Stream, Flags) = 1)
  
End Function

Public Function FreeImage_AcquireMemory(ByVal Stream As Long, _
                                        ByRef DataPtr As Long, _
                                        ByRef SizeInBytes As Long) As Boolean
                                        
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_AcquireMemory = (FreeImage_AcquireMemoryInt(Stream, DataPtr, SizeInBytes) = 1)
           
End Function

Public Function FreeImage_SeekMemory(ByVal Stream As Long, _
                                     ByVal Offset As Long, _
                                     ByVal Origin As Long) As Boolean
                                     
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_SeekMemory = (FreeImage_SeekMemoryInt(Stream, Offset, Origin) = 1)

End Function

Public Function FreeImage_IsLittleEndian() As Boolean
   
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_IsLittleEndian = (FreeImage_IsLittleEndianInt() = 1)

End Function

Public Function FreeImage_LookupX11Color(ByVal Color As String, _
                                         ByRef Red As Long, _
                                         ByRef Green As Long, _
                                         ByRef Blue As Long) As Boolean
                                         
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_LookupX11Color = (FreeImage_LookupX11ColorInt(Color, Red, Green, Blue) = 1)
           
End Function

Public Function FreeImage_LookupSVGColor(ByVal Color As String, _
                                         ByRef Red As Long, _
                                         ByRef Green As Long, _
                                         ByRef Blue As Long) As Boolean
                                         
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_LookupSVGColor = (FreeImage_LookupSVGColorInt(Color, Red, Green, Blue) = 1)
         
End Function

Public Function FreeImage_PreMultiplyWithAlpha(ByVal BITMAP As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value
   
   FreeImage_PreMultiplyWithAlpha = (FreeImage_PreMultiplyWithAlphaInt(BITMAP) = 1)

End Function

Public Function FreeImage_SetThumbnail(ByVal BITMAP As Long, ByVal Thumbnail As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value
   
   FreeImage_SetThumbnail = (FreeImage_SetThumbnailInt(BITMAP, Thumbnail) = 1)

End Function


Public Function FreeImage_OpenMultiBitmap(ByVal Format As FREE_IMAGE_FORMAT, _
                                          ByVal FileName As String, _
                                 Optional ByVal CreateNew As Boolean, _
                                 Optional ByVal ReadOnly As Boolean, _
                                 Optional ByVal KeepCacheInMemory As Boolean, _
                                 Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

   FreeImage_OpenMultiBitmap = FreeImage_OpenMultiBitmapInt(Format, FileName, IIf(CreateNew, 1, 0), _
         IIf(ReadOnly And Not CreateNew, 1, 0), IIf(KeepCacheInMemory, 1, 0), Flags)

End Function

Public Sub FreeImage_UnlockPage(ByVal BITMAP As Long, ByVal PageBitmap As Long, ByVal ApplyChanges As Boolean)

Dim lApplyChanges As Long

   If (ApplyChanges) Then
      lApplyChanges = 1
   End If
   Call FreeImage_UnlockPageInt(BITMAP, PageBitmap, lApplyChanges)

End Sub

Public Function FreeImage_RotateEx(ByVal BITMAP As Long, _
                                   ByVal Angle As Double, _
                          Optional ByVal ShiftX As Double, _
                          Optional ByVal ShiftY As Double, _
                          Optional ByVal OriginX As Double, _
                          Optional ByVal OriginY As Double, _
                          Optional ByVal UseMask As Boolean) As Long

    Dim lUseMask As Long

    If UseMask Then lUseMask = 1 Else lUseMask = 0
    FreeImage_RotateEx = FreeImage_RotateExInt(BITMAP, Angle, ShiftX, ShiftY, OriginX, OriginY, lUseMask)

End Function

Public Function FreeImage_MakeThumbnail(ByVal BITMAP As Long, _
                                        ByVal MaxPixelSize As Long, _
                               Optional ByVal Convert As Boolean) As Long

Dim lConvert As Long

   If (Convert) Then
      lConvert = 1
   End If
   FreeImage_MakeThumbnail = FreeImage_MakeThumbnailInt(BITMAP, MaxPixelSize, lConvert)

End Function

Public Function FreeImage_ConvertToStandardType(ByVal BITMAP As Long, _
                                       Optional ByVal ScaleLinear As Boolean = True) As Long
                                       
   If (ScaleLinear) Then
      FreeImage_ConvertToStandardType = FreeImage_ConvertToStandardTypeInt(BITMAP, 1)
   Else
      FreeImage_ConvertToStandardType = FreeImage_ConvertToStandardTypeInt(BITMAP, 0)
   End If
   
End Function

Public Function FreeImage_ConvertToType(ByVal BITMAP As Long, _
                                        ByVal DestinationType As FREE_IMAGE_TYPE, _
                               Optional ByVal ScaleLinear As Boolean = True) As Long
                                       
   If (ScaleLinear) Then
      FreeImage_ConvertToType = FreeImage_ConvertToTypeInt(BITMAP, DestinationType, 1)
   Else
      FreeImage_ConvertToType = FreeImage_ConvertToTypeInt(BITMAP, DestinationType, 0)
   End If
   
End Function



'--------------------------------------------------------------------------------
' Color conversion helper functions
'--------------------------------------------------------------------------------

Public Function ConvertColor(ByVal Color As Long) As Long

   ' This helper function converts a VB-style color value (like vbRed), which
   ' uses the ABGR format into a RGBQUAD compatible color value, using the ARGB
   ' format, needed by FreeImage and vice versa.

   ConvertColor = ((Color And &HFF000000) Or _
                   ((Color And &HFF&) * &H10000) Or _
                   ((Color And &HFF00&)) Or _
                   ((Color And &HFF0000) \ &H10000))

End Function

Public Function ConvertOleColor(ByVal Color As OLE_COLOR) As Long

   ' This helper function converts an OLE_COLOR value (like vbButtonFace), which
   ' uses the BGR format into a RGBQUAD compatible color value, using the ARGB
   ' format, needed by FreeImage.
   
   ' This function generally ingnores the specified color's alpha value but, in
   ' contrast to ConvertColor, also has support for system colors, which have the
   ' format &H80bbggrr.
   
   ' You should not use this function to convert any color provided by FreeImage
   ' in ARGB format into a VB-style ABGR color value. Use function ConvertColor
   ' instead.

Dim lColorRef As Long

   If (OleTranslateColor(Color, 0, lColorRef) = 0) Then
      ConvertOleColor = ConvertColor(lColorRef)
   End If

End Function



'--------------------------------------------------------------------------------
' Extended functions derived from FreeImage 3 functions usually dealing
' with arrays
'--------------------------------------------------------------------------------

Public Sub FreeImage_UnloadEx(ByRef BITMAP As Long)

   ' Extended version of FreeImage_Unload, which additionally sets the
   ' passed Bitmap handle to zero after unloading.

   If (BITMAP <> 0) Then
      Call FreeImage_Unload(BITMAP)
      BITMAP = 0
   End If

End Sub


' Memory and Stream functions

Public Function FreeImage_GetFileTypeFromMemoryEx(ByRef Data As Variant, _
                                         Optional ByRef SizeInBytes As Long) As FREE_IMAGE_FORMAT

Dim hStream As Long
Dim lDataPtr As Long

   ' This function extends the FreeImage function FreeImage_GetFileTypeFromMemory()
   ' to a more VB suitable function. The parameter data of type Variant my
   ' me either an array of type Byte, Integer or Long or may contain the pointer
   ' to a memory block, what in VB is always the address of the memory block,
   ' since VB actually doesn's support native pointers.
   
   ' In case of providing the memory block as an array, the SizeInBytes may
   ' be omitted, zero or less than zero. Then, the size of the memory block
   ' is calculated correctly. When SizeInBytes is given, it is up to the caller
   ' to ensure, it is correct.
   
   ' In case of providing an address of a memory block, SizeInBytes must not
   ' be omitted.
  

   ' get both pointer and size in bytes of the memory block provided
   ' through the Variant parameter 'data'.
   lDataPtr = pGetMemoryBlockPtrFromVariant(Data, SizeInBytes)
   
   ' open the memory stream
   hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
   If (hStream) Then
      ' on success, detect image type
      FreeImage_GetFileTypeFromMemoryEx = FreeImage_GetFileTypeFromMemory(hStream)
      Call FreeImage_CloseMemory(hStream)
   Else
      FreeImage_GetFileTypeFromMemoryEx = FIF_UNKNOWN
   End If

End Function

'NOTE: modified by Tanner to support direct pointer retrieval
Public Function FreeImage_LoadFromMemoryEx(ByRef Data As Variant, _
                                  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS = 0, _
                                  Optional ByRef SizeInBytes As Long, _
                                  Optional ByRef Format As FREE_IMAGE_FORMAT = FIF_UNKNOWN, _
                                  Optional ByVal ptrToDataInstead As Long = 0) As Long

Dim hStream As Long
Dim lDataPtr As Long

   ' This function extends the FreeImage function FreeImage_LoadFromMemory()
   ' to a more VB suitable function. The parameter data of type Variant my
   ' me either an array of type Byte, Integer or Long or may contain the pointer
   ' to a memory block, what in VB is always the address of the memory block,
   ' since VB actually doesn's support native pointers.
   
   ' The parameter 'Flags' works according to the FreeImage API documentation.
   
   ' In case of providing the memory block as an array, the SizeInBytes may
   ' be omitted, zero or less than zero. Then, the size of the memory block
   ' is calculated correctly. When SizeInBytes is given, it is up to the caller
   ' to ensure, it is correct.
   
   ' In case of providing an address of a memory block, SizeInBytes must not
   ' be omitted.
   
   ' The parameter fif is an OUT parameter, that will contain the image type
   ' detected. Any values set by the caller will never be used within this
   ' function.
   

   ' get both pointer and size in bytes of the memory block provided
   ' through the Variant parameter 'data'.
   
    'EDIT BY TANNER: use the pointer directly, if provided
    If ptrToDataInstead <> 0 Then
        lDataPtr = ptrToDataInstead
    Else
        lDataPtr = pGetMemoryBlockPtrFromVariant(Data, SizeInBytes)
    End If
   
   ' open the memory stream
   hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
   If (hStream) Then
   
        ' on success, detect image type
        If Format = FIF_UNKNOWN Then
            Format = FreeImage_GetFileTypeFromMemory(hStream)
            Debug.Print "FreeImage_LoadFromMemoryEx auto-detected format " & Format
        End If
      
      If (Format <> FIF_UNKNOWN) Then
         ' load the image from memory stream only, if known image type
         FreeImage_LoadFromMemoryEx = FreeImage_LoadFromMemory(Format, hStream, Flags)
      End If
      
      ' close the memory stream when open
      Call FreeImage_CloseMemory(hStream)
   Else
        Debug.Print "Couldn't obtain hStream pointer in FreeImage_LoadFromMemoryEx; sorry!"
   End If

End Function

Public Function FreeImage_LoadFromMemoryEx_Tanner(ByVal DataPtr As Long, ByVal SizeInBytes As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef fileFormat As FREE_IMAGE_FORMAT = FIF_UNKNOWN) As Long

    Dim hStream As Long
    
    'FreeImage_LoadFromMemoryEx routinely fails without explanation, and I'm hoping to find out why!

   ' get both pointer and size in bytes of the memory block provided
   ' through the Variant parameter 'data'.
   'lDataPtr = pGetMemoryBlockPtrFromVariant(Data, SizeInBytes)
   
   ' open the memory stream
   hStream = FreeImage_OpenMemoryByPtr(DataPtr, SizeInBytes)
   If (hStream) Then
   
      ' on success, detect image type
      If fileFormat = FIF_UNKNOWN Then fileFormat = FreeImage_GetFileTypeFromMemory(hStream)
      
      If (fileFormat <> FIF_UNKNOWN) Then
         ' load the image from memory stream only, if known image type
         FreeImage_LoadFromMemoryEx_Tanner = FreeImage_LoadFromMemory(fileFormat, hStream, Flags)
      Else
        Debug.Print "Format could not be ascertained!!"
      
      End If
      
      ' close the memory stream when open
      Call FreeImage_CloseMemory(hStream)
   End If

End Function

Public Function FreeImage_SaveToMemoryEx(ByVal Format As FREE_IMAGE_FORMAT, _
                                         ByVal BITMAP As Long, _
                                         ByRef Data() As Byte, _
                                Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, _
                                Optional ByVal UnloadSource As Boolean) As Boolean

Dim hStream As Long
Dim lpData As Long
Dim lSizeInBytes As Long

   ' This function saves a FreeImage DIB into memory by using the VB Byte
   ' array Data(). It makes a deep copy of the image data and closes the
   ' memory stream opened before it returns to the caller.
   
   ' The Byte array 'Data()' must not be a fixed sized array and will be
   ' redimensioned according to the size needed to hold all the data.
   
   ' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage 3
   ' API documentation.
   
   ' The optional 'UnloadSource' parameter is for unloading the original image
   ' after it has been saved into memory. There is no need to clean up the DIB
   ' at the caller's site.
   
   ' The function returns True on success and False otherwise.
   
   
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to save a 'header-only' bitmap.")
      End If
   
      hStream = FreeImage_OpenMemory()
      If (hStream) Then
         FreeImage_SaveToMemoryEx = FreeImage_SaveToMemory(Format, BITMAP, hStream, Flags)
         
         If (FreeImage_SaveToMemoryEx) Then
            If (FreeImage_AcquireMemoryInt(hStream, lpData, lSizeInBytes)) Then
               On Error Resume Next
               ReDim Data(lSizeInBytes - 1) As Byte
               If (Err.Number = ERROR_SUCCESS) Then
                  On Error GoTo 0
                  Call CopyMemory(Data(0), ByVal lpData, lSizeInBytes)
               Else
                  On Error GoTo 0
                  FreeImage_SaveToMemoryEx = False
               End If
            Else
               FreeImage_SaveToMemoryEx = False
            End If
         
         Else
         
            
         
         End If
         
         
         Call FreeImage_CloseMemory(hStream)
      Else
         FreeImage_SaveToMemoryEx = False
      End If
      
      If (UnloadSource) Then
         Call FreeImage_Unload(BITMAP)
      End If
   End If

End Function

Public Function FreeImage_SaveToMemoryEx2(ByVal Format As FREE_IMAGE_FORMAT, _
                                          ByVal BITMAP As Long, _
                                          ByRef Data() As Byte, _
                                          ByRef Stream As Long, _
                                 Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, _
                                 Optional ByVal UnloadSource As Boolean) As Boolean

   ' This function saves a FreeImage DIB into memory by using the VB Byte
   ' array Data(). It does not makes a deep copy of the image data, but uses
   ' the function 'FreeImage_AcquireMemoryEx' to wrap the array 'Data()'
   ' around the memory block pointed to by the result of the
   ' 'FreeImage_AcquireMemory' function.
   
   ' The Byte array 'Data()' must not be a fixed sized array and will be
   ' redimensioned according to the size needed to hold all the data.
   
   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the 'FreeImage_DestroyLockedArray' function.
   
   ' The parameter 'stream' is an IN/OUT parameter, tracking the memory
   ' stream, the VB array 'Data()' is based on. This parameter may contain
   ' an already opened FreeImage memory stream when the function is called and
   ' contains a valid memory stream when the function returns in each case.
   ' After all, it is up to the caller to close that memory stream correctly.
   ' The array 'Data()' will no longer be valid and accessable after the stream
   ' has been closed, so it should only be closed after the passed byte array
   ' variable either goes out of the caller's scope or is redimensioned.
   
   ' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage 3
   ' API documentation.
   
   ' The optional 'UnloadSource' parameter is for unloading the original image
   ' after it has been saved to memory. There is no need to clean up the DIB
   ' at the caller's site.
   
   ' The function returns True on success and False otherwise.

   
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to save a 'header-only' bitmap.")
      End If
   
      If (Stream = 0) Then
         Stream = FreeImage_OpenMemory()
      End If
      If (Stream) Then
         FreeImage_SaveToMemoryEx2 = FreeImage_SaveToMemory(Format, BITMAP, Stream, Flags)
         If (FreeImage_SaveToMemoryEx2) Then
            FreeImage_SaveToMemoryEx2 = FreeImage_AcquireMemoryEx(Stream, Data)
         End If
         
         ' do not close the memory stream, since the returned array data()
         ' points to the stream's data
         ' the caller must close the stream after he is done
         ' with the array
      Else
         FreeImage_SaveToMemoryEx2 = False
      End If
      
      If (UnloadSource) Then
         Call FreeImage_Unload(BITMAP)
      End If
   End If

End Function

Public Function FreeImage_AcquireMemoryEx(ByVal Stream As Long, _
                                          ByRef Data() As Byte, _
                                 Optional ByRef SizeInBytes As Long) As Boolean
                                          
Dim lpData As Long
Dim tSA As SAVEARRAY1D
Dim lpSA As Long

   ' This function wraps the byte array Data() around acquired memory
   ' of the memory stream specified by then stream parameter. The adjusted
   ' array then points directly to the stream's data pointer and so
   ' provides full read and write access.
   
   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the 'FreeImage_DestroyLockedArray' function.


   If (Stream) Then
      If (FreeImage_AcquireMemoryInt(Stream, lpData, SizeInBytes)) Then
         With tSA
            .cbElements = 1                           ' one element is one byte
            .cDims = 1                                ' the array has only 1 dimension
            .cElements = SizeInBytes                  ' the number of elements in the array is
                                                      ' the size in bytes of the memory block
            .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues,
                                                      ' so the array can not be modified in size
                                                      ' or erased; according to Matthew Curland never
                                                      ' use FIXEDSIZE alone
            .pvData = lpData                          ' let the array point to the memory block
                                                      ' received by FreeImage_AcquireMemory
         End With
         
         lpSA = pDeref(VarPtrArray(Data))
         If (lpSA = 0) Then
            ' allocate memory for an array descriptor
            Call SafeArrayAllocDescriptor(1, lpSA)
            Call CopyMemory(ByVal VarPtrArray(Data), lpSA, 4)
         Else
            Call SafeArrayDestroyData(lpSA)
         End If
         Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
      Else
         FreeImage_AcquireMemoryEx = False
      End If
   Else
      FreeImage_AcquireMemoryEx = False
   End If

End Function

Public Function FreeImage_ReadMemoryEx(ByRef Buffer As Variant, _
                                       ByVal Stream As Long, _
                              Optional ByRef Count As Long, _
                              Optional ByRef Size As Long) As Long
                                       
Dim lBufferPtr As Long
Dim lSizeInBytes As Long
Dim lSize As Long
Dim lCount As Long

   ' This function is a wrapper for 'FreeImage_ReadMemory()' using VB style
   ' arrays instead of a void pointer.
   
   ' The variant parameter 'Buffer' may be a Byte, Integer or Long array or
   ' may contain a pointer to a memory block (the memory block's address).
   
   ' In the latter case, this function behaves exactly
   ' like 'FreeImage_ReadMemory()'. Then, 'Count' and 'Size' must be valid
   ' upon entry.
   
   ' If 'Buffer' is an initialized (dimensioned) array, 'Count' and 'Size' may
   ' be omitted. Then, the array's layout is used to determine 'Count'
   ' and 'Size'. In that case, any provided value in 'Count' and 'Size' upon
   ' entry will override these calculated values as long as they are not
   ' exceeding the size of the array in 'Buffer'.
   
   ' If 'Buffer' is an uninitialized (not yet dimensioned) array of any valid
   ' type (Byte, Integer or Long) and, at least 'Count' is specified, the
   ' array in 'Buffer' is redimensioned by this function. If 'Buffer' is a
   ' fixed-size or otherwise locked array, a runtime error (10) occurs.
   ' If 'Size' is omitted, the array's element size is assumed to be the
   ' desired value.
   
   ' As FreeImage's function 'FreeImage_ReadMemory()', this function returns
   ' the number of items actually read.
   
   ' Example: (very freaky...)
   '
   ' Dim alLongBuffer() As Long
   ' Dim lRet as Long
   '
   '    ' now reading 303 integers (2 byte) into an array of Longs
   '    lRet = FreeImage_ReadMemoryEx(alLongBuffer, lMyStream, 303, 2)
   '
   '    ' now, lRet contains 303 and UBound(alLongBuffer) is 151 since
   '    ' we need at least 152 Longs (0..151) to store (303 * 2) = 606 bytes
   '    ' so, the higest two bytes of alLongBuffer(151) contain only unset
   '    ' bits. Got it?
   
   ' Remark: This function's parameter order differs from FreeImage's
   '         original funtion 'FreeImage_ReadMemory()'!
                                       
   If (VarType(Buffer) And vbArray) Then
      ' get both pointer and size in bytes of the memory block provided
      ' through the Variant parameter 'Buffer'.
      lBufferPtr = pGetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
      If (lBufferPtr = 0) Then
         ' array is not initialized
         If (Count > 0) Then
            ' only if we have a 'Count' value, redim the array
            If (Size <= 0) Then
               ' if 'Size' is omitted, use array's element size
               Size = lSize
            End If
            
            Select Case lSize
            
            Case 2
               ' Remark: -Int(-a) == ceil(a); a > 0
               ReDim Buffer(-Int(-Count * Size / 2) - 1) As Integer
            
            Case 4
               ' Remark: -Int(-a) == ceil(a); a > 0
               ReDim Buffer(-Int(-Count * Size / 4) - 1) As Long
            
            Case Else
               ReDim Buffer((Count * Size) - 1) As Byte
            
            End Select
            lBufferPtr = pGetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
         End If
      End If
      If (lBufferPtr) Then
         lCount = lSizeInBytes / lSize
         If (Size <= 0) Then
            ' use array's natural value for 'Size' when
            ' omitted
            Size = lSize
         End If
         If (Count <= 0) Then
            ' use array's natural value for 'Count' when
            ' omitted
            Count = lCount
         End If
         If ((Size * Count) > (lSize * lCount)) Then
            If (Size = lSize) Then
               Count = lCount
            Else
               ' Remark: -Fix(-a) == floor(a); a > 0
               Count = -Fix(-lSizeInBytes / Size)
               If (Count = 0) Then
                  Size = lSize
                  Count = lCount
               End If
            End If
         End If
         FreeImage_ReadMemoryEx = FreeImage_ReadMemory(lBufferPtr, Size, Count, Stream)
      End If
   
   ElseIf (VarType(Buffer) = vbLong) Then
      ' if Buffer is a Long, it specifies the address of a memory block
      ' then, we do not know anything about its size, so assume that 'Size'
      ' and 'Count' are correct and forward these directly to the FreeImage
      ' call.
      FreeImage_ReadMemoryEx = FreeImage_ReadMemory(CLng(Buffer), Size, Count, Stream)
   
   End If

End Function

Public Function FreeImage_WriteMemoryEx(ByRef Buffer As Variant, _
                                        ByVal Stream As Long, _
                               Optional ByRef Count As Long, _
                               Optional ByRef Size As Long) As Long
                                       
Dim lBufferPtr As Long
Dim lSizeInBytes As Long
Dim lSize As Long
Dim lCount As Long

   ' This function is a wrapper for 'FreeImage_WriteMemory()' using VB style
   ' arrays instead of a void pointer.
   
   ' The variant parameter 'Buffer' may be a Byte, Integer or Long array or
   ' may contain a pointer to a memory block (the memory block's address).
   
   ' In the latter case, this function behaves exactly
   ' like 'FreeImage_WriteMemory()'. Then, 'Count' and 'Size' must be valid
   ' upon entry.
   
   ' If 'Buffer' is an initialized (dimensioned) array, 'Count' and 'Size' may
   ' be omitted. Then, the array's layout is used to determine 'Count'
   ' and 'Size'. In that case, any provided value in 'Count' and 'Size' upon
   ' entry will override these calculated values as long as they are not
   ' exceeding the size of the array in 'Buffer'.
   
   ' If 'Buffer' is an uninitialized (not yet dimensioned) array of any
   ' type, the function will do nothing an returns 0.
   
   ' Remark: This function's parameter order differs from FreeImage's
   '         original funtion 'FreeImage_ReadMemory()'!

   If (VarType(Buffer) And vbArray) Then
      ' get both pointer and size in bytes of the memory block provided
      ' through the Variant parameter 'Buffer'.
      lBufferPtr = pGetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
      If (lBufferPtr) Then
         lCount = lSizeInBytes / lSize
         If (Size <= 0) Then
            ' use array's natural value for 'Size' when
            ' omitted
            Size = lSize
         End If
         If (Count <= 0) Then
            ' use array's natural value for 'Count' when
            ' omitted
            Count = lCount
         End If
         If ((Size * Count) > (lSize * lCount)) Then
            If (Size = lSize) Then
               Count = lCount
            Else
               ' Remark: -Fix(-a) == floor(a); a > 0
               Count = -Fix(-lSizeInBytes / Size)
               If (Count = 0) Then
                  Size = lSize
                  Count = lCount
               End If
            End If
         End If
         FreeImage_WriteMemoryEx = FreeImage_WriteMemory(lBufferPtr, Size, Count, Stream)
      End If
   
   ElseIf (VarType(Buffer) = vbLong) Then
      ' if Buffer is a Long, it specifies the address of a memory block
      ' then, we do not know anything about its size, so assume that 'Size'
      ' and 'Count' are correct and forward these directly to the FreeImage
      ' call.
      FreeImage_WriteMemoryEx = FreeImage_WriteMemory(CLng(Buffer), Size, Count, Stream)
   
   End If

End Function


Public Function FreeImage_UnsignedLong(ByVal Value As Long) As Variant

   ' This function converts a signed long (VB's Long data type) into
   ' an unsigned long (not really supported by VB).
   
   ' Basically, this function checks, whether the positive range of
   ' a signed long is sufficient to hold the value (indeed, it checks
   ' the value since the range is obviously constant). If yes,
   ' it returns a Variant with subtype Long ('Variant/Long' in VB's
   ' watch window). In this case, the function did not make any real
   ' changes at all. If not, the value is stored in a Currency variable,
   ' which is able to store the whole range of an unsigned long. Then,
   ' the function returns a Variant with subtype Currency
   ' ('Variant/Currency' in VB's watch window).
   
   If (Value < 0) Then
      Dim curTemp As Currency
      Call CopyMemory(curTemp, Value, 4)
      FreeImage_UnsignedLong = curTemp * 10000
   Else
      FreeImage_UnsignedLong = Value
   End If

End Function

Public Function FreeImage_UnsignedShort(ByVal Value As Integer) As Variant

   ' This function converts a signed short (VB's Integer data type) into
   ' an unsigned short (not really supported by VB).
   
   ' Basically, this function checks, whether the positive range of
   ' a signed short is sufficient to hold the value (indeed, it checks
   ' the value since the range is obviously constant). If yes,
   ' it returns a Variant with subtype Integer ('Variant/Integer' in VB's
   ' watch window). In this case, the function did not make any real
   ' changes at all. If not, the value is stored in a Long variable,
   ' which is able to store the whole range of an unsigned short. Then,
   ' the function returns a Variant with subtype Long
   ' ('Variant/Long' in VB's watch window).
   
   If (Value < 0) Then
      Dim lTemp As Long
      Call CopyMemory(lTemp, Value, 2)
      FreeImage_UnsignedShort = lTemp
   Else
      FreeImage_UnsignedShort = Value
   End If

End Function



'--------------------------------------------------------------------------------
' Derived and hopefully useful functions
'--------------------------------------------------------------------------------

' Plugin and filename functions

Public Function FreeImage_IsExtensionValidForFIF(ByVal Format As FREE_IMAGE_FORMAT, _
                                                 ByVal Extension As String, _
                                        Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
   
   ' This function tests, whether a given filename extension is valid
   ' for a certain image format (fif).
   
   FreeImage_IsExtensionValidForFIF = (InStr(1, _
                                             FreeImage_GetFIFExtensionList(Format) & ",", _
                                             Extension & ",", _
                                             Compare) > 0)

End Function

Public Function FreeImage_IsFilenameValidForFIF(ByVal Format As FREE_IMAGE_FORMAT, _
                                                ByVal FileName As String, _
                                       Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
                                                
Dim strExtension As String
Dim i As Long

   ' This function tests, whether a given complete filename is valid
   ' for a certain image format (fif).

   i = InStrRev(FileName, ".")
   If (i > 0) Then
      strExtension = Mid$(FileName, i + 1)
      FreeImage_IsFilenameValidForFIF = (InStr(1, _
                                               FreeImage_GetFIFExtensionList(Format) & ",", _
                                               strExtension & ",", _
                                               Compare) > 0)
   End If
   
End Function

Public Function FreeImage_GetPrimaryExtensionFromFIF(ByVal Format As FREE_IMAGE_FORMAT) As String

Dim strExtensionList As String
Dim i As Long

   ' This function returns the primary (main or most commonly used?) extension
   ' of a certain image format (fif). This is done by returning the first of
   ' all possible extensions returned by FreeImage_GetFIFExtensionList(). That
   ' assumes, that the plugin returns the extensions in ordered form. If not,
   ' in most cases it is even enough, to receive any extension.
   
   ' This function is primarily used by the function 'SavePictureEx'.

   strExtensionList = FreeImage_GetFIFExtensionList(Format)
   i = InStr(strExtensionList, ",")
   If (i > 0) Then
      FreeImage_GetPrimaryExtensionFromFIF = Left$(strExtensionList, i - 1)
   Else
      FreeImage_GetPrimaryExtensionFromFIF = strExtensionList
   End If

End Function

' Bitmap resolution functions

Public Function FreeImage_GetResolutionX(ByVal BITMAP As Long) As Long

   ' This function gets a DIB's resolution in X-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.
   
   FreeImage_GetResolutionX = Int(0.5 + 0.0254 * FreeImage_GetDotsPerMeterX(BITMAP))

End Function

Public Sub FreeImage_SetResolutionX(ByVal BITMAP As Long, ByVal Resolution As Long)

   ' This function sets a DIB's resolution in X-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   Call FreeImage_SetDotsPerMeterX(BITMAP, Int(Resolution / 0.0254 + 0.5))

End Sub

Public Function FreeImage_GetResolutionY(ByVal BITMAP As Long) As Long

   ' This function gets a DIB's resolution in Y-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   FreeImage_GetResolutionY = Int(0.5 + 0.0254 * FreeImage_GetDotsPerMeterY(BITMAP))

End Function

Public Sub FreeImage_SetResolutionY(ByVal BITMAP As Long, ByVal Resolution As Long)

   ' This function sets a DIB's resolution in Y-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   Call FreeImage_SetDotsPerMeterY(BITMAP, Int(Resolution / 0.0254 + 0.5))

End Sub

' ICC Color Profile functions

Public Function FreeImage_GetICCProfile(ByVal BITMAP As Long) As FIICCPROFILE

   ' This function is a wrapper for the FreeImage_GetICCProfile() function, returning
   ' a real FIICCPROFILE structure.
   
   ' Since the original FreeImage function returns a pointer to the FIICCPROFILE
   ' structure (FIICCPROFILE *), as with string returning functions, this wrapper is
   ' needed as VB can't declare a function returning a pointer to anything. So,
   ' analogous to string returning functions, FreeImage_GetICCProfile() is declared
   ' private as FreeImage_GetICCProfileInt() and made publicly available with this
   ' wrapper function.

   Call CopyMemory(FreeImage_GetICCProfile, _
                   ByVal FreeImage_GetICCProfileInt(BITMAP), _
                   LenB(FreeImage_GetICCProfile))

End Function

Public Function FreeImage_GetICCProfileColorModel(ByVal BITMAP As Long) As FREE_IMAGE_ICC_COLOR_MODEL

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' the color model in which the ICC color profile data is in, if there is actually
   ' a ICC color profile available for the Bitmap specified.
   
   ' If there is NO color profile along with that bitmap, this function returns the color
   ' model that should (or must) be used for any color profile data to be assigned to the
   ' Bitmap. That depends on the bitmap's color type.

   If (FreeImage_HasICCProfile(BITMAP)) Then
      FreeImage_GetICCProfileColorModel = (pDeref(FreeImage_GetICCProfileInt(BITMAP)) _
            And FREE_IMAGE_ICC_COLOR_MODEL_MASK)
   Else
      ' use FreeImage_GetColorType() to determine, whether this is a CMYK bitmap or not
      If (FreeImage_GetColorType(BITMAP) = FIC_CMYK) Then
         FreeImage_GetICCProfileColorModel = FIICC_COLOR_MODEL_CMYK
      Else
         FreeImage_GetICCProfileColorModel = FIICC_COLOR_MODEL_RGB
      End If
   End If

End Function

Public Function FreeImage_GetICCProfileSize(ByVal BITMAP As Long) As Long

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' only the size in bytes of the ICC profile data for the Bitmap specified or zero,
   ' if there is no ICC profile data for the Bitmap.

   FreeImage_GetICCProfileSize = pDeref(FreeImage_GetICCProfileInt(BITMAP) + 4)

End Function

Public Function FreeImage_GetICCProfileDataPointer(ByVal BITMAP As Long) As Long

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' only the pointer (the address) of the ICC profile data for the Bitmap specified,
   ' or zero if there is no ICC profile data for the Bitmap.

   FreeImage_GetICCProfileDataPointer = pDeref(FreeImage_GetICCProfileInt(BITMAP) + 8)

End Function

Public Function FreeImage_HasICCProfile(ByVal BITMAP As Long) As Boolean

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' True, if there is an ICC color profile available for the Bitmap specified or
   ' returns False otherwise.

   FreeImage_HasICCProfile = (FreeImage_GetICCProfileSize(BITMAP) <> 0)

End Function

' Image color depth conversion wrapper

Public Function FreeImage_GetPaletteEx(ByVal BITMAP As Long) As RGBQUAD()

Dim tSA As SAVEARRAY1D
Dim lpSA As Long

   ' This function returns a VB style array of type RGBQUAD, containing
   ' the palette data of the Bitmap. This array provides read and write access
   ' to the actual palette data provided by FreeImage. This is done by
   ' creating a VB array with an own SAFEARRAY descriptor making the
   ' array point to the palette pointer returned by FreeImage_GetPalette().
   
   ' This makes you use code like you would in C/C++:
   
   ' // this code assumes there is a bitmap loaded and
   ' // present in a variable called ‘dib’
   ' if(FreeImage_GetBPP(Bitmap) == 8) {
   '   // Build a greyscale palette
   '   RGBQUAD *pal = FreeImage_GetPalette(Bitmap);
   '   for (int i = 0; i < 256; i++) {
   '     pal[i].rgbRed = i;
   '     pal[i].rgbGreen = i;
   '     pal[i].rgbBlue = i;
   '   }
   
   ' As in C/C++ the array is only valid while the DIB is loaded and the
   ' palette data remains where the pointer returned by FreeImage_GetPalette
   ' has pointed to when this function was called. So, a good thing would
   ' be, not to keep the returned array in scope over the lifetime of the
   ' Bitmap. Best practise is, to use this function within another routine and
   ' assign the return value (the array) to a local variable only. As soon
   ' as this local variable goes out of scope (when the calling function
   ' returns to it's caller), the array and the descriptor is automatically
   ' cleaned up by VB.
   
   ' This function does not make a deep copy of the palette data, but only
   ' wraps a VB array around the FreeImage palette data. So, it can be called
   ' frequently "on demand" or somewhat "in place" without a significant
   ' performance loss.
   
   ' To learn more about this technique I recommend reading chapter 2 (Leveraging
   ' Arrays) of Matthew Curland's book "Advanced Visual Basic 6"
   
   ' The parameter 'Bitmap' works according to the FreeImage 3 API documentation.
   
   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the FreeImage_DestroyLockedArrayRGBQUAD() function.
   
   
   If (BITMAP) Then
      
      ' create a proper SAVEARRAY descriptor
      With tSA
         .cbElements = 4                              ' size in bytes of RGBQUAD structure
         .cDims = 1                                   ' the array has only 1 dimension
         .cElements = FreeImage_GetColorsUsed(BITMAP) ' the number of elements in the array is
                                                      ' the number of used colors in the Bitmap
         .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE     ' need AUTO and FIXEDSIZE for safety issues,
                                                      ' so the array can not be modified in size
                                                      ' or erased; according to Matthew Curland never
                                                      ' use FIXEDSIZE alone
         .pvData = FreeImage_GetPalette(BITMAP)       ' let the array point to the memory block, the
                                                      ' FreeImage palette pointer points to
      End With
      
      ' allocate memory for an array descriptor
      ' we cannot use the memory block used by tSA, since it is
      ' released when tSA goes out of scope, leaving us with an
      ' array with zeroed descriptor
      ' we use nearly the same method that VB uses, so VB is able
      ' to cleanup the array variable and it's descriptor; the
      ' array data is not touched when cleaning up, since both AUTO
      ' and FIXEDSIZE flags are set
      Call SafeArrayAllocDescriptor(1, lpSA)
      
      ' copy our own array descriptor over the descriptor allocated
      ' by SafeArrayAllocDescriptor; lpSA is a pointer to that memory
      ' location
      Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
      
      ' the implicit variable named as the function is an array variable in VB
      ' make it point to the allocated array descriptor
      Call CopyMemory(ByVal VarPtrArray(FreeImage_GetPaletteEx), lpSA, 4)
   End If

End Function

Public Function FreeImage_IsGreyscaleImage(ByVal BITMAP As Long) As Boolean

Dim atRGB() As RGBQUAD
Dim i As Long

   ' This function returns a boolean value that is true, if the DIB is actually
   ' a greyscale image. Here, the only test condition is, that each palette
   ' entry must be a grey value, what means that each color component has the
   ' same value (red = green = blue).
   
   ' The FreeImage libraray doesn't offer a function to determine if a DIB is
   ' greyscale. The only thing you can do is to use the 'FreeImage_GetColorType'
   ' function, that returns either FIC_MINISWHITE or FIC_MINISBLACK for
   ' greyscale images. However, a DIB needs to have a ordered greyscale palette
   ' (linear ramp or inverse linear ramp) to be judged as FIC_MINISWHITE or
   ' FIC_MINISBLACK. DIB's with an unordered palette that are actually (visually)
   ' greyscale, are said to be (color-)palletized. That's also true for any 4 bpp
   ' image, since it will never have a palette that satifies the tests done
   ' in the 'FreeImage_GetColorType' function.
   
   ' So, there is a chance to omit some color depth conversions, when displaying
   ' an image in greyscale fashion. Maybe the problem will be solved in the
   ' FreeImage library one day.

   Select Case FreeImage_GetBPP(BITMAP)
   
   Case 1, 4, 8
      atRGB = FreeImage_GetPaletteEx(BITMAP)
      FreeImage_IsGreyscaleImage = True
      For i = 0 To UBound(atRGB)
         With atRGB(i)
            If ((.Red <> .Green) Or (.Red <> .Blue)) Then
               FreeImage_IsGreyscaleImage = False
               Exit For
            End If
         End With
      Next i
   
   End Select

End Function

Public Function FreeImage_ConvertColorDepth(ByVal BITMAP As Long, _
                                            ByVal Conversion As FREE_IMAGE_CONVERSION_FLAGS, _
                                   Optional ByVal UnloadSource As Boolean, _
                                   Optional ByVal Threshold As Byte = 128, _
                                   Optional ByVal DitherMethod As FREE_IMAGE_DITHER = FID_FS, _
                                   Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT) As Long
                                            
Dim hDIBNew As Long
Dim hDIBTemp As Long
Dim lBPP As Long
Dim bForceLinearRamp As Boolean
'Dim lpReservePalette As Long
'Dim bAdjustReservePaletteSize As Boolean

   ' This function is an easy-to-use wrapper for color depth conversion, intended
   ' to work around some tweaks in the FreeImage library.
   
   ' The parameters 'Threshold' and 'eDitherMode' control how thresholding or
   ' dithering are performed. The 'QuantizeMethod' parameter determines, what
   ' quantization algorithm will be used when converting to 8 bit color images.
   
   ' The 'Conversion' parameter, which can contain a single value or an OR'ed
   ' combination of some of the FREE_IMAGE_CONVERSION_FLAGS enumeration values,
   ' determines the desired output image format.
   
   ' The optional 'UnloadSource' parameter is for unloading the original image, so
   ' you can "change" an image with this function rather than getting a new DIB
   ' pointer. There is no more need for a second DIB variable at the caller's site.
   
   bForceLinearRamp = ((Conversion And FICF_REORDER_GREYSCALE_PALETTE) = 0)
   lBPP = FreeImage_GetBPP(BITMAP)

   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to convert a 'header-only' bitmap.")
      End If
   
      Select Case (Conversion And (Not FICF_REORDER_GREYSCALE_PALETTE))
      
      Case FICF_MONOCHROME_THRESHOLD
         If (lBPP > 1) Then
            hDIBNew = FreeImage_Threshold(BITMAP, Threshold)
         End If

      Case FICF_MONOCHROME_DITHER
         If (lBPP > 1) Then
            hDIBNew = FreeImage_Dither(BITMAP, DitherMethod)
         End If
      
      Case FICF_GREYSCALE_4BPP
         If (lBPP <> 4) Then
            ' If the color depth is 1 bpp and the we don't have a linear ramp palette
            ' the bitmap is first converted to an 8 bpp greyscale bitmap with a linear
            ' ramp palette and then to 4 bpp.
            If ((lBPP = 1) And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE)) Then
               hDIBTemp = BITMAP
               BITMAP = FreeImage_ConvertToGreyscale(BITMAP)
               Call FreeImage_Unload(hDIBTemp)
            End If
            hDIBNew = FreeImage_ConvertTo4Bits(BITMAP)
         Else
            ' The bitmap is already 4 bpp but may not have a linear ramp.
            ' If we force a linear ramp the bitmap is converted to 8 bpp with a linear ramp
            ' and then back to 4 bpp.
            If (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(BITMAP))) Or _
                (bForceLinearRamp And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE))) Then
               hDIBTemp = FreeImage_ConvertToGreyscale(BITMAP)
               hDIBNew = FreeImage_ConvertTo4Bits(hDIBTemp)
               Call FreeImage_Unload(hDIBTemp)
            End If
         End If
            
      Case FICF_GREYSCALE_8BPP
         ' Convert, if the bitmap is not at 8 bpp or does not have a linear ramp palette.
         If ((lBPP <> 8) Or (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(BITMAP))) Or _
                             (bForceLinearRamp And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE)))) Then
            hDIBNew = FreeImage_ConvertToGreyscale(BITMAP)
         End If
         
      Case FICF_PALLETISED_8BPP
         ' note, that the FreeImage library only quantizes 24 bit images
         ' do not convert any 8 bit images
         If (lBPP <> 8) Then
            ' images with a color depth of 24 bits can directly be
            ' converted with the FreeImage_ColorQuantize function;
            ' other images need to be converted to 24 bits first
            If (lBPP = 24) Then
               hDIBNew = FreeImage_ColorQuantize(BITMAP, QuantizeMethod)
            Else
               hDIBTemp = FreeImage_ConvertTo24Bits(BITMAP)
               hDIBNew = FreeImage_ColorQuantize(hDIBTemp, QuantizeMethod)
               Call FreeImage_Unload(hDIBTemp)
            End If
         End If
         
      Case FICF_RGB_15BPP
         If (lBPP <> 15) Then
            hDIBNew = FreeImage_ConvertTo16Bits555(BITMAP)
         End If
      
      Case FICF_RGB_16BPP
         If (lBPP <> 16) Then
            hDIBNew = FreeImage_ConvertTo16Bits565(BITMAP)
         End If
      
      Case FICF_RGB_24BPP
         If (lBPP <> 24) Then
            hDIBNew = FreeImage_ConvertTo24Bits(BITMAP)
         End If
      
      Case FICF_RGB_32BPP
         If (lBPP <> 32) Then
            hDIBNew = FreeImage_ConvertTo32Bits(BITMAP)
         End If
      
      End Select
      
      If (hDIBNew) Then
         FreeImage_ConvertColorDepth = hDIBNew
         If (UnloadSource) Then
            Call FreeImage_Unload(BITMAP)
         End If
      Else
         FreeImage_ConvertColorDepth = BITMAP
      End If
   
   End If

End Function

Public Function FreeImage_ColorQuantizeEx(ByVal BITMAP As Long, _
                                 Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, _
                                 Optional ByVal UnloadSource As Boolean, _
                                 Optional ByVal PaletteSize As Long = 256, _
                                 Optional ByVal ReserveSize As Long, _
                                 Optional ByRef ReservePalette As Variant = Null) As Long
  
Dim hTmp As Long
Dim lpPalette As Long
Dim lBlockSize As Long
Dim lElementSize As Long

   ' This function is a more VB-friendly wrapper around FreeImage_ColorQuantizeEx,
   ' which lets you specify the ReservePalette to be used not only as a pointer, but
   ' also as a real VB-style array of type Long, where each Long item takes a color
   ' in ARGB format (&HAARRGGBB). The native FreeImage function FreeImage_ColorQuantizeEx
   ' is declared private and named FreeImage_ColorQuantizeExInt and so hidden from the
   ' world outside the wrapper.
   
   ' In contrast to the FreeImage API documentation, ReservePalette is of type Variant
   ' and may either be a pointer to palette data (pointer to an array of type RGBQUAD
   ' == VarPtr(atMyPalette(0)) in VB) or an array of type Long, which then must contain
   ' the palette data in ARGB format. You can receive palette data as an array Longs
   ' from function FreeImage_GetPaletteExLong.
   ' Although ReservePalette is of type Variant, arrays of type RGBQUAD can not be
   ' passed, as long as RGBQUAD is not declared as a public type in a public object
   ' module. So, when dealing with RGBQUAD arrays, you are stuck on VarPtr or may
   ' use function FreeImage_GetPalettePtr, which is a more meaningfully named
   ' convenience wrapper around VarPtr.
   
   ' The optional 'UnloadSource' parameter is for unloading the original image, so
   ' you can "change" an image with this function rather than getting a new DIB
   ' pointer. There is no more need for a second DIB variable at the caller's site.
   
   ' All other parameters work according to the FreeImage API documentation.
   
   ' Note: Currently, any provided ReservePalette is only used, if quantize is
   '       FIQ_NNQUANT. This seems to be either a bug or an undocumented
   '       limitation of the FreeImage library (up to version 3.11.0).

   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to quantize a 'header-only' bitmap.")
      End If
      
      If (FreeImage_GetBPP(BITMAP) <> 24) Then
         hTmp = BITMAP
         BITMAP = FreeImage_ConvertTo24Bits(BITMAP)
         If (UnloadSource) Then
            Call FreeImage_Unload(hTmp)
         End If
         UnloadSource = True
      End If
      
      ' adjust PaletteSize
      If (PaletteSize < 2) Then
         PaletteSize = 2
      ElseIf (PaletteSize > 256) Then
         PaletteSize = 256
      End If
      
      lpPalette = pGetMemoryBlockPtrFromVariant(ReservePalette, lBlockSize, lElementSize)
      FreeImage_ColorQuantizeEx = FreeImage_ColorQuantizeExInt(BITMAP, QuantizeMethod, _
            PaletteSize, ReserveSize, lpPalette)
      
      If (UnloadSource) Then
         Call FreeImage_Unload(BITMAP)
      End If
   End If

End Function

' Image Rescale wrapper functions

Public Function FreeImage_RescaleEx(ByVal BITMAP As Long, _
                           Optional ByVal Width As Variant, _
                           Optional ByVal Height As Variant, _
                           Optional ByVal IsPercentValue As Boolean, _
                           Optional ByVal UnloadSource As Boolean, _
                           Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, _
                           Optional ByVal ForceCloneCreation As Boolean) As Long
                     
Dim lNewWidth As Long
Dim lNewHeight As Long
Dim hDIBNew As Long

   ' This function is a easy-to-use wrapper for rescaling an image with the
   ' FreeImage library. It returns a pointer to a new rescaled DIB provided
   ' by FreeImage.
   
   ' The parameters 'Width', 'Height' and 'IsPercentValue' control
   ' the size of the new image. Here, the function tries to fake something like
   ' overloading known from Java. It depends on the parameter's data type passed
   ' through the Variant, how the provided values for width and height are
   ' actually interpreted. The following rules apply:
   
   ' In general, non integer values are either interpreted as percent values or
   ' factors, the original image size will be multiplied with. The 'IsPercentValue'
   ' parameter controls whether the values are percent values or factors. Integer
   ' values are always considered to be the direct new image size, not depending on
   ' the original image size. In that case, the 'IsPercentValue' parameter has no
   ' effect. If one of the parameters is omitted, the image will not be resized in
   ' that direction (either in width or height) and keeps it's original size. It is
   ' possible to omit both, but that makes actually no sense.
   
   ' The following table shows some of possible data type and value combinations
   ' that might by used with that function: (assume an original image sized 100x100 px)
   
   ' Parameter         |  Values |  Values |  Values |  Values |     Values |
   ' ----------------------------------------------------------------------
   ' Width             |    75.0 |    0.85 |     200 |     120 |      400.0 |
   ' Height            |   120.0 |     1.3 |     230 |       - |      400.0 |
   ' IsPercentValue    |    True |   False |    d.c. |    d.c. |      False | <- wrong option?
   ' ----------------------------------------------------------------------
   ' Result Size       |  75x120 |  85x130 | 200x230 | 120x100 |40000x40000 |
   ' Remarks           | percent |  factor |  direct |         |maybe not   |
   '                                                           |what you    |
   '                                                           |wanted,     |
   '                                                           |right?      |
   
   ' The optional 'UnloadSource' parameter is for unloading the original image, so
   ' you can "change" an image with this function rather than getting a new DIB
   ' pointer. There is no more need for a second DIB variable at the caller's site.
   
   ' As of version 2.0 of the FreeImage VB wrapper, this function and all it's derived
   ' functions like FreeImage_RescaleByPixel() or FreeImage_RescaleByPercent(), do NOT
   ' return a clone of the image, if the new size desired is the same as the source
   ' image's size. That behaviour can be forced by setting the new parameter
   ' 'ForceCloneCreation' to True. Then, an image is also rescaled (and so
   ' effectively cloned), if the new width and height is exactly the same as the source
   ' image's width and height.
   
   ' Since this diversity may be confusing to VB developers, this function is also
   ' callable through three different functions called 'FreeImage_RescaleByPixel',
   ' 'FreeImage_RescaleByPercent' and 'FreeImage_RescaleByFactor'.
   
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to rescale a 'header-only' bitmap.")
      End If
   
      If (Not IsMissing(Width)) Then
         Select Case VarType(Width)
         
         Case vbDouble, vbSingle, vbDecimal, vbCurrency
            lNewWidth = FreeImage_GetWidth(BITMAP) * Width
            If (IsPercentValue) Then
               lNewWidth = lNewWidth / 100
            End If
         
         Case Else
            lNewWidth = Width
         
         End Select
      End If
      
      If (Not IsMissing(Height)) Then
         Select Case VarType(Height)
         
         Case vbDouble, vbSingle, vbDecimal
            lNewHeight = FreeImage_GetHeight(BITMAP) * Height
            If (IsPercentValue) Then
               lNewHeight = lNewHeight / 100
            End If
         
         Case Else
            lNewHeight = Height
         
         End Select
      End If
      
      If ((lNewWidth > 0) And (lNewHeight > 0)) Then
         If (ForceCloneCreation) Then
            hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
         
         ElseIf ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or _
                 (lNewHeight <> FreeImage_GetHeight(BITMAP))) Then
            hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
         
         End If
          
      ElseIf (lNewWidth > 0) Then
         If ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or _
             (ForceCloneCreation)) Then
            lNewHeight = lNewWidth / (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
            hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
         End If
      
      ElseIf (lNewHeight > 0) Then
         If ((lNewHeight <> FreeImage_GetHeight(BITMAP)) Or _
             (ForceCloneCreation)) Then
            lNewWidth = lNewHeight * (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
            hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
         End If
      
      End If
      
      If (hDIBNew) Then
         FreeImage_RescaleEx = hDIBNew
         If (UnloadSource) Then
            Call FreeImage_Unload(BITMAP)
         End If
      Else
         FreeImage_RescaleEx = BITMAP
      End If
   End If
                     
End Function

Public Function FreeImage_RescaleByPixel(ByVal BITMAP As Long, _
                                Optional ByVal WidthInPixels As Long, _
                                Optional ByVal HeightInPixels As Long, _
                                Optional ByVal UnloadSource As Boolean, _
                                Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, _
                                Optional ByVal ForceCloneCreation As Boolean) As Long
                                
   ' Thin wrapper for function 'FreeImage_RescaleEx' for removing method
   ' overload fake. This function rescales the image directly to the size
   ' specified by the 'WidthInPixels' and 'HeightInPixels' parameters.

   FreeImage_RescaleByPixel = FreeImage_RescaleEx(BITMAP, WidthInPixels, HeightInPixels, False, _
         UnloadSource, Filter, ForceCloneCreation)

End Function


' Painting functions

Public Function FreeImage_PaintDC(ByVal hDC As Long, _
                                  ByVal BITMAP As Long, _
                         Optional ByVal xDst As Long, _
                         Optional ByVal yDst As Long, _
                         Optional ByVal xSrc As Long, _
                         Optional ByVal ySrc As Long, _
                         Optional ByVal Width As Long, _
                         Optional ByVal Height As Long) As Long
 
   ' This function draws a FreeImage DIB directly onto a device context (DC). There
   ' are many (selfexplaining?) parameters that control the visual result.
   
   ' Parameters 'XDst' and 'YDst' specify the point where the output should
   ' be painted and 'XSrc', 'YSrc', 'Width' and 'Height' span a rectangle
   ' in the source image 'Bitmap' that defines the area to be painted.
   
   ' If any of parameters 'Width' and 'Height' is zero, it is transparently substituted
   ' by the width or height of teh bitmap to be drawn, resprectively.
   
   If ((hDC <> 0) And (BITMAP <> 0)) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to paint a 'header-only' bitmap.")
      End If
            
      If (Width = 0) Then
         Width = FreeImage_GetWidth(BITMAP)
      End If
      
      If (Height = 0) Then
         Height = FreeImage_GetHeight(BITMAP)
      End If
      
      FreeImage_PaintDC = SetDIBitsToDevice(hDC, xDst, yDst - ySrc, Width, Height, xSrc, ySrc, 0, _
            Height, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS)
   End If

End Function

Public Function FreeImage_PaintDCEx(ByVal hDC As Long, _
                                    ByVal BITMAP As Long, _
                           Optional ByVal xDst As Long, _
                           Optional ByVal yDst As Long, _
                           Optional ByVal WidthDst As Long, _
                           Optional ByVal HeightDst As Long, _
                           Optional ByVal xSrc As Long, _
                           Optional ByVal ySrc As Long, _
                           Optional ByVal WidthSrc As Long, _
                           Optional ByVal HeightSrc As Long, _
                           Optional ByVal DrawMode As DRAW_MODE = DM_DRAW_DEFAULT, _
                           Optional ByVal RasterOperator As RASTER_OPERATOR = ROP_SRCCOPY, _
                           Optional ByVal StretchMode As STRETCH_MODE = SM_COLORONCOLOR) As Long

Dim eLastStretchMode As STRETCH_MODE

   ' This function draws a FreeImage DIB directly onto a device context (DC). There
   ' are many (selfexplaining?) parameters that control the visual result.
   
   ' The main difference of this function compared to the 'FreeImage_PaintDC' is,
   ' that this function supports both mirroring and stretching of the image to be
   ' painted and so, is somewhat slower than 'FreeImage_PaintDC'.
   
   If ((hDC <> 0) And (BITMAP <> 0)) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to paint a 'header-only' bitmap.")
      End If
      
      eLastStretchMode = GetStretchBltMode(hDC)
      Call SetStretchBltMode(hDC, StretchMode)
      
      If (WidthSrc = 0) Then
         WidthSrc = FreeImage_GetWidth(BITMAP)
      End If
      If (WidthDst = 0) Then
         WidthDst = WidthSrc
      End If
      
      If (HeightSrc = 0) Then
         HeightSrc = FreeImage_GetHeight(BITMAP)
      End If
      If (HeightDst = 0) Then
         HeightDst = HeightSrc
      End If
      
      If (DrawMode And DM_MIRROR_VERTICAL) Then
         yDst = yDst + HeightDst
         HeightDst = -HeightDst
      End If
     
      If (DrawMode And DM_MIRROR_HORIZONTAL) Then
         xDst = xDst + WidthDst
         WidthDst = -WidthDst
      End If

      Call StretchDIBits(hDC, xDst, yDst, WidthDst, HeightDst, xSrc, ySrc, WidthSrc, HeightSrc, _
            FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, _
            RasterOperator)
      
      ' restore last mode
      Call SetStretchBltMode(hDC, eLastStretchMode)
   End If

End Function


Public Function FreeImage_GetTransparencyTableEx(ByVal BITMAP As Long) As Byte()

Dim tSA As SAVEARRAY1D
Dim lpSA As Long

   ' This function returns a VB style Byte array, containing the transparency
   ' table of the Bitmap. This array provides read and write access to the actual
   ' transparency table provided by FreeImage. This is done by creating a VB array
   ' with an own SAFEARRAY descriptor making the array point to the transparency
   ' table pointer returned by FreeImage_GetTransparencyTable().
   
   ' This makes you use code like you would in C/C++:
   
   ' // this code assumes there is a bitmap loaded and
   ' // present in a variable called ‘dib’
   ' if(FreeImage_GetBPP(Bitmap) == 8) {
   '   // Remove transparency information
   '   byte *transt = FreeImage_GetTransparencyTable(Bitmap);
   '   for (int i = 0; i < 256; i++) {
   '     transt[i].rgbRed = 255;
   '   }
   
   ' As in C/C++ the array is only valid while the DIB is loaded and the transparency
   ' table remains where the pointer returned by FreeImage_GetTransparencyTable() has
   ' pointed to when this function was called. So, a good thing would be, not to keep
   ' the returned array in scope over the lifetime of the DIB. Best practise is, to use
   ' this function within another routine and assign the return value (the array) to a
   ' local variable only. As soon as this local variable goes out of scope (when the
   ' calling function returns to it's caller), the array and the descriptor is
   ' automatically cleaned up by VB.
   
   ' This function does not make a deep copy of the transparency table, but only
   ' wraps a VB array around the FreeImage transparency table. So, it can be called
   ' frequently "on demand" or somewhat "in place" without a significant
   ' performance loss.
   
   ' To learn more about this technique I recommend reading chapter 2 (Leveraging
   ' Arrays) of Matthew Curland's book "Advanced Visual Basic 6"
   
   ' The parameter 'Bitmap' works according to the FreeImage 3 API documentation.
   
   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the FreeImage_DestroyLockedArray() function.
   
   
   If (BITMAP) Then
      
      ' create a proper SAVEARRAY descriptor
      With tSA
         .cbElements = 1                                     ' size in bytes of a byte element
         .cDims = 1                                          ' the array has only 1 dimension
         .cElements = FreeImage_GetTransparencyCount(BITMAP) ' the number of elements in the array is
                                                             ' equal to the number transparency table entries
         .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE            ' need AUTO and FIXEDSIZE for safety issues,
                                                             ' so the array can not be modified in size
                                                             ' or erased; according to Matthew Curland never
                                                             ' use FIXEDSIZE alone
         .pvData = FreeImage_GetTransparencyTable(BITMAP)    ' let the array point to the memory block, the
                                                             ' FreeImage transparency table pointer points to
      End With
      
      ' allocate memory for an array descriptor
      ' we cannot use the memory block used by tSA, since it is
      ' released when tSA goes out of scope, leaving us with an
      ' array with zeroed descriptor
      ' we use nearly the same method that VB uses, so VB is able
      ' to cleanup the array variable and it's descriptor; the
      ' array data is not touched when cleaning up, since both AUTO
      ' and FIXEDSIZE flags are set
      Call SafeArrayAllocDescriptor(1, lpSA)
      
      ' copy our own array descriptor over the descriptor allocated
      ' by SafeArrayAllocDescriptor(); lpSA is a pointer to that memory
      ' location
      Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
      
      ' the implicit variable named as the function is an array variable in VB
      ' make it point to the allocated array descriptor
      Call CopyMemory(ByVal VarPtrArray(FreeImage_GetTransparencyTableEx), lpSA, 4)
   End If

End Function

Public Function FreeImage_PaintTransparent(ByVal hDC As Long, _
                                           ByVal BITMAP As Long, _
                                  Optional ByVal xDst As Long = 0, _
                                  Optional ByVal yDst As Long = 0, _
                                  Optional ByVal WidthDst As Long, _
                                  Optional ByVal HeightDst As Long, _
                                  Optional ByVal xSrc As Long = 0, _
                                  Optional ByVal ySrc As Long = 0, _
                                  Optional ByVal WidthSrc As Long, _
                                  Optional ByVal HeightSrc As Long, _
                                  Optional ByVal Alpha As Byte = 255) As Long
                                  
Dim lpPalette As Long
Dim bIsTransparent As Boolean

   ' This function paints a device independent bitmap to any device context and
   ' thereby honors any transparency information associated with the bitmap.
   ' Furthermore, through the 'Alpha' parameter, an overall transparency level
   ' may be specified.
   
   ' For palletised images, any color set to be transparent in the transparency
   ' table, will be transparent. For high color images, only 32-bit images may
   ' have any transparency information associated in their alpha channel. Only
   ' these may be painted with transparency by this function.
   
   ' Since this is a wrapper for the Windows GDI function AlphaBlend(), 31-bit
   ' images, containing alpha (or per-pixel) transparency, must be 'premultiplied'
   ' for alpha transparent regions to actually show transparent. See MSDN help
   ' on the AlphaBlend() function.
   
   ' FreeImage also offers a function to premultiply 32-bit bitmaps with their alpha
   ' channel, according to the needs of AlphaBlend(). Have a look at function
   ' FreeImage_PreMultiplyWithAlpha().
   
   ' Overall transparency level may be specified for all bitmaps in all color
   ' depths supported by FreeImage. If needed, bitmaps are transparently converted
   ' to 32-bit and unloaded after the paint operation. This is also true for palletised
   ' bitmaps.
   
   ' Parameters 'hDC' and 'Bitmap' seem to be very self-explanatory. All other parameters
   ' are optional. The group of '*Dest*' parameters span a rectangle on the destination
   ' device context, used as drawing area for the bitmap. If these are omitted, the
   ' bitmap will be drawn starting at position 0,0 in the bitmap's actual size.
   ' The group of '*Src*' parameters span a rectangle on the source bitmap, used as
   ' cropping area for the paint operation. If both rectangles differ in size in any
   ' direction, the part of the image actually painted is stretched for to fit into
   ' the drawing area. If any of the parameters '*Width' or '*Height' are omitted,
   ' the bitmap's actual size (width or height) will be used.
   
   ' The 'Alpha' parameter specifies the overall transparency. It takes values in the
   ' range from 0 to 255. Using 0 will paint the bitmap fully transparent, 255 will
   ' paint the image fully opaque. The 'Alpha' value controls, how the non per-pixel
   ' portions of the image will be drawn.
                                  
   If ((hDC <> 0) And (BITMAP <> 0)) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to paint a 'header-only' bitmap.")
      End If
   
      ' get image width if not specified
      If (WidthSrc = 0) Then
         WidthSrc = FreeImage_GetWidth(BITMAP)
      End If
      If (WidthDst = 0) Then
         WidthDst = WidthSrc
      End If
      
      ' get image height if not specified
      If (HeightSrc = 0) Then
         HeightSrc = FreeImage_GetHeight(BITMAP)
      End If
      If (HeightDst = 0) Then
         HeightDst = HeightSrc
      End If
      
      lpPalette = FreeImage_GetPalette(BITMAP)
      If (lpPalette) Then
      
         Dim lPaletteSize As Long
         Dim alPalOrg(255) As Long
         Dim alPalMod(255) As Long
         Dim alPalMask(255) As Long
         Dim abTT() As Byte
         Dim i As Long
         
         lPaletteSize = FreeImage_GetColorsUsed(BITMAP) * 4
         Call CopyMemory(alPalOrg(0), ByVal lpPalette, lPaletteSize)
         Call CopyMemory(alPalMod(0), ByVal lpPalette, lPaletteSize)
         abTT = FreeImage_GetTransparencyTableEx(BITMAP)
         
         If ((Alpha = 255) And _
             (HeightDst >= HeightSrc) And (WidthDst >= WidthSrc)) Then
            
            ' create a mask palette and a modified version of the
            ' original palette
            For i = 0 To UBound(abTT)
               If (abTT(i) = 0) Then
                  alPalMask(i) = &HFFFFFFFF   ' white
                  alPalMod(i) = &H0           ' black
                  bIsTransparent = True
               End If
            Next i

            If (Not bIsTransparent) Then
               
               ' if there is no transparency in the image, paint it with
               ' a single SRCCOPY
               Call StretchDIBits(hDC, _
                                  xDst, yDst, WidthDst, HeightDst, _
                                  xSrc, ySrc, WidthSrc, HeightSrc, _
                                  FreeImage_GetBits(BITMAP), _
                                  FreeImage_GetInfo(BITMAP), _
                                  DIB_RGB_COLORS, SRCCOPY)
            Else
            
               ' set mask palette and paint with SRCAND
               Call CopyMemory(ByVal lpPalette, alPalMask(0), lPaletteSize)
               Call StretchDIBits(hDC, _
                                  xDst, yDst, WidthDst, HeightDst, _
                                  xSrc, ySrc, WidthSrc, HeightSrc, _
                                  FreeImage_GetBits(BITMAP), _
                                  FreeImage_GetInfo(BITMAP), _
                                  DIB_RGB_COLORS, SRCAND)
               
               ' set mask modified and paint with SRCPAINT
               Call CopyMemory(ByVal lpPalette, alPalMod(0), lPaletteSize)
               Call StretchDIBits(hDC, _
                                  xDst, yDst, WidthDst, HeightDst, _
                                  xSrc, ySrc, WidthSrc, HeightSrc, _
                                  FreeImage_GetBits(BITMAP), _
                                  FreeImage_GetInfo(BITMAP), _
                                  DIB_RGB_COLORS, SRCPAINT)
                                  
               ' restore original palette
               Call CopyMemory(ByVal lpPalette, alPalOrg(0), lPaletteSize)
            End If
            
            ' we are done, do not paint with AlphaBlend() any more
            BITMAP = 0
         Else
            
            ' create a premultiplied palette
            ' since we have no real per pixel transparency in a palletized
            ' image, we only need to set all transparent colors to zero.
            For i = 0 To UBound(abTT)
               If (abTT(i) = 0) Then
                  alPalMod(i) = 0
               End If
            Next i
            
            ' set premultiplied palette and convert to 32 bits
            Call CopyMemory(ByVal lpPalette, alPalMod(0), lPaletteSize)
            BITMAP = FreeImage_ConvertTo32Bits(BITMAP)
            
            ' restore original palette
            Call CopyMemory(ByVal lpPalette, alPalOrg(0), lPaletteSize)
         End If
      End If

      If (BITMAP) Then
         Dim hMemDC As Long
         Dim hBitmap As Long
         Dim hBitmapOld As Long
         Dim tBF As BLENDFUNCTION
         Dim lBF As Long
         
         hMemDC = Drawing.GetMemoryDC()
         If (hMemDC) Then
            hBitmap = FreeImage_GetBitmap(BITMAP, hMemDC)
            hBitmapOld = SelectObject(hMemDC, hBitmap)
            
            With tBF
               .BlendOp = AC_SRC_OVER
               .SourceConstantAlpha = Alpha
               If (FreeImage_GetBPP(BITMAP) = 32) Then
                  .AlphaFormat = AC_SRC_ALPHA
               End If
            End With
            Call CopyMemory(lBF, tBF, 4)
            
            Call AlphaBlend(hDC, xDst, yDst, WidthDst, HeightDst, _
                            hMemDC, xSrc, ySrc, WidthSrc, HeightSrc, _
                            lBF)
                            
            Call SelectObject(hMemDC, hBitmapOld)
            Call DeleteObject(hBitmap)
            Drawing.FreeMemoryDC hMemDC
            If (lpPalette) Then
               Call FreeImage_Unload(BITMAP)
            End If
         End If
      End If
   End If

End Function

'--------------------------------------------------------------------------------
' Pixel access functions
'--------------------------------------------------------------------------------

Public Function FreeImage_GetBitsEx(ByVal BITMAP As Long) As Byte()

Dim tSA As SAVEARRAY2D
Dim lpSA As Long

   ' This function returns a two dimensional Byte array containing a DIB's
   ' data-bits. This is done by wrapping a true VB array around the memory
   ' block the returned pointer of FreeImage_GetBits() is pointing to. So, the
   ' array returned provides full read and write acces to the image's data.

   ' To reuse the caller's array variable, this function's result was assigned to,
   ' before it goes out of scope, the caller's array variable must be destroyed with
   ' the FreeImage_DestroyLockedArray() function.

   If (BITMAP) Then
      
      ' create a proper SAVEARRAY descriptor
      With tSA
         .cbElements = 1                           ' size in bytes per array element
         .cDims = 2                                ' the array has 2 dimensions
         .cElements1 = FreeImage_GetHeight(BITMAP) ' the number of elements in y direction (height of Bitmap)
         .cElements2 = FreeImage_GetPitch(BITMAP)  ' the number of elements in x direction (byte width of Bitmap)
         .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues,
                                                   ' so the array can not be modified in size
                                                   ' or erased; according to Matthew Curland never
                                                   ' use FIXEDSIZE alone
         .pvData = FreeImage_GetBits(BITMAP)       ' let the array point to the memory block, the
                                                   ' FreeImage scanline data pointer points to
      End With
      
      ' allocate memory for an array descriptor
      ' we cannot use the memory block used by tSA, since it is
      ' released when tSA goes out of scope, leaving us with an
      ' array with zeroed descriptor
      ' we use nearly the same method that VB uses, so VB is able
      ' to cleanup the array variable and it's descriptor; the
      ' array data is not touched when cleaning up, since both AUTO
      ' and FIXEDSIZE flags are set
      Call SafeArrayAllocDescriptor(2, lpSA)
      
      ' copy our own array descriptor over the descriptor allocated
      ' by SafeArrayAllocDescriptor; lpSA is a pointer to that memory
      ' location
      Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
      
      ' the implicit variable named like the function is an array
      ' variable in VB
      ' make it point to the allocated array descriptor
      Call CopyMemory(ByVal VarPtrArray(FreeImage_GetBitsEx), lpSA, 4)
   End If

End Function


'--------------------------------------------------------------------------------el a
' HBITMAP conversion functions
'--------------------------------------------------------------------------------

Public Function FreeImage_GetBitmap(ByVal BITMAP As Long, _
                           Optional ByVal hDC As Long, _
                           Optional ByVal UnloadSource As Boolean) As Long
                               
Dim bReleaseDC As Boolean
Dim ppvBits As Long
   
   ' This function returns an HBITMAP created by the CreateDIBSection() function which
   ' in turn has the same color depth as the original DIB. A reference DC may be provided
   ' through the 'hDC' parameter. The desktop DC will be used, if no reference DC is
   ' specified.

   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to create a bitmap from a 'header-only' bitmap.")
      End If
   
      If (hDC = 0) Then
         hDC = GetDC(0)
         bReleaseDC = True
      End If
      If (hDC) Then
         FreeImage_GetBitmap = CreateDIBSection(hDC, FreeImage_GetInfo(BITMAP), _
               DIB_RGB_COLORS, ppvBits, 0, 0)
         If ((FreeImage_GetBitmap <> 0) And (ppvBits <> 0)) Then
            Call CopyMemory(ByVal ppvBits, ByVal FreeImage_GetBits(BITMAP), _
                  FreeImage_GetHeight(BITMAP) * FreeImage_GetPitch(BITMAP))
         End If
         If (UnloadSource) Then
            Call FreeImage_Unload(BITMAP)
         End If
         If (bReleaseDC) Then
            Call ReleaseDC(0, hDC)
         End If
      End If
   End If

End Function


Public Function FreeImage_GetBitmapForDevice(ByVal BITMAP As Long, _
                                    Optional ByVal hDC As Long, _
                                    Optional ByVal UnloadSource As Boolean) As Long
                                    
Dim bReleaseDC As Boolean

   ' This function returns an HBITMAP created by the CreateDIBitmap() function which
   ' in turn has always the same color depth as the reference DC, which may be provided
   ' through the 'hDC' parameter. The desktop DC will be used, if no reference DC is
   ' specified.
                              
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to create a bitmap from a 'header-only' bitmap.")
      End If
   
      If (hDC = 0) Then
         hDC = GetDC(0)
         bReleaseDC = True
      End If
      If (hDC) Then
         FreeImage_GetBitmapForDevice = _
               CreateDIBitmap(hDC, FreeImage_GetInfoHeader(BITMAP), CBM_INIT, _
                     FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), _
                           DIB_RGB_COLORS)
         If (UnloadSource) Then
            Call FreeImage_Unload(BITMAP)
         End If
         If (bReleaseDC) Then
            Call ReleaseDC(0, hDC)
         End If
      End If
   End If

End Function

'--------------------------------------------------------------------------------
' OlePicture conversion functions
'--------------------------------------------------------------------------------

Public Function FreeImage_GetOlePicture(ByVal BITMAP As Long, _
                               Optional ByVal hDC As Long, _
                               Optional ByVal UnloadSource As Boolean) As IPicture

Dim hBitmap As Long
Dim tPicDesc As PictDesc
Dim tGuid As Guid
Dim cPictureDisp As IPictureDisp

   ' This function creates a VB Picture object (OlePicture) from a FreeImage DIB.
   ' The original image need not remain valid nor loaded after the VB Picture
   ' object has been created.
   
   ' The optional parameter 'hDC' determines the device context (DC) used for
   ' transforming the device independent bitmap (DIB) to a device dependent
   ' bitmap (DDB). This device context's color depth is responsible for this
   ' transformation. This parameter may be null or omitted. In that case, the
   ' windows desktop's device context will be used, what will be the desired
   ' way in almost any cases.
   
   ' The optional 'UnloadSource' parameter is for unloading the original image
   ' after the OlePicture has been created, so you can easily "switch" from a
   ' FreeImage DIB to a VB Picture object. There is no need to unload the DIB
   ' at the caller's site if this argument is True.
   
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to create a picture from a 'header-only' bitmap.")
      End If
   
      hBitmap = FreeImage_GetBitmapForDevice(BITMAP, hDC, UnloadSource)
      If (hBitmap) Then
         ' fill tPictDesc structure with necessary parts
         With tPicDesc
            .cbSizeofStruct = Len(tPicDesc)
            ' the vbPicTypeBitmap constant is not available in VBA environemnts
            .picType = 1  'vbPicTypeBitmap
            .hImage = hBitmap
         End With
   
         ' fill in IDispatch Interface ID
         With tGuid
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With
   
         ' create a picture object
         Call OleCreatePictureIndirect(tPicDesc, tGuid, True, cPictureDisp)
         Set FreeImage_GetOlePicture = cPictureDisp
      End If
   End If

End Function



Public Function FreeImage_CreateFromOlePicture(ByRef Picture As IPicture) As Long

Dim hBitmap As Long
Dim tBM As BITMAP_API
Dim hDIB As Long
Dim hDC As Long
Dim lResult As Long
Dim nColors As Long
Dim lpInfo As Long

   ' Creates a FreeImage DIB from a VB Picture object (OlePicture). This function
   ' returns a pointer to the DIB as, for instance, the FreeImage function
   ' 'FreeImage_Load' does. So, this could be a real replacement for 'FreeImage_Load'
   ' when working with VB Picture objects.

   If (Not Picture Is Nothing) Then
      hBitmap = Picture.Handle
      If (hBitmap) Then
         lResult = GetObjectAPI(hBitmap, Len(tBM), tBM)
         If (lResult) Then
            hDIB = FreeImage_Allocate(tBM.bmWidth, _
                                      tBM.bmHeight, _
                                      tBM.bmBitsPixel)
            If (hDIB) Then
               ' The GetDIBits function clears the biClrUsed and biClrImportant BITMAPINFO
               ' members (dont't know why). So we save these infos below.
               ' This is needed for palletized images only.
               nColors = FreeImage_GetColorsUsed(hDIB)
            
               hDC = GetDC(0)
               lResult = GetDIBits(hDC, hBitmap, 0, _
                                   FreeImage_GetHeight(hDIB), _
                                   FreeImage_GetBits(hDIB), _
                                   FreeImage_GetInfo(hDIB), _
                                   DIB_RGB_COLORS)
               If (lResult) Then
                  FreeImage_CreateFromOlePicture = hDIB
                  If (nColors) Then
                     ' restore BITMAPINFO members
                     ' FreeImage_GetInfo(Bitmap)->biClrUsed = nColors;
                     ' FreeImage_GetInfo(Bitmap)->biClrImportant = nColors;
                     lpInfo = FreeImage_GetInfo(hDIB)
                     Call CopyMemory(ByVal lpInfo + 32, nColors, 4)
                     Call CopyMemory(ByVal lpInfo + 36, nColors, 4)
                  End If
               Else
                  Call FreeImage_Unload(hDIB)
               End If
               Call ReleaseDC(0, hDC)
            End If
         End If
      End If
   End If

End Function

Public Function FreeImage_CreateFromDC(ByVal hDC As Long, _
                              Optional ByRef hBitmap As Long) As Long

Dim tBM As BITMAP_API
Dim hDIB As Long
Dim lResult As Long
Dim nColors As Long
Dim lpInfo As Long

   ' Creates a FreeImage DIB from a Device Context/Compatible Bitmap. This
   ' function returns a pointer to the DIB as, for instance, 'FreeImage_Load()'
   ' does. So, this could be a real replacement for FreeImage_Load() or
   ' 'FreeImage_CreateFromOlePicture()' when working with DCs and BITMAPs directly
   
   ' The 'hDC' parameter specifies a window device context (DC), the optional
   ' parameter 'hBitmap' may specify a handle to a memory bitmap. When 'hBitmap' is
   ' omitted, the bitmap currently selected into the given DC is used to create
   ' the DIB.
   
   ' When 'hBitmap' is not missing but NULL (0), the function uses the DC's currently
   ' selected bitmap. This bitmap's handle is stored in the ('ByRef'!) 'hBitmap' parameter
   ' and so, is avaliable at the caller's site when the function returns.
   
   ' The DIB returned by this function is a copy of the image specified by 'hBitmap' or
   ' the DC's current bitmap when 'hBitmap' is missing. The 'hDC' and also the 'hBitmap'
   ' remain untouched in this function, there will be no objects destroyed or freed.
   ' The caller is responsible to destroy or free the DC and BITMAP if necessary.
   
   ' first, check whether we got a hBitmap or not
   If (hBitmap = 0) Then
      ' if not, the parameter may be missing or is NULL so get the
      ' DC's current bitmap
      hBitmap = GetCurrentObject(hDC, OBJ_BITMAP)
   End If

   lResult = GetObjectAPI(hBitmap, Len(tBM), tBM)
   If (lResult) Then
      hDIB = FreeImage_Allocate(tBM.bmWidth, _
                                tBM.bmHeight, _
                                tBM.bmBitsPixel)
      If (hDIB) Then
         ' The GetDIBits function clears the biClrUsed and biClrImportant BITMAPINFO
         ' members (dont't know why). So we save these infos below.
         ' This is needed for palletized images only.
         nColors = FreeImage_GetColorsUsed(hDIB)
         
         lResult = GetDIBits(hDC, hBitmap, 0, _
                             FreeImage_GetHeight(hDIB), _
                             FreeImage_GetBits(hDIB), _
                             FreeImage_GetInfo(hDIB), _
                             DIB_RGB_COLORS)
                             
         If (lResult) Then
            FreeImage_CreateFromDC = hDIB
            If (nColors) Then
               ' restore BITMAPINFO members
               ' FreeImage_GetInfo(Bitmap)->biClrUsed = nColors;
               ' FreeImage_GetInfo(Bitmap)->biClrImportant = nColors;
               lpInfo = FreeImage_GetInfo(hDIB)
               Call CopyMemory(ByVal lpInfo + 32, nColors, 4)
               Call CopyMemory(ByVal lpInfo + 36, nColors, 4)
            End If
         Else
            Call FreeImage_Unload(hDIB)
         End If
      End If
   End If

End Function



Public Function FreeImage_LoadEx(ByVal FileName As String, _
                        Optional ByVal Options As FREE_IMAGE_LOAD_OPTIONS, _
                        Optional ByVal Width As Variant, _
                        Optional ByVal Height As Variant, _
                        Optional ByVal InPercent As Boolean, _
                        Optional ByVal Filter As FREE_IMAGE_FILTER, _
                        Optional ByRef Format As FREE_IMAGE_FORMAT) As Long

Const vbInvalidPictureError As Long = 481

   ' The function provides all image formats, the FreeImage library can read. The
   ' image format is determined from the image file to load, the optional parameter
   ' 'Format' is an OUT parameter that will contain the image format that has
   ' been loaded.
   
   ' The parameters 'Width', 'Height', 'InPercent' and 'Filter' make it possible
   ' to "load" the image in a resized version. 'Width', 'Height' specify the desired
   ' width and height, 'Filter' determines, what image filter should be used
   ' on the resizing process.
   
   ' The parameters 'Width', 'Height', 'InPercent' and 'Filter' map directly to the
   ' according parameters of the 'FreeImage_RescaleEx' function. So, read the
   ' documentation of the 'FreeImage_RescaleEx' for a complete understanding of the
   ' usage of these parameters.
   

   Format = FreeImage_GetFileTypeU(StrPtr(FileName))
   If (Format <> FIF_UNKNOWN) Then
      If (FreeImage_FIFSupportsReading(Format)) Then
         FreeImage_LoadEx = FreeImage_LoadUInt(Format, StrPtr(FileName), Options)
         If (FreeImage_LoadEx) Then
            
            If ((Not IsMissing(Width)) Or _
                (Not IsMissing(Height))) Then
               FreeImage_LoadEx = FreeImage_RescaleEx(FreeImage_LoadEx, Width, Height, _
                     InPercent, True, Filter)
            End If
         Else
            Call Err.Raise(vbInvalidPictureError)
         End If
      Else
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & _
                        "does not support reading.")
      End If
   Else
      Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                     "The file specified has an unknown image format.")
   End If

End Function

Public Function LoadPictureEx(Optional ByRef FileName As Variant, _
                              Optional ByRef Options As FREE_IMAGE_LOAD_OPTIONS, _
                              Optional ByRef Width As Variant, _
                              Optional ByRef Height As Variant, _
                              Optional ByRef InPercent As Boolean, _
                              Optional ByRef Filter As FREE_IMAGE_FILTER, _
                              Optional ByRef Format As FREE_IMAGE_FORMAT) As IPicture
                              
Dim hDIB As Long

   ' This function is an extended version of the VB method 'LoadPicture'. As
   ' the VB version it takes a filename parameter to load the image and throws
   ' the same errors in most cases.
   
   ' This function now is only a thin wrapper for the FreeImage_LoadEx() wrapper
   ' function (as compared to releases of this wrapper prior to version 1.8). So,
   ' have a look at this function's discussion of the parameters.
   
   ' However, we do mask out the FILO_LOAD_NOPIXELS load option, since this
   ' function shall create a VB Picture object, which does not support
   ' FreeImage's header-only loading option.


   If (Not IsMissing(FileName)) Then
      hDIB = FreeImage_LoadEx(FileName, (Options And (Not FILO_LOAD_NOPIXELS)), _
            Width, Height, InPercent, Filter, Format)
      Set LoadPictureEx = FreeImage_GetOlePicture(hDIB, , True)
   End If

End Function

Public Function FreeImage_SaveEx(ByVal BITMAP As Long, _
                                 ByVal FileName As String, _
                        Optional ByVal Format As FREE_IMAGE_FORMAT = FIF_UNKNOWN, _
                        Optional ByVal Options As FREE_IMAGE_SAVE_OPTIONS, _
                        Optional ByVal colorDepth As FREE_IMAGE_COLOR_DEPTH, _
                        Optional ByVal Width As Variant, _
                        Optional ByVal Height As Variant, _
                        Optional ByVal InPercent As Boolean, _
                        Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, _
                        Optional ByVal UnloadSource As Boolean) As Boolean
                     
Dim hDIBRescale As Long
Dim bConvertedOnRescale As Boolean
Dim bIsNewDIB As Boolean
Dim lBPP As Long
Dim lBPPOrg As Long
Dim strExtension As String

   ' This function is an easy to use replacement for FreeImage's FreeImage_Save()
   ' function which supports inline size- and color conversions as well as an
   ' auto image format detection algorithm that determines the desired image format
   ' by the given filename. An even more sophisticated algorithm may auto-detect
   ' the proper color depth for a explicitly given or auto-detected image format.

   ' The function provides all image formats, and save options, the FreeImage
   ' library can write. The optional parameter 'Format' may contain the desired
   ' image format. When omitted, the function tries to get the image format from
   ' the filename extension.
   
   ' The optional parameter 'ColorDepth' may contain the desired color depth for
   ' the saved image. This can be either any value of the FREE_IMAGE_COLOR_DEPTH
   ' enumeration or the value FICD_AUTO what is the default value of the parameter.
   ' When 'ColorDepth' is FICD_AUTO, the function tries to get the most suitable
   ' color depth for the specified image format if the image's current color depth
   ' is not supported by the specified image format. Therefore, the function
   ' firstly reduces the color depth step by step until a proper color depth is
   ' found since an incremention would only increase the file's size with no
   ' quality benefit. Only when there is no lower color depth is found for the
   ' image format, the function starts to increase the color depth.
   
   ' Keep in mind that an explicitly specified color depth that is not supported
   ' by the image format results in a runtime error. For example, when saving
   ' a 24 bit image as GIF image, a runtime error occurs.
   
   ' The function checks, whether the given filename has a valid extension or
   ' not. If not, the "primary" extension for the used image format will be
   ' appended to the filename. The parameter 'Filename' remains untouched in
   ' this case.
   
   ' To learn more about the "primary" extension, read the documentation for
   ' the 'FreeImage_GetPrimaryExtensionFromFIF' function.
   
   ' The parameters 'Width', 'Height', 'InPercent' and 'Filter' make it possible
   ' to save the image in a resized version. 'Width', 'Height' specify the desired
   ' width and height, 'Filter' determines, what image filter should be used
   ' on the resizing process. Since FreeImage_SaveEx relies on FreeImage_RescaleEx,
   ' please refer to the documentation of FreeImage_RescaleEx to learn more
   ' about these four parameters.
   
   ' The optional 'UnloadSource' parameter is for unloading the saved image, so
   ' you can save and unload an image with this function in one operation.
   ' CAUTION: at current, the image is unloaded, even if the image was not
   '          saved correctly!

   
   If (BITMAP) Then
   
      If (Not FreeImage_HasPixels(BITMAP)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to save 'header-only' bitmaps.")
      End If
   
      If ((Not IsMissing(Width)) Or _
          (Not IsMissing(Height))) Then
          
         lBPP = FreeImage_GetBPP(BITMAP)
         hDIBRescale = FreeImage_RescaleEx(BITMAP, Width, Height, InPercent, UnloadSource, Filter)
         bIsNewDIB = (hDIBRescale <> BITMAP)
         BITMAP = hDIBRescale
         bConvertedOnRescale = (lBPP <> FreeImage_GetBPP(BITMAP))
      End If
      
      If (Format = FIF_UNKNOWN) Then
         Format = FreeImage_GetFIFFromFilenameU(StrPtr(FileName))
      End If
      If (Format <> FIF_UNKNOWN) Then
         If ((FreeImage_FIFSupportsWriting(Format)) And _
             (FreeImage_FIFSupportsExportType(Format, FIT_BITMAP))) Then
            
            If (Not FreeImage_IsFilenameValidForFIF(Format, FileName)) Then
               'Edit by Tanner: don't prevent me from writing whatever file extensions I damn well please!  ;)
               'strExtension = "." & FreeImage_GetPrimaryExtensionFromFIF(Format)
            End If
            
            ' check color depth
            If (colorDepth <> FICD_AUTO) Then
               ' mask out bit 1 (0x02) for the case ColorDepth is FICD_MONOCHROME_DITHER (0x03)
               ' FREE_IMAGE_COLOR_DEPTH values are true bit depths in general except FICD_MONOCHROME_DITHER
               ' by masking out bit 1, 'FreeImage_FIFSupportsExportBPP()' tests for bitdepth 1
               ' what is correct again for dithered images.
               colorDepth = (colorDepth And (Not &H2))
               If (Not FreeImage_FIFSupportsExportBPP(Format, colorDepth)) Then
                  Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                                 "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & _
                                 "is unable to write images with a color depth " & _
                                 "of " & colorDepth & " bpp.")
               
               ElseIf (FreeImage_GetBPP(BITMAP) <> colorDepth) Then
               
                  BITMAP = FreeImage_ConvertColorDepth(BITMAP, colorDepth, (UnloadSource Or bIsNewDIB))
                  bIsNewDIB = True
               
               End If
            Else
            
               If (lBPP = 0) Then
                  lBPP = FreeImage_GetBPP(BITMAP)
               End If
               
               If (Not FreeImage_FIFSupportsExportBPP(Format, lBPP)) Then
                  lBPPOrg = lBPP
                  Do
                     lBPP = pGetPreviousColorDepth(lBPP)
                  Loop While ((Not FreeImage_FIFSupportsExportBPP(Format, lBPP)) Or _
                              (lBPP = 0))
                  If (lBPP = 0) Then
                     lBPP = lBPPOrg
                     Do
                        lBPP = pGetNextColorDepth(lBPP)
                     Loop While ((Not FreeImage_FIFSupportsExportBPP(Format, lBPP)) Or _
                                 (lBPP = 0))
                  End If
                  
                  If (lBPP <> 0) Then
                     BITMAP = FreeImage_ConvertColorDepth(BITMAP, lBPP, (UnloadSource Or bIsNewDIB))
                     bIsNewDIB = True
                  End If
               
               ElseIf (bConvertedOnRescale) Then
                  ' restore original color depth
                  ' always unload current DIB here, since 'bIsNewDIB' is True
                  BITMAP = FreeImage_ConvertColorDepth(BITMAP, lBPP, True)
                  
               End If
            End If
            
            FreeImage_SaveEx = FreeImage_Save(Format, BITMAP, FileName & strExtension, Options)
            If ((bIsNewDIB) Or (UnloadSource)) Then
               Call FreeImage_Unload(BITMAP)
            End If
         Else
            Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                           "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & _
                           "is unable to write images of the image format requested.")
         End If
      Else
         ' unknown image format error
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unknown image format. Neither an explicit image format " & _
                        "was specified nor any known image format was determined " & _
                        "from the filename specified.")
      End If
   End If

End Function

Public Function SavePictureEx(ByRef Picture As IPicture, _
                              ByRef FileName As String, _
                     Optional ByRef Format As FREE_IMAGE_FORMAT, _
                     Optional ByRef Options As FREE_IMAGE_SAVE_OPTIONS, _
                     Optional ByRef colorDepth As FREE_IMAGE_COLOR_DEPTH, _
                     Optional ByRef Width As Variant, _
                     Optional ByRef Height As Variant, _
                     Optional ByRef InPercent As Boolean, _
                     Optional ByRef Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As Boolean
                     
Dim hDIB As Long

Const vbObjectOrWithBlockVariableNotSet As Long = 91
Const vbInvalidPictureError As Long = 481

   ' This function is an extended version of the VB method 'SavePicture'. As
   ' the VB version it takes a Picture object and a filename parameter to
   ' save the image and throws the same errors in most cases.
   
   ' This function now is only a thin wrapper for the FreeImage_SaveEx() wrapper
   ' function (as compared to releases of this wrapper prior to version 1.8). So,
   ' have a look at this function's discussion of the parameters.
   
   
   If (Not Picture Is Nothing) Then
      hDIB = FreeImage_CreateFromOlePicture(Picture)
      If (hDIB) Then
         SavePictureEx = FreeImage_SaveEx(hDIB, FileName, Format, Options, _
                                          colorDepth, Width, Height, InPercent, _
                                          FILTER_BICUBIC, True)
      Else
         Call Err.Raise(vbInvalidPictureError)
      End If
   Else
      Call Err.Raise(vbObjectOrWithBlockVariableNotSet)
   End If

End Function

Public Function SaveImageContainerEx(ByRef Container As Object, _
                                     ByRef FileName As String, _
                            Optional ByVal IncludeDrawings As Boolean, _
                            Optional ByRef Format As FREE_IMAGE_FORMAT, _
                            Optional ByRef Options As FREE_IMAGE_SAVE_OPTIONS, _
                            Optional ByRef colorDepth As FREE_IMAGE_COLOR_DEPTH, _
                            Optional ByRef Width As Variant, _
                            Optional ByRef Height As Variant, _
                            Optional ByRef InPercent As Boolean, _
                            Optional ByRef Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As Long
                            
   ' This function is an extended version of the VB method 'SavePicture'. As
   ' the VB version it takes an image hosting control and a filename parameter to
   ' save the image and throws the same errors in most cases.
   
   ' This function merges the functionality of both wrapper functions
   ' 'SavePictureEx()' and 'FreeImage_CreateFromImageContainer()'. Basically this
   ' function is identical to 'SavePictureEx' expect that is does not take a
   ' IOlePicture (IPicture) object but a VB image hosting container control.
   
   ' Please, refer to each of this two function's inline documentation for a
   ' more detailed description.
                            
   Call SavePictureEx(pGetIOlePictureFromContainer(Container, IncludeDrawings), _
            FileName, Format, Options, colorDepth, Width, Height, InPercent, Filter)

End Function

Public Function FreeImage_OpenMultiBitmapEx(ByVal FileName As String, _
                                   Optional ByVal ReadOnly As Boolean, _
                                   Optional ByVal KeepCacheInMemory As Boolean, _
                                   Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, _
                                   Optional ByRef Format As FREE_IMAGE_FORMAT) As Long

   Format = FreeImage_GetFileTypeU(StrPtr(FileName))
   If (Format <> FIF_UNKNOWN) Then
      Select Case Format
      
      Case FIF_TIFF, FIF_GIF, FIF_ICO
         FreeImage_OpenMultiBitmapEx = FreeImage_OpenMultiBitmap(Format, FileName, False, _
               ReadOnly, KeepCacheInMemory, Flags)
      
      Case Else
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & _
                        "does not have any support for multi-page bitmaps.")
      End Select
   Else
      Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                     "The file specified has an unknown image format.")
   End If
   
End Function

Public Function FreeImage_CreateMultiBitmapEx(ByVal FileName As String, _
                                     Optional ByVal KeepCacheInMemory As Boolean, _
                                     Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, _
                                     Optional ByRef Format As FREE_IMAGE_FORMAT) As Long

   If (Format = FIF_UNKNOWN) Then
      Format = FreeImage_GetFIFFromFilenameU(StrPtr(FileName))
   End If
   
   If (Format <> FIF_UNKNOWN) Then
      Select Case Format
      
      Case FIF_TIFF, FIF_GIF, FIF_ICO
         FreeImage_CreateMultiBitmapEx = FreeImage_OpenMultiBitmap(Format, FileName, True, _
               False, KeepCacheInMemory, Flags)
      
      Case Else
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                       "FreeImage Library plugin '" & _
                       FreeImage_GetFormatFromFIF(Format) & "' " & _
                       "does not have any support for multi-page bitmaps.")
      End Select
   Else
      ' unknown image format error
      Call Err.Raise(5, _
                     "MFreeImage", _
                     Error$(5) & vbCrLf & vbCrLf & _
                     "Unknown image format. Neither an explicit image format " & _
                     "was specified nor any known image format was determined " & _
                     "from the filename specified.")
   End If

End Function

'NOTE FROM TANNER: the original .bas file contained many more OLE picture wrappers,
' but I have removed them from PhotoDemon's copy for brevity's sake

Public Function FreeImage_RotateIOP(ByRef Picture As IPicture, _
                                    ByVal Angle As Double, _
                           Optional ByVal ColorPtr As Long) As IPicture

Dim hDIB As Long
Dim hDIBNew As Long

   ' IOlePicture based wrapper for FreeImage function FreeImage_Rotate()
   
   ' The optional ColorPtr parameter takes a pointer to (e.g. the address of) an
   ' RGB color value. So, all these assignments are valid for ColorPtr:
   '
   'Dim tColor As RGBQUAD
   'tColor.rgbRed = 255
   'tColor.rgbGreen = 255
   'tColor.rgbBlue = 255
'
   'ColorPtr = VarPtr(tColor)
   ' VarPtr(&H33FF80)
   ' VarPtr(vbWhite) ' However, the VB color constants are in BGR format!

   hDIB = FreeImage_CreateFromOlePicture(Picture)
   If (hDIB) Then
      Select Case FreeImage_GetBPP(hDIB)
      
      Case 1, 8, 24, 32
         hDIBNew = FreeImage_Rotate(hDIB, Angle, ByVal ColorPtr)
         Set FreeImage_RotateIOP = FreeImage_GetOlePicture(hDIBNew, , True)
         
      End Select
      Call FreeImage_Unload(hDIB)
   End If

End Function

'--------------------------------------------------------------------------------
' Compression functions wrappers
'--------------------------------------------------------------------------------

Public Function FreeImage_ZLibCompressEx(ByRef target As Variant, _
                                Optional ByRef TargetSize As Long, _
                                Optional ByRef Source As Variant, _
                                Optional ByVal SourceSize As Long, _
                                Optional ByVal Offset As Long) As Long
                                
Dim lSourceDataPtr As Long
Dim lTargetDataPtr As Long
Dim bTargetCreated As Boolean

   ' This function is a more VB friendly wrapper for compressing data with
   ' the 'FreeImage_ZLibCompress' function.
   
   ' The parameter 'Target' may either be a VB style array of Byte, Integer
   ' or Long or a pointer to a memory block. If 'Target' is a pointer to a
   ' memory block (when it contains an address), 'TargetSize' must be
   ' specified and greater than zero. If 'Target' is an initialized array,
   ' the whole array will be used to store compressed data when 'TargetSize'
   ' is missing or below or equal to zero. If 'TargetSize' is specified, only
   ' the first TargetSize bytes of the array will be used.
   ' In each case, all rules according to the FreeImage API documentation
   ' apply, what means that the target buffer must be at least 0.1% greater
   ' than the source buffer plus 12 bytes.
   ' If 'Target' is an uninitialized array, the contents of 'TargetSize'
   ' will be ignored and the size of the array 'Target' will be handled
   ' internally. When the function returns, 'Target' will be initialized
   ' as an array of Byte and sized correctly to hold all the compressed
   ' data.
   
   ' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
   ' is also true for 'Source' and 'SourceSize', expect that 'Source' should
   ' never be an uninitialized array. In that case, the function returns
   ' immediately.
   
   ' The optional parameter 'Offset' may contain a number of bytes to remain
   ' untouched at the beginning of 'Target', when an uninitialized array is
   ' provided through 'Target'. When 'Target' is either a pointer or an
   ' initialized array, 'Offset' will be ignored. This parameter is currently
   ' used by 'FreeImage_ZLibCompressVB' to store the length of the uncompressed
   ' data at the first four bytes of 'Target'.

   
   ' get the pointer and the size in bytes of the source
   ' memory block
   lSourceDataPtr = pGetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
      If (lTargetDataPtr = 0) Then
         ' if 'Target' is a null pointer, we will initialized it as an array
         ' of bytes; here we will take 'Offset' into account
         ReDim target(SourceSize + Int(SourceSize * 0.1) + _
                      12 + Offset) As Byte
         ' get pointer and size in bytes (will never be a null pointer)
         lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
         ' adjust according to 'Offset'
         lTargetDataPtr = lTargetDataPtr + Offset
         TargetSize = TargetSize - Offset
         bTargetCreated = True
      End If
      
      ' compress source data
      FreeImage_ZLibCompressEx = FreeImage_ZLibCompress(lTargetDataPtr, _
                                                        TargetSize, _
                                                        lSourceDataPtr, _
                                                        SourceSize)
      
      ' the function returns the number of bytes needed to store the
      ' compressed data or zero on failure
      If (FreeImage_ZLibCompressEx) Then
         If (bTargetCreated) Then
            ' when we created the array, we need to adjust it's size
            ' according to the length of the compressed data
            ReDim Preserve target(FreeImage_ZLibCompressEx - 1 + Offset)
         End If
      End If
   End If
                                
End Function

Public Function FreeImage_ZLibUncompressEx(ByRef target As Variant, _
                                  Optional ByRef TargetSize As Long, _
                                  Optional ByRef Source As Variant, _
                                  Optional ByVal SourceSize As Long) As Long
                                
Dim lSourceDataPtr As Long
Dim lTargetDataPtr As Long

   ' This function is a more VB friendly wrapper for compressing data with
   ' the 'FreeImage_ZLibUncompress' function.
   
   ' The parameter 'Target' may either be a VB style array of Byte, Integer
   ' or Long or a pointer to a memory block. If 'Target' is a pointer to a
   ' memory block (when it contains an address), 'TargetSize' must be
   ' specified and greater than zero. If 'Target' is an initialized array,
   ' the whole array will be used to store uncompressed data when 'TargetSize'
   ' is missing or below or equal to zero. If 'TargetSize' is specified, only
   ' the first TargetSize bytes of the array will be used.
   ' In each case, all rules according to the FreeImage API documentation
   ' apply, what means that the target buffer must be at least as large, to
   ' hold all the uncompressed data.
   ' Unlike the function 'FreeImage_ZLibCompressEx', 'Target' can not be
   ' an uninitialized array, since the size of the uncompressed data can
   ' not be determined by the ZLib functions, but must be specified by a
   ' mechanism outside the FreeImage compression functions' scope.
   
   ' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
   ' is also true for 'Source' and 'SourceSize'.
   
   
   ' get the pointer and the size in bytes of the source
   ' memory block
   lSourceDataPtr = pGetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
      If (lTargetDataPtr) Then
         ' if we do not have a null pointer, uncompress the data
         FreeImage_ZLibUncompressEx = FreeImage_ZLibUncompress(lTargetDataPtr, _
                                                               TargetSize, _
                                                               lSourceDataPtr, _
                                                               SourceSize)
      End If
   End If
                                
End Function

Public Function FreeImage_ZLibGZipEx(ByRef target As Variant, _
                            Optional ByRef TargetSize As Long, _
                            Optional ByRef Source As Variant, _
                            Optional ByVal SourceSize As Long, _
                            Optional ByVal Offset As Long) As Long
                                
Dim lSourceDataPtr As Long
Dim lTargetDataPtr As Long
Dim bTargetCreated As Boolean

   ' This function is a more VB friendly wrapper for compressing data with
   ' the 'FreeImage_ZLibGZip' function.
   
   ' The parameter 'Target' may either be a VB style array of Byte, Integer
   ' or Long or a pointer to a memory block. If 'Target' is a pointer to a
   ' memory block (when it contains an address), 'TargetSize' must be
   ' specified and greater than zero. If 'Target' is an initialized array,
   ' the whole array will be used to store compressed data when 'TargetSize'
   ' is missing or below or equal to zero. If 'TargetSize' is specified, only
   ' the first TargetSize bytes of the array will be used.
   ' In each case, all rules according to the FreeImage API documentation
   ' apply, what means that the target buffer must be at least 0.1% greater
   ' than the source buffer plus 24 bytes.
   ' If 'Target' is an uninitialized array, the contents of 'TargetSize'
   ' will be ignored and the size of the array 'Target' will be handled
   ' internally. When the function returns, 'Target' will be initialized
   ' as an array of Byte and sized correctly to hold all the compressed
   ' data.
   
   ' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
   ' is also true for 'Source' and 'SourceSize', expect that 'Source' should
   ' never be an uninitialized array. In that case, the function returns
   ' immediately.
   
   ' The optional parameter 'Offset' may contain a number of bytes to remain
   ' untouched at the beginning of 'Target', when an uninitialized array is
   ' provided through 'Target'. When 'Target' is either a pointer or an
   ' initialized array, 'Offset' will be ignored. This parameter is currently
   ' used by 'FreeImage_ZLibGZipVB' to store the length of the uncompressed
   ' data at the first four bytes of 'Target'.

   
   ' get the pointer and the size in bytes of the source
   ' memory block
   lSourceDataPtr = pGetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
      If (lTargetDataPtr = 0) Then
         ' if 'Target' is a null pointer, we will initialized it as an array
         ' of bytes; here we will take 'Offset' into account
         ReDim target(SourceSize + Int(SourceSize * 0.1) + _
                      24 + Offset) As Byte
         ' get pointer and size in bytes (will never be a null pointer)
         lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
         ' adjust according to 'Offset'
         lTargetDataPtr = lTargetDataPtr + Offset
         TargetSize = TargetSize - Offset
         bTargetCreated = True
      End If
      
      ' compress source data
      FreeImage_ZLibGZipEx = FreeImage_ZLibGZip(lTargetDataPtr, _
                                                TargetSize, _
                                                lSourceDataPtr, _
                                                SourceSize)
      
      ' the function returns the number of bytes needed to store the
      ' compressed data or zero on failure
      If (FreeImage_ZLibGZipEx) Then
         If (bTargetCreated) Then
            ' when we created the array, we need to adjust it's size
            ' according to the length of the compressed data
            ReDim Preserve target(FreeImage_ZLibGZipEx - 1 + Offset)
         End If
      End If
   End If
                                
End Function

Public Function FreeImage_ZLibCRC32Ex(ByVal CRC As Long, _
                             Optional ByRef Source As Variant, _
                             Optional ByVal SourceSize As Long) As Long
                                
Dim lSourceDataPtr As Long

   ' This function is a more VB friendly wrapper for compressing data with
   ' the 'FreeImage_ZLibCRC32' function.
   
   ' The parameter 'Source' may either be a VB style array of Byte, Integer
   ' or Long or a pointer to a memory block. If 'Source' is a pointer to a
   ' memory block (when it contains an address), 'SourceSize' must be
   ' specified and greater than zero. If 'Source' is an initialized array,
   ' the whole array will be used to calculate the new CRC when 'SourceSize'
   ' is missing or below or equal to zero. If 'SourceSize' is specified, only
   ' the first SourceSize bytes of the array will be used.

   
   ' get the pointer and the size in bytes of the source
   ' memory block
   lSourceDataPtr = pGetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' if we do not have a null pointer, calculate the CRC including 'crc'
      FreeImage_ZLibCRC32Ex = FreeImage_ZLibCRC32(CRC, _
                                                  lSourceDataPtr, _
                                                  SourceSize)
   End If
                                
End Function

Public Function FreeImage_ZLibGUnzipEx(ByRef target As Variant, _
                              Optional ByRef TargetSize As Long, _
                              Optional ByRef Source As Variant, _
                              Optional ByVal SourceSize As Long) As Long
                                
Dim lSourceDataPtr As Long
Dim lTargetDataPtr As Long

   ' This function is a more VB friendly wrapper for compressing data with
   ' the 'FreeImage_ZLibGUnzip' function.
   
   ' The parameter 'Target' may either be a VB style array of Byte, Integer
   ' or Long or a pointer to a memory block. If 'Target' is a pointer to a
   ' memory block (when it contains an address), 'TargetSize' must be
   ' specified and greater than zero. If 'Target' is an initialized array,
   ' the whole array will be used to store uncompressed data when 'TargetSize'
   ' is missing or below or equal to zero. If 'TargetSize' is specified, only
   ' the first TargetSize bytes of the array will be used.
   ' In each case, all rules according to the FreeImage API documentation
   ' apply, what means that the target buffer must be at least as large, to
   ' hold all the uncompressed data.
   ' Unlike the function 'FreeImage_ZLibGZipEx', 'Target' can not be
   ' an uninitialized array, since the size of the uncompressed data can
   ' not be determined by the ZLib functions, but must be specified by a
   ' mechanism outside the FreeImage compression functions' scope.
   
   ' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
   ' is also true for 'Source' and 'SourceSize'.
   
   
   ' get the pointer and the size in bytes of the source
   ' memory block
   lSourceDataPtr = pGetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = pGetMemoryBlockPtrFromVariant(target, TargetSize)
      If (lTargetDataPtr) Then
         ' if we do not have a null pointer, uncompress the data
         FreeImage_ZLibGUnzipEx = FreeImage_ZLibGUnzip(lTargetDataPtr, _
                                                       TargetSize, _
                                                       lSourceDataPtr, _
                                                       SourceSize)
      End If
   End If
                                
End Function

Public Function FreeImage_ZLibCompressVB(ByRef Data() As Byte, _
                                Optional ByVal IncludeSize As Boolean = True) As Byte()
                                
Dim lOffset As Long
Dim lArrayDataPtr As Long

   ' This function is another, even more VB friendly wrapper for the FreeImage
   ' 'FreeImage_ZLibCompress' function, that uses the 'FreeImage_ZLibCompressEx'
   ' function. This function is very easy to use, since it deals only with VB
   ' style Byte arrays.
   
   ' The parameter 'Data()' is a Byte array, providing the uncompressed source
   ' data, that will be compressed.
   
   ' The optional parameter 'IncludeSize' determines whether the size of the
   ' uncompressed data should be stored in the first four bytes of the returned
   ' byte buffer containing the compressed data or not. When 'IncludeSize' is
   ' True, the size of the uncompressed source data will be stored. This works
   ' in conjunction with the corresponding 'FreeImage_ZLibUncompressVB' function.
   
   ' The function returns a VB style Byte array containing the compressed data.
   

   ' start population the memory block with compressed data
   ' at offset 4 bytes, when the unclompressed size should
   ' be included
   If (IncludeSize) Then
      lOffset = 4
   End If
   
   Call FreeImage_ZLibCompressEx(FreeImage_ZLibCompressVB, , Data, , lOffset)
                                 
   If (IncludeSize) Then
      ' get the pointer actual pointing to the array data of
      ' the Byte array 'FreeImage_ZLibCompressVB'
      lArrayDataPtr = pDeref(pDeref(VarPtrArray(FreeImage_ZLibCompressVB)) + 12)

      ' copy uncompressed size into the first 4 bytes
      Call CopyMemory(ByVal lArrayDataPtr, UBound(Data) + 1, 4)
   End If

End Function

Public Function FreeImage_ZLibUncompressVB(ByRef Data() As Byte, _
                                  Optional ByVal SizeIncluded As Boolean = True, _
                                  Optional ByVal SizeNeeded As Long) As Byte()

Dim abBuffer() As Byte

   ' This function is another, even more VB friendly wrapper for the FreeImage
   ' 'FreeImage_ZLibUncompress' function, that uses the 'FreeImage_ZLibUncompressEx'
   ' function. This function is very easy to use, since it deals only with VB
   ' style Byte arrays.
   
   ' The parameter 'Data()' is a Byte array, providing the compressed source
   ' data that will be uncompressed either withthe size of the uncompressed
   ' data included or not.
   
   ' When the optional parameter 'SizeIncluded' is True, the function assumes,
   ' that the first four bytes contain the size of the uncompressed data as a
   ' Long value. In that case, 'SizeNeeded' will be ignored.
   
   ' When the size of the uncompressed data is not included in the buffer 'Data()'
   ' containing the compressed data, the optional parameter 'SizeNeeded' must
   ' specify the size in bytes needed to hold all the uncompressed data.
   
   ' The function returns a VB style Byte array containing the uncompressed data.


   If (SizeIncluded) Then
      ' get uncompressed size from the first 4 bytes and allocate
      ' buffer accordingly
      Call CopyMemory(SizeNeeded, Data(0), 4)
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibUncompressEx(abBuffer, , VarPtr(Data(4)), UBound(Data) - 3)
      Call pSwap(VarPtrArray(FreeImage_ZLibUncompressVB), VarPtrArray(abBuffer))
   
   ElseIf (SizeNeeded) Then
      ' no size included in compressed data, so just forward the
      ' call to 'FreeImage_ZLibUncompressEx' and trust on SizeNeeded
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibUncompressEx(abBuffer, , Data)
      Call pSwap(VarPtrArray(FreeImage_ZLibUncompressVB), VarPtrArray(abBuffer))
   
   End If

End Function

Public Function FreeImage_ZLibGZipVB(ByRef Data() As Byte, _
                            Optional ByVal IncludeSize As Boolean = True) As Byte()
                                
Dim lOffset As Long
Dim lArrayDataPtr As Long

   ' This function is another, even more VB friendly wrapper for the FreeImage
   ' 'FreeImage_ZLibGZip' function, that uses the 'FreeImage_ZLibGZipEx'
   ' function. This function is very easy to use, since it deals only with VB
   ' style Byte arrays.
   
   ' The parameter 'Data()' is a Byte array, providing the uncompressed source
   ' data that will be compressed.
   
   ' The optional parameter 'IncludeSize' determines whether the size of the
   ' uncompressed data should be stored in the first four bytes of the returned
   ' byte buffer containing the compressed data or not. When 'IncludeSize' is
   ' True, the size of the uncompressed source data will be stored. This works
   ' in conjunction with the corresponding 'FreeImage_ZLibGUnzipVB' function.
   
   ' The function returns a VB style Byte array containing the compressed data.


   ' start population the memory block with compressed data
   ' at offset 4 bytes, when the unclompressed size should
   ' be included
   If (IncludeSize) Then
      lOffset = 4
   End If
   
   Call FreeImage_ZLibGZipEx(FreeImage_ZLibGZipVB, , Data, , lOffset)
                                 
   If (IncludeSize) Then
      ' get the pointer actual pointing to the array data of
      ' the Byte array 'FreeImage_ZLibCompressVB'
      lArrayDataPtr = pDeref(pDeref(VarPtrArray(FreeImage_ZLibGZipVB)) + 12)

      ' copy uncompressed size into the first 4 bytes
      Call CopyMemory(ByVal lArrayDataPtr, UBound(Data) + 1, 4)
   End If

End Function

Public Function FreeImage_ZLibGUnzipVB(ByRef Data() As Byte, _
                              Optional ByVal SizeIncluded As Boolean = True, _
                              Optional ByVal SizeNeeded As Long) As Byte()

Dim abBuffer() As Byte

   ' This function is another, even more VB friendly wrapper for the FreeImage
   ' 'FreeImage_ZLibGUnzip' function, that uses the 'FreeImage_ZLibGUnzipEx'
   ' function. This function is very easy to use, since it deals only with VB
   ' style Byte arrays.
   
   ' The parameter 'Data()' is a Byte array, providing the compressed source
   ' data that will be uncompressed either withthe size of the uncompressed
   ' data included or not.
   
   ' When the optional parameter 'SizeIncluded' is True, the function assumes,
   ' that the first four bytes contain the size of the uncompressed data as a
   ' Long value. In that case, 'SizeNeeded' will be ignored.
   
   ' When the size of the uncompressed data is not included in the buffer 'Data()'
   ' containing the compressed data, the optional parameter 'SizeNeeded' must
   ' specify the size in bytes needed to hold all the uncompressed data.
   
   ' The function returns a VB style Byte array containing the uncompressed data.


   If (SizeIncluded) Then
      ' get uncompressed size from the first 4 bytes and allocate
      ' buffer accordingly
      Call CopyMemory(SizeNeeded, Data(0), 4)
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibGUnzipEx(abBuffer, , VarPtr(Data(4)), UBound(Data) - 3)
      Call pSwap(VarPtrArray(FreeImage_ZLibGUnzipVB), VarPtrArray(abBuffer))
   
   ElseIf (SizeNeeded) Then
      ' no size included in compressed data, so just forward the
      ' call to 'FreeImage_ZLibUncompressEx' and trust on SizeNeeded
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibGUnzipEx(abBuffer, , Data)
      Call pSwap(VarPtrArray(FreeImage_ZLibGUnzipVB), VarPtrArray(abBuffer))
   
   End If

End Function


'--------------------------------------------------------------------------------
' Public functions to destroy custom safearrays
'--------------------------------------------------------------------------------

Public Function FreeImage_DestroyLockedArray(ByRef Data As Variant) As Long

Dim lpArrayPtr As Long

   ' This function destroys an array, that was self created with a custom
   ' array descriptor of type ('fFeatures' member) 'FADF_AUTO Or FADF_FIXEDSIZE'.
   ' Such arrays are returned by mostly all of the array-dealing wrapper
   ' functions. Since these should not destroy the actual array data, when
   ' going out of scope, they are craeted as 'FADF_FIXEDSIZE'.'
   
   ' So, VB sees them as fixed or temporarily locked, when you try to manipulate
   ' the array's dimensions. There will occur some strange effects, you should
   ' know about:
   
   ' 1. When trying to 'ReDim' the array, this run-time error will occur:
   '    Error #10, 'This array is fixed or temporarily locked'
   
   ' 2. When trying to assign another array to the array variable, this
   '    run-time error will occur:
   '    Error #13, 'Type mismatch'
   
   ' 3. The 'Erase' statement has no effect on the array
   
   ' Although VB clears up these arrays correctly, when the array variable
   ' goes out of scope, you have to destroy the array manually, when you want
   ' to reuse the array variable in current scope.
   
   ' For an example assume, that you want do walk all scanlines in an image:
   
   ' For i = 0 To FreeImage_GetHeight(Bitmap)
   '
   '    ' assign scanline-array to array variable
   '    abByte = FreeImage_GetScanLineEx(Bitmap, i)
   '
   '    ' do some work on it...
   '
   '    ' destroy the array (only the array, not the actual data)
   '    Call FreeImage_DestroyLockedArray(dbByte)
   ' Next i
   
   ' The function returns zero on success and any other value on failure
   
   ' !! Attention !!
   ' This function uses a Variant parameter for passing the array to be
   ' destroyed. Since VB does not allow to pass an array of non public
   ' structures through a Variant parameter, this function can not be used
   ' with arrays of cutom type.
   
   ' You will get this compiler error: "Only public user defined types defined
   ' in public object modules can be used as parameters or return types for
   ' public procedures of class modules or as fields of public user defined types"
   
   ' So, there is a function in the wrapper called 'FreeImage_DestroyLockedArrayByPtr'
   ' that takes a pointer to the array variable which can be used to work around
   ' that VB limitation and furthermore can be used for any of these self-created
   ' arrays. To get the array variable's pointer, a declared version of the
   ' VB 'VarPtr' function can be used which works for all types of arrays expect
   ' String arrays. Declare this function like this in your code:
   
   ' Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
         ByRef Ptr() As Any) As Long
         
   ' Then an array could be destroyed by calling the 'FreeImage_DestroyLockedArrayByPtr'
   ' function like this:
   
   ' lResult = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(MyLockedArray))
   
   ' Additionally there are some handy wrapper functions available, one for each
   ' commonly used structure in FreeImage like RGBTRIPLE, RGBQUAD, FICOMPLEX etc.
   
   
   ' Currently, these functions do return 'FADF_AUTO Or FADF_FIXEDSIZE' arrays
   ' that must be destroyed using this or any of it's derived functions:
   
   ' FreeImage_GetPaletteEx()           with FreeImage_DestroyLockedArrayRGBQUAD()
   ' FreeImage_GetPaletteLong()         with FreeImage_DestroyLockedArray()
   ' FreeImage_SaveToMemoryEx2()        with FreeImage_DestroyLockedArray()
   ' FreeImage_AcquireMemoryEx()        with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineEx()          with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineBITMAP8()     with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineBITMAP16()    with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineBITMAP24()    with FreeImage_DestroyLockedArrayRGBTRIPLE()
   ' FreeImage_GetScanLineBITMAP32()    with FreeImage_DestroyLockedArrayRGBQUAD()
   ' FreeImage_GetScanLineINT16()       with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineINT32()       with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineFLOAT()       with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineDOUBLE()      with FreeImage_DestroyLockedArray()
   ' FreeImage_GetScanLineCOMPLEX()     with FreeImage_DestroyLockedArrayFICOMPLEX()
   ' FreeImage_GetScanLineRGB16()       with FreeImage_DestroyLockedArrayFIRGB16()
   ' FreeImage_GetScanLineRGBA16()      with FreeImage_DestroyLockedArrayFIRGBA16()
   ' FreeImage_GetScanLineRGBF()        with FreeImage_DestroyLockedArrayFIRGBF()
   ' FreeImage_GetScanLineRGBAF()       with FreeImage_DestroyLockedArrayFIRGBAF()

   
   ' ensure, this is an array
   If (VarType(Data) And vbArray) Then
   
      ' data is a VB array, what means a SAFEARRAY in C/C++, that is
      ' passed through a ByRef Variant variable, that is a pointer to
      ' a VARIANTARG structure
      
      ' the VARIANTARG structure looks like this:
      
      ' typedef struct tagVARIANT VARIANTARG;
      ' struct tagVARIANT
      '     {
      '     Union
      '         {
      '         struct __tagVARIANT
      '             {
      '             VARTYPE vt;
      '             WORD wReserved1;
      '             WORD wReserved2;
      '             WORD wReserved3;
      '             Union
      '                 {
      '                 [...]
      '             SAFEARRAY *parray;    // used when not VT_BYREF
      '                 [...]
      '             SAFEARRAY **pparray;  // used when VT_BYREF
      '                 [...]
      
      ' the data element (SAFEARRAY) has an offset of 8, since VARTYPE
      ' and WORD both have a length of 2 bytes; the pointer to the
      ' VARIANTARG structure is the VarPtr of the Variant variable in VB
      
      ' getting the contents of the data element (in C/C++: *(data + 8))
      lpArrayPtr = pDeref(VarPtr(Data) + 8)
      
      ' call the 'FreeImage_DestroyLockedArrayByPtr' function to destroy
      ' the array properly
      Call FreeImage_DestroyLockedArrayByPtr(lpArrayPtr)
   Else
      
      FreeImage_DestroyLockedArray = -1
   End If

End Function

Public Function FreeImage_DestroyLockedArrayByPtr(ByVal ArrayPtr As Long) As Long

Dim lpSA As Long

   ' This function destroys a self-created array with a custom array
   ' descriptor by a pointer to the array variable.

   ' dereference the pointer once (in C/C++: *ArrayPtr)
   lpSA = pDeref(ArrayPtr)
   ' now 'lpSA' is a pointer to the actual SAFEARRAY structure
   ' and could be a null pointer when the array is not initialized
   ' then, we have nothing to do here but return (-1) to indicate
   ' an "error"
   If (lpSA) Then
      
      ' destroy the array descriptor
      Call SafeArrayDestroyDescriptor(lpSA)
      
      ' make 'lpSA' a null pointer, that is an uninitialized array;
      ' keep in mind, that we here use 'ArrayPtr' as a ByVal argument,
      ' since 'ArrayPtr' is a pointer to lpSA (the address of lpSA);
      ' we need to zero these four bytes, 'ArrayPtr' points to
      Call CopyMemory(ByVal ArrayPtr, 0&, 4)
   Else
      
      ' the array is already uninitialized, so return an "error" value
      FreeImage_DestroyLockedArrayByPtr = -1
   End If

End Function

Public Function FreeImage_DestroyLockedArrayRGBTRIPLE(ByRef Data() As RGBTRIPLE) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'RGBTRIPLE'.
   
   FreeImage_DestroyLockedArrayRGBTRIPLE = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayRGBQUAD(ByRef Data() As RGBQUAD) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'RGBQUAD'.

   FreeImage_DestroyLockedArrayRGBQUAD = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayFICOMPLEX(ByRef Data() As FICOMPLEX) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'FICOMPLEX'.

   FreeImage_DestroyLockedArrayFICOMPLEX = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayFIRGB16(ByRef Data() As FIRGB16) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'FIRGB16'.

   FreeImage_DestroyLockedArrayFIRGB16 = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayFIRGBA16(ByRef Data() As FIRGBA16) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'FIRGBA16'.

   FreeImage_DestroyLockedArrayFIRGBA16 = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayFIRGBF(ByRef Data() As FIRGBF) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'FIRGBF'.

   FreeImage_DestroyLockedArrayFIRGBF = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function

Public Function FreeImage_DestroyLockedArrayFIRGBAF(ByRef Data() As FIRGBAF) As Long

   ' This function is a thin wrapper for 'FreeImage_DestroyLockedArrayByPtr'
   ' for destroying arrays of type 'FIRGBAF'.

   FreeImage_DestroyLockedArrayFIRGBAF = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data))

End Function



'--------------------------------------------------------------------------------
' Private IOlePicture related helper functions
'--------------------------------------------------------------------------------

Private Function pGetIOlePictureFromContainer(ByRef Container As Object, _
                                     Optional ByVal IncludeDrawings As Boolean) As IPicture

   ' Returns a VB IOlePicture object (IPicture) from a VB image hosting control.
   ' See the inline documentation of function 'FreeImage_CreateFromImageContainer'
   ' for a detailed description of this helper function.

   If (Not Container Is Nothing) Then
      
      Select Case TypeName(Container)
      
      Case "PictureBox", "Form"
         If (IncludeDrawings) Then
            If (Not Container.AutoRedraw) Then
               Call Err.Raise(5, _
                              "MFreeImage", _
                              Error$(5) & vbCrLf & vbCrLf & _
                              "Custom drawings can only be included into the DIB when " & _
                              "the container's 'AutoRedraw' property is set to True.")
               Exit Function
            End If
            Set pGetIOlePictureFromContainer = Container.Image
         Else
            Set pGetIOlePictureFromContainer = Container.Picture
         End If
      
      Case Else
      
         Dim bHasPicture As Boolean
         Dim bHasImage As Boolean
         Dim bIsAutoRedraw As Boolean
         
         On Error Resume Next
         bHasPicture = (Container.Picture <> 0)
         bHasImage = (Container.Image <> 0)
         bIsAutoRedraw = Container.AutoRedraw
         On Error GoTo 0
         
         If ((IncludeDrawings) And _
             (bHasImage) And _
             (bIsAutoRedraw)) Then
            Set pGetIOlePictureFromContainer = Container.Image
         
         ElseIf (bHasPicture) Then
            Set pGetIOlePictureFromContainer = Container.Picture
            
         Else
            Call Err.Raise(5, _
                           "MFreeImage", _
                           Error$(5) & vbCrLf & vbCrLf & _
                           "Cannot create DIB from container control. Container " & _
                           "control has no 'Picture' property.")
         
         End If
      
      End Select
      
   End If

End Function



'--------------------------------------------------------------------------------
' Private image and color helper functions
'--------------------------------------------------------------------------------

Private Function pGetPreviousColorDepth(ByVal Bpp As Long) As Long

   ' This function returns the 'previous' color depth of a given
   ' color depth. Here, 'previous' means the next smaller color
   ' depth.
   
   Select Case Bpp
   
   Case 32
      pGetPreviousColorDepth = 24
   
   Case 24
      pGetPreviousColorDepth = 16
   
   Case 16
      pGetPreviousColorDepth = 15
   
   Case 15
      pGetPreviousColorDepth = 8
   
   Case 8
      pGetPreviousColorDepth = 4
      
   Case 4
      pGetPreviousColorDepth = 1
   
   End Select
   
End Function

Private Function pGetNextColorDepth(ByVal Bpp As Long) As Long

   ' This function returns the 'next' color depth of a given
   ' color depth. Here, 'next' means the next greater color
   ' depth.
   
   Select Case Bpp
   
   Case 1
      pGetNextColorDepth = 4
      
   Case 4
      pGetNextColorDepth = 8
      
   Case 8
      pGetNextColorDepth = 15
      
   Case 15
      pGetNextColorDepth = 16
      
   Case 16
      pGetNextColorDepth = 24
      
   Case 24
      pGetNextColorDepth = 32
      
   End Select
   
End Function

'--------------------------------------------------------------------------------
' Private pointer manipulation helper functions
'--------------------------------------------------------------------------------

'Edited by Tanner: the old function was wasteful; this is simpler
Private Function pGetStringFromPointerA(ByRef Ptr As Long) As String
    Dim cUnicode As pdUnicode
    Set cUnicode = New pdUnicode
    pGetStringFromPointerA = cUnicode.ConvertCharPointerToVBString(Ptr, False)
End Function

Private Function pDeref(ByVal Ptr As Long) As Long

   ' This function dereferences a pointer and returns the
   ' contents as it's return value.
   
   ' in C/C++ this would be:
   ' return *(ptr);
   
   Call CopyMemory(pDeref, ByVal Ptr, 4)

End Function

Private Sub pSwap(ByVal lpSrc As Long, _
                  ByVal lpDst As Long)

Dim lpTmp As Long

   ' This function swaps two DWORD memory blocks pointed to
   ' by lpSrc and lpDst, whereby lpSrc and lpDst are actually
   ' no pointer types but contain the pointer's address.
   
   ' in C/C++ this would be:
   ' void pSwap(int lpSrc, int lpDst) {
   '   int tmp = *(int*)lpSrc;
   '   *(int*)lpSrc = *(int*)lpDst;
   '   *(int*)lpDst = tmp;
   ' }
  
   Call CopyMemory(lpTmp, ByVal lpSrc, 4)
   Call CopyMemory(ByVal lpSrc, ByVal lpDst, 4)
   Call CopyMemory(ByVal lpDst, lpTmp, 4)

End Sub

Private Function pGetMemoryBlockPtrFromVariant(ByRef Data As Variant, _
                                      Optional ByRef SizeInBytes As Long, _
                                      Optional ByRef ElementSize As Long) As Long
                                            
   ' This function returns the pointer to the memory block provided through
   ' the Variant parameter 'data', which could be either a Byte, Integer or
   ' Long array or the address of the memory block itself. In the last case,
   ' the parameter 'SizeInBytes' must not be omitted or zero, since it's
   ' correct value (the size of the memory block) can not be determined by
   ' the address only. So, the function fails, if 'SizeInBytes' is omitted
   ' or zero and 'data' is not an array but contains a Long value (the address
   ' of a memory block) by returning Null.
   
   ' If 'data' contains either a Byte, Integer or Long array, the pointer to
   ' the actual array data is returned. The parameter 'SizeInBytes' will
   ' be adjusted correctly, if it was less or equal zero upon entry.
   
   ' The function returns Null (zero) if there was no supported memory block
   ' provided.
   
   ' do we have an array?
   If (VarType(Data) And vbArray) Then
      Select Case (VarType(Data) And (Not vbArray))
      
      Case vbByte
         ElementSize = 1
         pGetMemoryBlockPtrFromVariant = pGetArrayPtrFromVariantArray(Data)
         If (pGetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1)
            
            ElseIf (SizeInBytes > (UBound(Data) + 1)) Then
               SizeInBytes = (UBound(Data) + 1)
            
            End If
         End If
      
      Case vbInteger
         ElementSize = 2
         pGetMemoryBlockPtrFromVariant = pGetArrayPtrFromVariantArray(Data)
         If (pGetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1) * 2
            
            ElseIf (SizeInBytes > ((UBound(Data) + 1) * 2)) Then
               SizeInBytes = (UBound(Data) + 1) * 2
            
            End If
         End If
      
      Case vbLong
         ElementSize = 4
         pGetMemoryBlockPtrFromVariant = pGetArrayPtrFromVariantArray(Data)
         If (pGetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1) * 4
            
            ElseIf (SizeInBytes > ((UBound(Data) + 1) * 4)) Then
               SizeInBytes = (UBound(Data) + 1) * 4
            
            End If
         End If
      
      End Select
   Else
      ElementSize = 1
      If ((VarType(Data) = vbLong) And _
          (SizeInBytes >= 0)) Then
         pGetMemoryBlockPtrFromVariant = Data
      End If
   End If
                                            
End Function

Private Function pGetArrayPtrFromVariantArray(ByRef Data As Variant) As Long

Dim eVarType As VbVarType
Dim lDataPtr As Long

   ' This function returns a pointer to the first array element of
   ' a VB array (SAFEARRAY) that is passed through a Variant type
   ' parameter. (Don't try this at home...)
   
   ' cache VarType in variable
   eVarType = VarType(Data)
   
   ' ensure, this is an array
   If (eVarType And vbArray) Then
      
      ' data is a VB array, what means a SAFEARRAY in C/C++, that is
      ' passed through a ByRef Variant variable, that is a pointer to
      ' a VARIANTARG structure
      
      ' the VARIANTARG structure looks like this:
      
      ' typedef struct tagVARIANT VARIANTARG;
      ' struct tagVARIANT
      '     {
      '     Union
      '         {
      '         struct __tagVARIANT
      '             {
      '             VARTYPE vt;
      '             WORD wReserved1;
      '             WORD wReserved2;
      '             WORD wReserved3;
      '             Union
      '                 {
      '                 [...]
      '             SAFEARRAY *parray;    // used when not VT_BYREF
      '                 [...]
      '             SAFEARRAY **pparray;  // used when VT_BYREF
      '                 [...]
      
      ' the data element (SAFEARRAY) has an offset of 8, since VARTYPE
      ' and WORD both have a length of 2 bytes; the pointer to the
      ' VARIANTARG structure is the VarPtr of the Variant variable in VB
      
      ' getting the contents of the data element (in C/C++: *(data + 8))
      lDataPtr = pDeref(VarPtr(Data) + 8)
      
      ' dereference the pointer again (in C/C++: *(lDataPtr))
      lDataPtr = pDeref(lDataPtr)
      
      ' test, whether 'lDataPtr' now is a Null pointer
      ' in that case, the array is not yet initialized and so we can't dereference
      ' it another time since we have no permisson to acces address 0
      
      ' the contents of 'lDataPtr' may be Null now in case of an uninitialized
      ' array; then we can't access any of the SAFEARRAY members since the array
      ' variable doesn't event point to a SAFEARRAY structure, so we will return
      ' the null pointer
      
      If (lDataPtr) Then
         ' the contents of lDataPtr now is a pointer to the SAFEARRAY structure
            
         ' the SAFEARRAY structure looks like this:
         
         ' typedef struct FARSTRUCT tagSAFEARRAY {
         '    unsigned short cDims;       // Count of dimensions in this array.
         '    unsigned short fFeatures;   // Flags used by the SafeArray
         '                                // routines documented below.
         ' #if defined(WIN32)
         '    unsigned long cbElements;   // Size of an element of the array.
         '                                // Does not include size of
         '                                // pointed-to data.
         '    unsigned long cLocks;       // Number of times the array has been
         '                                // locked without corresponding unlock.
         ' #Else
         '    unsigned short cbElements;
         '    unsigned short cLocks;
         '    unsigned long handle;       // Used on Macintosh only.
         ' #End If
         '    void HUGEP* pvData;               // Pointer to the data.
         '    SAFEARRAYBOUND rgsabound[1];      // One bound for each dimension.
         ' } SAFEARRAY;
         
         ' since we live in WIN32, the pvData element has an offset
         ' of 12 bytes from the base address of the structure,
         ' so dereference the pvData pointer, what indeed is a pointer
         ' to the actual array (in C/C++: *(lDataPtr + 12))
         lDataPtr = pDeref(lDataPtr + 12)
      End If
      
      ' return this value
      pGetArrayPtrFromVariantArray = lDataPtr
      
      ' a more shorter form of this function would be:
      ' (doesn't work for uninitialized arrays, but will likely crash!)
      'pGetArrayPtrFromVariantArray = pDeref(pDeref(pDeref(VarPtr(data) + 8)) + 12)
   End If

End Function

'--------------------------------------------------------------------------------
' Error handling functions
'--------------------------------------------------------------------------------

Public Sub FreeImage_InitErrorHandler()

   ' Call this function once for using the FreeImage 3 error handling callback.
   ' The 'FreeImage_ErrorHandler' function is called on each FreeImage 3 error.

   Call FreeImage_SetOutputMessage(AddressOf FreeImage_ErrorHandler)
   
   'Initialize the error message array
   ReDim g_FreeImageErrorMessages(0) As String

End Sub

Private Sub FreeImage_ErrorHandler(ByVal Format As FREE_IMAGE_FORMAT, ByVal Message As Long)

    Dim strErrorMessage As String
    Dim strImageFormat As String

   ' This function is called whenever the FreeImage 3 libraray throws an error.
   ' Currently this function gets the error message and the format name of the
   ' involved image type as VB string and prints both to the VB Debug console. Feel
   ' free to modify this function to call an error handling routine of your own.

   strErrorMessage = Trim$(pGetStringFromPointerA(Message))
   strImageFormat = FreeImage_GetFormatFromFIF(Format)
   
    'Save a copy of the FreeImage error in a public string, where other functions can retrieve it
    If Len(g_FreeImageErrorMessages(UBound(g_FreeImageErrorMessages))) <> 0 Then
        
        'See if this error already exists in the log
        Dim errorFound As Boolean
        errorFound = False
        
        Dim i As Long
        For i = 0 To UBound(g_FreeImageErrorMessages)
            If StrComp(g_FreeImageErrorMessages(i), strErrorMessage, vbTextCompare) = 0 Then
                errorFound = True
                Exit For
            End If
        Next i
        
        'If the error was not found in the log, add it now
        If Not errorFound Then
            ReDim Preserve g_FreeImageErrorMessages(0 To UBound(g_FreeImageErrorMessages) + 1) As String
            g_FreeImageErrorMessages(UBound(g_FreeImageErrorMessages)) = Trim$(strErrorMessage)
        End If
    Else
        g_FreeImageErrorMessages(UBound(g_FreeImageErrorMessages)) = Trim$(strErrorMessage)
    End If
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "FreeImage returned the following internal error:", PDM_EXTERNAL_LIB
        pdDebug.LogAction vbTab & strErrorMessage, PDM_EXTERNAL_LIB
        pdDebug.LogAction vbTab & "Image format in question was: " & strImageFormat, PDM_EXTERNAL_LIB
    #End If
   
End Sub
