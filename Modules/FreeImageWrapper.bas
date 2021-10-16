Attribute VB_Name = "Outside_FreeImageV3"
'Note: this file has been heavily modified for use within PhotoDemon.

'The vast majority of the code is copied directly from the official FreeImage VB6 wrapper
' by Carsten Klein, but I have stripped out unused functions, retyped certain enums
' (to work more nicely with PD's custom systems), and directly modified many functions to
' handle data more easily for my own purposes.

'Said another way: IF YOU WANT TO USE THIS CODE IN YOUR OWN PROJECT, PLEASE DOWNLOAD AN
' ORIGINAL COPY FROM THIS LINK (good as of August 2020):
' http://freeimage.sourceforge.net/download.html

'Thank you to Carsten Klein and the FreeImage team for their excellent library and VB6 wrapper.


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
' versions calles LoadPictureEx() and SavePictureEx() offering the FreeImage 3's
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

' NOTE FROM TANNER: a very detailed changelog follows this line in the original, but it has been removed for brevity's sake

'--------------------------------------------------------------------------------
' Win32 API function, struct and constant declarations
'--------------------------------------------------------------------------------

Private Const ERROR_SUCCESS As Long = 0
    
Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32.dll" ( _
    ByVal cDims As Long, _
    ByRef ppsaOut As Long) As Long
    
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

Private Type Bitmap_API
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
    
'GDI32
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
    
Private Declare Function GetCurrentObject Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    ByVal uObjectType As Long) As Long

Private Const OBJ_BITMAP As Long = 7
Private Const DIB_RGB_COLORS As Long = 0

'--------------------------------------------------------------------------------
' FreeImage 3 types, constants and enumerations
'--------------------------------------------------------------------------------

'FREEIMAGE

' Version information
'Private Const FREEIMAGE_MAJOR_VERSION As Long = 3
'Private Const FREEIMAGE_MINOR_VERSION As Long = 18
'Private Const FREEIMAGE_RELEASE_SERIAL As Long = 0

' Load / Save flag constants
Public Const FIF_LOAD_NOPIXELS = &H8000              ' load the image header only (not supported by all plugins)

Public Const BMP_DEFAULT As Long = 0
Public Const BMP_SAVE_RLE As Long = 1
Public Const EXR_DEFAULT As Long = 0                 ' save data as half with piz-based wavelet compression
Public Const EXR_FLOAT As Long = &H1                 ' save data as float instead of as half (not recommended)
Public Const EXR_NONE As Long = &H2                  ' save with no compression
Public Const EXR_ZIP As Long = &H4                   ' save with zlib compression, in blocks of 16 scan lines
Public Const EXR_PIZ As Long = &H8                   ' save with piz-based wavelet compression
Public Const EXR_PXR24 As Long = &H10                ' save with lossy 24-bit float compression
Public Const EXR_B44 As Long = &H20                  ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
Public Const EXR_LC As Long = &H40                   ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
Public Const GIF_DEFAULT As Long = 0
Public Const GIF_LOAD256 As Long = 1                 ' Load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
Public Const GIF_PLAYBACK As Long = 2                ''Play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
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
Public Const JXR_LOSSLESS As Long = &H64             ' save lossless
Public Const JXR_PROGRESSIVE As Long = &H2000        ' save as a progressive-JXR (use | to combine with other save flags)
Public Const PCD_DEFAULT As Long = 0
Public Const PCD_BASE As Long = 1                    ' load the bitmap sized 768 x 512
Public Const PCD_BASEDIV4 As Long = 2                ' load the bitmap sized 384 x 256
Public Const PCD_BASEDIV16 As Long = 3               ' load the bitmap sized 192 x 128
Public Const PNM_DEFAULT As Long = 0
Public Const PNM_SAVE_RAW As Long = 0                ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
Public Const PNM_SAVE_ASCII As Long = 1              ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
Public Const RAW_DEFAULT As Long = 0                 ' load the file as linear RGB 48-bit
Public Const RAW_PREVIEW As Long = 1                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
Public Const RAW_DISPLAY As Long = 2                 ' load the file as RGB 24-bit
Public Const RAW_HALFSIZE As Long = 4                ' load the file as half-size color image
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
Public Const WEBP_LOSSLESS As Long = &H100           ' save in lossless mode

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
   Private Const FIF_UNKNOWN = -1, FIF_BMP = 0, FIF_ICO = 1, FIF_JPEG = 2, FIF_JNG = 3, FIF_KOALA = 4, FIF_LBM = 5, FIF_IFF = FIF_LBM, FIF_MNG = 6, FIF_PBM = 7, FIF_PBMRAW = 8, FIF_PCD = 9
   Private Const FIF_PCX = 10, FIF_PGM = 11, FIF_PGMRAW = 12, FIF_PNG = 13, FIF_PPM = 14, FIF_PPMRAW = 15, FIF_RAS = 16, FIF_TARGA = 17, FIF_TIFF = 18, FIF_WBMP = 19
   Private Const FIF_PSD = 20, FIF_CUT = 21, FIF_XBM = 22, FIF_XPM = 23, FIF_DDS = 24, FIF_GIF = 25, FIF_HDR = 26, FIF_FAXG3 = 27, FIF_SGI = 28, FIF_EXR = 29
   Private Const FIF_J2K = 30, FIF_JP2 = 31, FIF_PFM = 32, FIF_PICT = 33, FIF_RAW = 34, FIF_WEBP = 35, FIF_JXR = 36
#End If

Public Enum FREE_IMAGE_LOAD_OPTIONS
   FILO_LOAD_NOPIXELS = FIF_LOAD_NOPIXELS         ' load the image header only (not supported by all plugins)
   FILO_LOAD_DEFAULT = 0
   FILO_GIF_DEFAULT = GIF_DEFAULT
   FILO_GIF_LOAD256 = GIF_LOAD256                 ' load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
   FILO_GIF_PLAYBACK = GIF_PLAYBACK               ' 'play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
   FILO_JPEG_DEFAULT = JPEG_DEFAULT               ' for loading this is a synonym for FILO_JPEG_FAST
   FILO_JPEG_FAST = JPEG_FAST                     ' load the file as fast as possible, sacrificing some quality
   FILO_JPEG_ACCURATE = JPEG_ACCURATE             ' load the file with the best quality, sacrificing some speed
   FILO_JPEG_CMYK = JPEG_CMYK                     ' load separated CMYK "as is" (use 'OR' to combine with other load flags)
   FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE         ' load and rotate according to Exif 'Orientation' tag if available
   FILO_JPEG_GREYSCALE = JPEG_GREYSCALE           ' load and convert to a 8-bit greyscale image
   FILO_PCD_DEFAULT = PCD_DEFAULT
   FILO_PCD_BASE = PCD_BASE                       ' load the bitmap sized 768 x 512
   FILO_PCD_BASEDIV4 = PCD_BASEDIV4               ' load the bitmap sized 384 x 256
   FILO_PCD_BASEDIV16 = PCD_BASEDIV16             ' load the bitmap sized 192 x 128
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
   Const FILO_JPEG_DEFAULT = JPEG_DEFAULT
   Const FILO_JPEG_FAST = JPEG_FAST
   Const FILO_JPEG_ACCURATE = JPEG_ACCURATE
   Const FILO_JPEG_CMYK = JPEG_CMYK
   Const FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE
   Const FILO_PCD_DEFAULT = PCD_DEFAULT
   Const FILO_PCD_BASE = PCD_BASE
   Const FILO_PCD_BASEDIV4 = PCD_BASEDIV4
   Const FILO_PCD_BASEDIV16 = PCD_BASEDIV16
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
   FIQ_LFPQUANT = 2          ' Lossless Fast Pseudo-Quantization Algorithm by Carsten Klein
End Enum
#If False Then
   Const FIQ_WUQUANT = 0, FIQ_NNQUANT = 1, FIQ_LFPQUANT = 2
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
   Private Const FIMD_NODATA = -1, FIMD_COMMENTS = 0, FIMD_EXIF_MAIN = 1, FIMD_EXIF_EXIF = 2, FIMD_EXIF_GPS = 3, FIMD_EXIF_MAKERNOTE = 4, FIMD_EXIF_INTEROP = 5, FIMD_IPTC = 6, FIMD_XMP = 7, FIMD_GEOTIFF = 8, FIMD_ANIMATION = 9
   Private Const FIMD_CUSTOM = 10, FIMD_EXIF_RAW = 11
#End If

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

Public Type FIICCPROFILE
   Flags As Integer
   Size As Long
   Data As Long
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
   numerator As Variant
   denominator As Variant
End Type

Public Type FREE_IMAGE_TAG
   Model As FREE_IMAGE_MDMODEL
   tagPtr As Long
   Key As String
   Description As String
   Id As Long
   Type As FREE_IMAGE_MDTYPE
   Count As Long
   Length As Long
   StringValue As String
   Palette() As RGBQuad
   RationalValue() As FIRATIONAL
   Value As Variant
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
Private Declare Function FreeImage_GetVersionInt Lib "FreeImage.dll" Alias "_FreeImage_GetVersion@0" () As Long

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
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_LoadUInt Lib "FreeImage.dll" Alias "_FreeImage_LoadU@12" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal srcFilename As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveUInt Lib "FreeImage.dll" Alias "_FreeImage_SaveU@16" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal Bitmap As Long, _
           ByVal srcFilename As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
  
Public Declare Sub FreeImage_Unload Lib "FreeImage.dll" Alias "_FreeImage_Unload@4" ( _
           ByVal Bitmap As Long)


' Bitmap information functions
Public Declare Function FreeImage_GetImageType Lib "FreeImage.dll" Alias "_FreeImage_GetImageType@4" ( _
           ByVal Bitmap As Long) As FREE_IMAGE_TYPE

Public Declare Function FreeImage_GetColorsUsed Lib "FreeImage.dll" Alias "_FreeImage_GetColorsUsed@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetBPP Lib "FreeImage.dll" Alias "_FreeImage_GetBPP@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetWidth Lib "FreeImage.dll" Alias "_FreeImage_GetWidth@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetHeight Lib "FreeImage.dll" Alias "_FreeImage_GetHeight@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetPitch Lib "FreeImage.dll" Alias "_FreeImage_GetPitch@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterX@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterY@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Sub FreeImage_SetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterX@8" ( _
           ByVal Bitmap As Long, _
           ByVal Resolution As Long)

Public Declare Sub FreeImage_SetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterY@8" ( _
           ByVal Bitmap As Long, _
           ByVal Resolution As Long)

Public Declare Function FreeImage_GetInfoHeader Lib "FreeImage.dll" Alias "_FreeImage_GetInfoHeader@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetInfo Lib "FreeImage.dll" Alias "_FreeImage_GetInfo@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetColorType Lib "FreeImage.dll" Alias "_FreeImage_GetColorType@4" ( _
           ByVal Bitmap As Long) As FREE_IMAGE_COLOR_TYPE

Public Declare Function FreeImage_GetTransparencyCount Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyCount@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_Invert Lib "FreeImage.dll" Alias "_FreeImage_Invert@4" (ByVal Bitmap As Long) As Long

Private Declare Function FreeImage_IsTransparentInt Lib "FreeImage.dll" Alias "_FreeImage_IsTransparent@4" ( _
           ByVal Bitmap As Long) As Long
           
Public Declare Function FreeImage_GetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_GetTransparentIndex@4" ( _
           ByVal Bitmap As Long) As Long
           
Public Declare Function FreeImage_SetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_SetTransparentIndex@8" ( _
           ByVal Bitmap As Long, _
           ByVal Index As Long) As Long

Private Declare Function FreeImage_HasBackgroundColorInt Lib "FreeImage.dll" Alias "_FreeImage_HasBackgroundColor@4" ( _
           ByVal Bitmap As Long) As Long
           
'Public Declare Function FreeImage_GetThumbnail Lib "FreeImage.dll" Alias "_FreeImage_GetThumbnail@4" ( _
           ByVal Bitmap As Long) As Long
           
' Filetype functions
Public Declare Function FreeImage_GetFileTypeU Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeU@8" ( _
           ByVal srcFilename As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_GetFileTypeFromMemory Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromMemory@8" ( _
           ByVal Stream As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT


' Pixel access functions
Public Declare Function FreeImage_GetBits Lib "FreeImage.dll" Alias "_FreeImage_GetBits@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetScanline Lib "FreeImage.dll" Alias "_FreeImage_GetScanLine@8" ( _
           ByVal Bitmap As Long, _
           ByVal Scanline As Long) As Long
        
        
' Conversion functions
Public Declare Function FreeImage_ConvertTo4Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo4Bits@4" ( _
           ByVal Bitmap As Long) As Long
           
Public Declare Function FreeImage_ConvertToGreyscale Lib "FreeImage.dll" Alias "_FreeImage_ConvertToGreyscale@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits555@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits565@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo24Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo24Bits@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo32Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo32Bits@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ColorQuantize Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantize@8" ( _
           ByVal Bitmap As Long, _
           ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE) As Long
           
Public Declare Function FreeImage_ColorQuantizeExInt Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantizeEx@20" ( _
           ByVal Bitmap As Long, _
  Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, _
  Optional ByVal paletteSize As Long = 256, _
  Optional ByVal ReserveSize As Long = 0, _
  Optional ByVal ReservePalettePtr As Long = 0) As Long

Public Declare Function FreeImage_Threshold Lib "FreeImage.dll" Alias "_FreeImage_Threshold@8" ( _
           ByVal Bitmap As Long, _
           ByVal threshold As Byte) As Long

Public Declare Function FreeImage_Dither Lib "FreeImage.dll" Alias "_FreeImage_Dither@8" ( _
           ByVal Bitmap As Long, _
           ByVal ditherMethod As FREE_IMAGE_DITHER) As Long

Private Declare Function FreeImage_ConvertFromRawBitsExInt Lib "FreeImage.dll" Alias "_FreeImage_ConvertFromRawBitsEx@44" ( _
           ByVal CopySource As Long, _
           ByVal BitsPtr As Long, _
           ByVal ImageType As FREE_IMAGE_TYPE, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal Pitch As Long, _
           ByVal BitsPerPixel As Long, _
           ByVal RedMask As Long, _
           ByVal GreenMask As Long, _
           ByVal BlueMask As Long, _
           ByVal TopDown As Long) As Long

Public Declare Function FreeImage_ConvertToFloat Lib "FreeImage.dll" Alias "_FreeImage_ConvertToFloat@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGBF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBF@4" ( _
           ByVal Bitmap As Long) As Long

'Manually patched by Tanner:
Public Declare Function FreeImage_ConvertToRGBAF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBAF@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertToUINT16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToUINT16@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGB16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGB16@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGBA16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBA16@4" ( _
           ByVal Bitmap As Long) As Long
           
Public Declare Function FreeImage_GetRedMask Lib "FreeImage.dll" Alias "_FreeImage_GetRedMask@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Function FreeImage_GetBlueMask Lib "FreeImage.dll" Alias "_FreeImage_GetBlueMask@4" ( _
           ByVal Bitmap As Long) As Long
           
'Red and blue masks are used to determine RGB vs BGR order.  Green masks aren't used at present.
'Public Declare Function FreeImage_GetGreenMask Lib "FreeImage.dll" Alias "_FreeImage_GetGreenMask@4" ( _
           ByVal Bitmap As Long) As Long

' Tone mapping operators
Public Declare Function FreeImage_TmoDrago03 Lib "FreeImage.dll" Alias "_FreeImage_TmoDrago03@20" ( _
           ByVal Bitmap As Long, _
  Optional ByVal gamma As Double = 2.2, _
  Optional ByVal Exposure As Double) As Long

Public Declare Function FreeImage_TmoReinhard05Ex Lib "FreeImage.dll" Alias "_FreeImage_TmoReinhard05Ex@36" ( _
           ByVal Bitmap As Long, _
  Optional ByVal Intensity As Double, _
  Optional ByVal Contrast As Double, _
  Optional ByVal Adaptation As Double = 1, _
  Optional ByVal ColorCorrection As Double) As Long

' ICC profile functions
Private Declare Function FreeImage_GetICCProfileInt Lib "FreeImage.dll" Alias "_FreeImage_GetICCProfile@4" ( _
           ByVal Bitmap As Long) As Long

' Plugin functions
Private Declare Function FreeImage_GetFormatFromFIFInt Lib "FreeImage.dll" Alias "_FreeImage_GetFormatFromFIF@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Public Declare Function FreeImage_GetFIFFromFilenameU Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFilenameU@4" ( _
           ByVal srcFilename As Long) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_FIFSupportsReadingInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsReading@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsWritingInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsWriting@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsExportTypeInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportType@8" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal ImageType As FREE_IMAGE_TYPE) As Long

Private Declare Function FreeImage_FIFSupportsExportBPPInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportBPP@8" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal BitsPerPixel As Long) As Long

Private Declare Function FreeImage_FIFSupportsICCProfilesInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsICCProfiles@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long
           
Private Declare Function FreeImage_FIFSupportsNoPixelsInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsNoPixels@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long


' Multipage functions
Private Declare Function FreeImage_OpenMultiBitmapInt Lib "FreeImage.dll" Alias "_FreeImage_OpenMultiBitmap@24" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal srcFilename As String, _
           ByVal CreateNew As Long, _
           ByVal ReadOnly As Long, _
           ByVal KeepCacheInMemory As Long, _
           ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_CloseMultiBitmapInt Lib "FreeImage.dll" Alias "_FreeImage_CloseMultiBitmap@8" ( _
           ByVal Bitmap As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Public Declare Function FreeImage_GetPageCount Lib "FreeImage.dll" Alias "_FreeImage_GetPageCount@4" ( _
           ByVal Bitmap As Long) As Long

Public Declare Sub FreeImage_AppendPage Lib "FreeImage.dll" Alias "_FreeImage_AppendPage@8" ( _
           ByVal Bitmap As Long, _
           ByVal PageBitmap As Long)

Public Declare Function FreeImage_LockPage Lib "FreeImage.dll" Alias "_FreeImage_LockPage@8" ( _
           ByVal Bitmap As Long, _
           ByVal Page As Long) As Long

Private Declare Sub FreeImage_UnlockPageInt Lib "FreeImage.dll" Alias "_FreeImage_UnlockPage@12" ( _
           ByVal Bitmap As Long, _
           ByVal PageBitmap As Long, _
           ByVal ApplyChanges As Long)

' Memory I/O streams
Public Declare Function FreeImage_OpenMemoryByPtr Lib "FreeImage.dll" Alias "_FreeImage_OpenMemory@8" ( _
  Optional ByVal dataPtr As Long, _
  Optional ByVal sizeInBytes As Long) As Long

Public Declare Sub FreeImage_CloseMemory Lib "FreeImage.dll" Alias "_FreeImage_CloseMemory@4" ( _
           ByVal Stream As Long)

Public Declare Function FreeImage_LoadFromMemory Lib "FreeImage.dll" Alias "_FreeImage_LoadFromMemory@12" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveToMemoryInt Lib "FreeImage.dll" Alias "_FreeImage_SaveToMemory@16" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal Bitmap As Long, _
           ByVal Stream As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long

Private Declare Function FreeImage_AcquireMemoryInt Lib "FreeImage.dll" Alias "_FreeImage_AcquireMemory@12" ( _
           ByVal Stream As Long, _
           ByRef dataPtr As Long, _
           ByRef sizeInBytes As Long) As Long
           
' Compression functions
Public Declare Function FreeImage_ZLibUncompress Lib "FreeImage.dll" Alias "_FreeImage_ZLibUncompress@16" ( _
           ByVal targetPtr As Long, _
           ByVal TargetSize As Long, _
           ByVal SourcePtr As Long, _
           ByVal SourceSize As Long) As Long
           
'--------------------------------------------------------------------------------
' Metadata functions
'--------------------------------------------------------------------------------

' tag creation / destruction
Private Declare Function FreeImage_CreateTag Lib "FreeImage.dll" Alias "_FreeImage_CreateTag@0" () As Long

Private Declare Sub FreeImage_DeleteTag Lib "FreeImage.dll" Alias "_FreeImage_DeleteTag@4" ( _
           ByVal Tag As Long)

' tag getters and setters (only those actually needed by wrapper functions)
Private Declare Function FreeImage_SetTagKey Lib "FreeImage.dll" Alias "_FreeImage_SetTagKey@8" ( _
           ByVal Tag As Long, _
           ByVal Key As String) As Long

Private Declare Function FreeImage_SetTagValue Lib "FreeImage.dll" Alias "_FreeImage_SetTagValue@8" ( _
           ByVal Tag As Long, _
           ByVal ValuePtr As Long) As Long

'Added by Tanner
Private Declare Function FreeImage_SetTagID Lib "FreeImage.dll" Alias "_FreeImage_SetTagID@8" ( _
           ByVal Tag As Long, _
           ByVal newValue As Integer) As Long
           
Private Declare Function FreeImage_SetTagType Lib "FreeImage.dll" Alias "_FreeImage_SetTagType@8" ( _
           ByVal Tag As Long, _
           ByVal newValue As Integer) As Long

Private Declare Function FreeImage_SetTagCount Lib "FreeImage.dll" Alias "_FreeImage_SetTagCount@8" ( _
           ByVal Tag As Long, _
           ByVal newValue As Long) As Long
           
Private Declare Function FreeImage_SetTagLength Lib "FreeImage.dll" Alias "_FreeImage_SetTagLength@8" ( _
           ByVal Tag As Long, _
           ByVal newValue As Long) As Long

' metadata setters and getters
Public Declare Function FreeImage_SetMetadataInt Lib "FreeImage.dll" Alias "_FreeImage_SetMetadata@16" ( _
           ByVal Model As Long, _
           ByVal Bitmap As Long, _
           ByVal Key As String, _
           ByVal Tag As Long) As Long

Private Declare Function FreeImage_GetMetadataInt Lib "FreeImage.dll" Alias "_FreeImage_GetMetadata@16" ( _
           ByVal Model As Long, _
           ByVal Bitmap As Long, _
           ByVal Key As String, _
           ByRef Tag As Long) As Long

' tag to string conversion functions
Private Declare Function FreeImage_TagToStringInt Lib "FreeImage.dll" Alias "_FreeImage_TagToString@12" ( _
           ByVal Model As Long, _
           ByVal Tag As Long, _
  Optional ByVal Make As String = vbNullString) As Long


'--------------------------------------------------------------------------------
' Toolkit functions
'--------------------------------------------------------------------------------

' Rotating and flipping
' EDIT BY TANNER: this function was deprecated in the 3.18 release
'Public Declare Function FreeImage_RotateClassic Lib "FreeImage.dll" Alias "_FreeImage_RotateClassic@12" ( _
           ByVal Bitmap As Long, _
           ByVal Angle As Double) As Long

Public Declare Function FreeImage_Rotate Lib "FreeImage.dll" Alias "_FreeImage_Rotate@16" ( _
           ByVal Bitmap As Long, _
           ByVal Angle As Double, _
  Optional ByRef Color As Any = 0) As Long

Private Declare Function FreeImage_RotateExInt Lib "FreeImage.dll" Alias "_FreeImage_RotateEx@48" ( _
           ByVal Bitmap As Long, _
           ByVal Angle As Double, _
           ByVal ShiftX As Double, _
           ByVal ShiftY As Double, _
           ByVal originX As Double, _
           ByVal originY As Double, _
           ByVal UseMask As Long) As Long

Private Declare Function FreeImage_FlipHorizontal Lib "FreeImage.dll" Alias "_FreeImage_FlipHorizontal@4" (ByVal FIBITMAP As Long) As Long
Private Declare Function FreeImage_FlipVertical Lib "FreeImage.dll" Alias "_FreeImage_FlipVertical@4" (ByVal FIBITMAP As Long) As Long

' Upsampling and downsampling
Public Declare Function FreeImage_Rescale Lib "FreeImage.dll" Alias "_FreeImage_Rescale@16" ( _
           ByVal Bitmap As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal Filter As FREE_IMAGE_FILTER) As Long
           
Private Declare Function FreeImage_MakeThumbnailInt Lib "FreeImage.dll" Alias "_FreeImage_MakeThumbnail@12" ( _
           ByVal Bitmap As Long, _
           ByVal MaxPixelSize As Long, _
  Optional ByVal Convert As Long) As Long

' Copy / Paste / Composite routines
Public Declare Function FreeImage_Composite Lib "FreeImage.dll" Alias "_FreeImage_Composite@16" ( _
           ByVal Bitmap As Long, _
  Optional ByVal UseFileBackColor As Long, _
  Optional ByRef AppBackColor As Any, _
  Optional ByVal BackgroundBitmap As Long) As Long

Private Declare Function FreeImage_PreMultiplyWithAlphaInt Lib "FreeImage.dll" Alias "_FreeImage_PreMultiplyWithAlpha@4" ( _
           ByVal Bitmap As Long) As Long
           


'--------------------------------------------------------------------------------
' Line converting functions
'--------------------------------------------------------------------------------

' convert to 4 bpp
Public Declare Sub FreeImage_ConvertLine1To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To4@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine8To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)
           
Public Declare Sub FreeImage_ConvertLine16To4_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
                     
Public Declare Sub FreeImage_ConvertLine16To4_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine24To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           
Public Declare Sub FreeImage_ConvertLine32To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To4@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 8 bpp
Public Declare Sub FreeImage_ConvertLine1To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine4To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To8@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To8_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To8_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To8@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To8@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)
           

' convert to 16 bpp
Public Declare Sub FreeImage_ConvertLine1To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_555@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_555@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_555@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16_565_To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_565_To16_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine1To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_565@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_565@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_565@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16_555_To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_555_To16_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 24 bpp
Public Declare Sub FreeImage_ConvertLine1To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To24@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To24@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16To24_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To24_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine32To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To24@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)


' convert to 32 bpp
Public Declare Sub FreeImage_ConvertLine1To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To32@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine4To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To32@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine8To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To32@16" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long, _
           ByVal PalettePtr As Long)

Public Declare Sub FreeImage_ConvertLine16To32_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_555@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine16To32_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_565@12" ( _
           ByVal targetPtr As Long, _
           ByVal SourcePtr As Long, _
           ByVal WidthInPixels As Long)

Public Declare Sub FreeImage_ConvertLine24To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To32@12" ( _
           ByVal targetPtr As Long, _
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

Public Function FreeImage_GetFormatFromFIF(ByVal imgFormat As FREE_IMAGE_FORMAT) As String

   ' This function returns the result of the 'FreeImage_GetFormatFromFIF' function
   ' as VB String. Read paragraph 2 of the "General notes on implementation
   ' and design" section to learn more about that technique.
   
   ' The parameter 'Format' works according to the FreeImage 3 API documentation.
   
   FreeImage_GetFormatFromFIF = pGetStringFromPointerA(FreeImage_GetFormatFromFIFInt(imgFormat))

End Function

Public Sub FreeImage_GetInfoHeaderEx(ByVal Bitmap As Long, ByVal ptrToBitmapInfoHeader As Long)

Dim lpInfoHeader As Long

   ' This function is a wrapper around FreeImage_GetInfoHeader() and returns a fully
   ' populated BITMAPINFOHEADER structure for a given bitmap.

   lpInfoHeader = FreeImage_GetInfoHeader(Bitmap)
   
   If (lpInfoHeader) Then
      Call CopyMemory(ByVal ptrToBitmapInfoHeader, ByVal lpInfoHeader, 40&)
   End If

End Sub

'--------------------------------------------------------------------------------
' BOOL/Boolean returning functions wrappers
'--------------------------------------------------------------------------------

Public Function FreeImage_HasPixels(ByVal Bitmap As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_HasPixels = (FreeImage_HasPixelsInt(Bitmap) = 1)

End Function

Public Function FreeImage_Save(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                               ByVal Bitmap As Long, _
                               ByVal srcFilename As String, _
                      Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_Save = (FreeImage_SaveUInt(imgFormat, Bitmap, StrPtr(srcFilename), Flags) = 1)

End Function

' Thin wrapper function returning a real VB Boolean value
Public Function FreeImage_IsTransparent(ByVal Bitmap As Long) As Boolean
    FreeImage_IsTransparent = (FreeImage_IsTransparentInt(Bitmap) = 1)
End Function
           
' Thin wrapper function returning a real VB Boolean value
Public Function FreeImage_HasBackgroundColor(ByVal Bitmap As Long) As Boolean
    FreeImage_HasBackgroundColor = (FreeImage_HasBackgroundColorInt(Bitmap) = 1)
End Function

Public Function FreeImage_FIFSupportsReading(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsReading = (FreeImage_FIFSupportsReadingInt(imgFormat) = 1)

End Function

Public Function FreeImage_FIFSupportsWriting(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsWriting = (FreeImage_FIFSupportsWritingInt(imgFormat) = 1)
   
End Function

Public Function FreeImage_FIFSupportsExportType(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                                                ByVal ImageType As FREE_IMAGE_TYPE) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsExportType = (FreeImage_FIFSupportsExportTypeInt(imgFormat, ImageType) = 1)

End Function

Public Function FreeImage_FIFSupportsExportBPP(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                                               ByVal BitsPerPixel As Long) As Boolean
   
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsExportBPP = (FreeImage_FIFSupportsExportBPPInt(imgFormat, BitsPerPixel) = 1)
                                             
End Function

Public Function FreeImage_FIFSupportsICCProfiles(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsICCProfiles = (FreeImage_FIFSupportsICCProfilesInt(imgFormat) = 1)

End Function

Public Function FreeImage_FIFSupportsNoPixels(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_FIFSupportsNoPixels = (FreeImage_FIFSupportsNoPixelsInt(imgFormat) = 1)

End Function

Public Function FreeImage_CloseMultiBitmap(ByVal Bitmap As Long, _
                                  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean

   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_CloseMultiBitmap = (FreeImage_CloseMultiBitmapInt(Bitmap, Flags) = 1)

End Function

Public Function FreeImage_SaveToMemory(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                                       ByVal Bitmap As Long, _
                                       ByVal Stream As Long, _
                              Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean
                              
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_SaveToMemory = (FreeImage_SaveToMemoryInt(imgFormat, Bitmap, Stream, Flags) = 1)
  
End Function

Public Function FreeImage_AcquireMemory(ByVal Stream As Long, _
                                        ByRef dataPtr As Long, _
                                        ByRef sizeInBytes As Long) As Boolean
                                        
   ' Thin wrapper function returning a real VB Boolean value

   FreeImage_AcquireMemory = (FreeImage_AcquireMemoryInt(Stream, dataPtr, sizeInBytes) = 1)
           
End Function

Public Function FreeImage_PreMultiplyWithAlpha(ByVal Bitmap As Long) As Boolean

   ' Thin wrapper function returning a real VB Boolean value
   
   FreeImage_PreMultiplyWithAlpha = (FreeImage_PreMultiplyWithAlphaInt(Bitmap) = 1)

End Function


Public Function FreeImage_OpenMultiBitmap(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                                          ByVal srcFilename As String, _
                                 Optional ByVal CreateNew As Boolean, _
                                 Optional ByVal ReadOnly As Boolean, _
                                 Optional ByVal KeepCacheInMemory As Boolean, _
                                 Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

   FreeImage_OpenMultiBitmap = FreeImage_OpenMultiBitmapInt(imgFormat, srcFilename, IIf(CreateNew, 1, 0), _
         IIf(ReadOnly And Not CreateNew, 1, 0), IIf(KeepCacheInMemory, 1, 0), Flags)

End Function

Public Sub FreeImage_UnlockPage(ByVal Bitmap As Long, ByVal PageBitmap As Long, ByVal ApplyChanges As Boolean)

Dim lApplyChanges As Long

   If (ApplyChanges) Then
      lApplyChanges = 1
   End If
   Call FreeImage_UnlockPageInt(Bitmap, PageBitmap, lApplyChanges)

End Sub

Public Function FreeImage_RotateEx(ByVal Bitmap As Long, _
                                   ByVal Angle As Double, _
                          Optional ByVal ShiftX As Double, _
                          Optional ByVal ShiftY As Double, _
                          Optional ByVal originX As Double, _
                          Optional ByVal originY As Double, _
                          Optional ByVal UseMask As Boolean) As Long
    Dim lUseMask As Long
    If UseMask Then lUseMask = 1 Else lUseMask = 0
    FreeImage_RotateEx = FreeImage_RotateExInt(Bitmap, Angle, ShiftX, ShiftY, originX, originY, lUseMask)
End Function

Public Function FreeImage_MakeThumbnail(ByVal Bitmap As Long, _
                                        ByVal MaxPixelSize As Long, _
                               Optional ByVal Convert As Boolean) As Long
    Dim lConvert As Long
    If Convert Then lConvert = 1
    FreeImage_MakeThumbnail = FreeImage_MakeThumbnailInt(Bitmap, MaxPixelSize, lConvert)
End Function

'Added by Tanner on 16-Nov-15, from the official FreeImage wrapper.  Tweaked slightly to better match PD's intended use-case.
Public Function FreeImage_ConvertFromRawBitsEx(ByVal CopySource As Boolean, _
                                               ByVal BitsPtr As Long, _
                                               ByVal ImageType As FREE_IMAGE_TYPE, _
                                               ByVal Width As Long, _
                                               ByVal Height As Long, _
                                               ByVal Pitch As Long, _
                                               ByVal BitsPerPixel As Long, _
                                      Optional ByVal RedMask As Long, _
                                      Optional ByVal GreenMask As Long, _
                                      Optional ByVal BlueMask As Long, _
                                      Optional ByVal TopDown As Boolean = False) As Long
    
    'Convert incoming VB booleans to C-style booleans
    Dim lCopySource As Long, lTopDown As Long
    If CopySource Then lCopySource = 1
    If TopDown Then lTopDown = 1
    
    'Ask FreeImage to simply wrap the data, rather than copying it (depending on CopySource, obviously)
    FreeImage_ConvertFromRawBitsEx = FreeImage_ConvertFromRawBitsExInt(lCopySource, BitsPtr, ImageType, _
         Width, Height, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, lTopDown)

End Function


'--------------------------------------------------------------------------------
' Extended functions derived from FreeImage 3 functions usually dealing
' with arrays
'--------------------------------------------------------------------------------

' Extended version of FreeImage_Unload, which additionally sets the passed Bitmap handle
' to zero after unloading.
Public Sub FreeImage_UnloadEx(ByRef Bitmap As Long)
    If (Bitmap <> 0) Then FreeImage_Unload Bitmap
    Bitmap = 0
End Sub


' Memory and Stream functions

'NOTE: modified by Tanner to support direct pointer retrieval
Public Function FreeImage_LoadFromMemoryEx(ByRef Data As Variant, _
                                  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS = 0, _
                                  Optional ByRef sizeInBytes As Long, _
                                  Optional ByRef imgFormat As FREE_IMAGE_FORMAT = FIF_UNKNOWN, _
                                  Optional ByVal ptrToDataInstead As Long = 0) As Long

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
     Dim lDataPtr As Long
     If (ptrToDataInstead <> 0) Then
         lDataPtr = ptrToDataInstead
     Else
         lDataPtr = pGetMemoryBlockPtrFromVariant(Data, sizeInBytes)
     End If
    
    ' open the memory stream
    Dim hStream As Long
    hStream = FreeImage_OpenMemoryByPtr(lDataPtr, sizeInBytes)
    If (hStream <> 0) Then
        
        ' on success, detect image type
        If (imgFormat = FIF_UNKNOWN) Then
            imgFormat = FreeImage_GetFileTypeFromMemory(hStream)
            Debug.Print "FreeImage_LoadFromMemoryEx auto-detected format " & imgFormat
        End If
      
        ' load the image from memory stream only, if known image type
        If (imgFormat <> FIF_UNKNOWN) Then
            FreeImage_LoadFromMemoryEx = FreeImage_LoadFromMemory(imgFormat, hStream, Flags)
        End If
      
        ' close the memory stream when open
        FreeImage_CloseMemory hStream
        
    Else
        Debug.Print "Couldn't obtain hStream pointer in FreeImage_LoadFromMemoryEx; sorry!"
    End If

End Function

'Modified LoadFromMemory function, created while testing unpredictable FreeImage LoadFromMemory failures
Public Function FreeImage_LoadFromMemoryEx_Tanner(ByVal dataPtr As Long, ByVal sizeInBytes As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef fileFormat As FREE_IMAGE_FORMAT = FIF_UNKNOWN) As Long

    Dim hStream As Long
    
    'FreeImage_LoadFromMemoryEx routinely fails without explanation, and I'm hoping to find out why!

   ' get both pointer and size in bytes of the memory block provided
   ' through the Variant parameter 'data'.
   'lDataPtr = pGetMemoryBlockPtrFromVariant(Data, SizeInBytes)
   
   ' open the memory stream
   hStream = FreeImage_OpenMemoryByPtr(dataPtr, sizeInBytes)
   If (hStream) Then
   
      ' on success, detect image type
      If (fileFormat = FIF_UNKNOWN) Then fileFormat = FreeImage_GetFileTypeFromMemory(hStream)
      
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

Public Function FreeImage_SaveToMemoryEx(ByVal imgFormat As FREE_IMAGE_FORMAT, _
                                         ByVal Bitmap As Long, _
                                         ByRef Data() As Byte, _
                                Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, _
                                Optional ByVal UnloadSource As Boolean, _
                                Optional ByRef dstSizeInBytes As Long) As Boolean

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
   
   
   If Bitmap Then
   
      If (Not FreeImage_HasPixels(Bitmap)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to save a 'header-only' bitmap.")
      End If
   
      hStream = FreeImage_OpenMemoryByPtr(0&, 0&)
      If hStream Then
         FreeImage_SaveToMemoryEx = FreeImage_SaveToMemory(imgFormat, Bitmap, hStream, Flags)
         
         If FreeImage_SaveToMemoryEx Then
            If FreeImage_AcquireMemoryInt(hStream, lpData, lSizeInBytes) Then
                
                On Error Resume Next
                
                'Change by Tanner: return the size in bytes, and only allocate new memory as necessary.
                ' (This allows the caller to reuse allocations that may already exist.)
                dstSizeInBytes = lSizeInBytes
                If Not VBHacks.IsArrayInitialized(Data) Then
                    ReDim Data(lSizeInBytes - 1) As Byte
                Else
                    If UBound(Data) < (lSizeInBytes - 1) Then ReDim Data(0 To lSizeInBytes - 1)
                End If
               
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
            Debug.Print "FreeImage_SaveToMemoryEx failed."
         End If
         
         Call FreeImage_CloseMemory(hStream)
         
      Else
         FreeImage_SaveToMemoryEx = False
      End If
      
      If UnloadSource Then Call FreeImage_Unload(Bitmap)
      
   End If

End Function

Private Function FreeImage_UnsignedLong(ByVal Value As Long) As Variant

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

Private Function FreeImage_UnsignedShort(ByVal Value As Integer) As Variant

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


' Bitmap resolution functions

Public Function FreeImage_GetResolutionX(ByVal Bitmap As Long) As Double

   ' This function gets a DIB's resolution in X-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.
   
   FreeImage_GetResolutionX = (0.0254 * FreeImage_GetDotsPerMeterX(Bitmap))

End Function

Public Sub FreeImage_SetResolutionX(ByVal Bitmap As Long, ByVal Resolution As Double)

   ' This function sets a DIB's resolution in X-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   Call FreeImage_SetDotsPerMeterX(Bitmap, Int(Resolution / 0.0254 + 0.5))

End Sub

Public Function FreeImage_GetResolutionY(ByVal Bitmap As Long) As Double

   ' This function gets a DIB's resolution in Y-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   FreeImage_GetResolutionY = (0.0254 * FreeImage_GetDotsPerMeterY(Bitmap))

End Function

Public Sub FreeImage_SetResolutionY(ByVal Bitmap As Long, ByVal Resolution As Double)

   ' This function sets a DIB's resolution in Y-direction measured
   ' in 'dots per inch' (DPI) and not in 'dots per meter'.

   Call FreeImage_SetDotsPerMeterY(Bitmap, Int(Resolution / 0.0254 + 0.5))

End Sub

' ICC Color Profile functions

Public Function FreeImage_GetICCProfile(ByVal Bitmap As Long) As FIICCPROFILE

   ' This function is a wrapper for the FreeImage_GetICCProfile() function, returning
   ' a real FIICCPROFILE structure.
   
   ' Since the original FreeImage function returns a pointer to the FIICCPROFILE
   ' structure (FIICCPROFILE *), as with string returning functions, this wrapper is
   ' needed as VB can't declare a function returning a pointer to anything. So,
   ' analogous to string returning functions, FreeImage_GetICCProfile() is declared
   ' private as FreeImage_GetICCProfileInt() and made publicly available with this
   ' wrapper function.

   Call CopyMemory(FreeImage_GetICCProfile, _
                   ByVal FreeImage_GetICCProfileInt(Bitmap), _
                   LenB(FreeImage_GetICCProfile))

End Function

Public Function FreeImage_GetICCProfileSize(ByVal Bitmap As Long) As Long

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' only the size in bytes of the ICC profile data for the Bitmap specified or zero,
   ' if there is no ICC profile data for the Bitmap.

   FreeImage_GetICCProfileSize = pDeref(FreeImage_GetICCProfileInt(Bitmap) + 4)

End Function

Public Function FreeImage_GetICCProfileDataPointer(ByVal Bitmap As Long) As Long

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' only the pointer (the address) of the ICC profile data for the Bitmap specified,
   ' or zero if there is no ICC profile data for the Bitmap.

   FreeImage_GetICCProfileDataPointer = pDeref(FreeImage_GetICCProfileInt(Bitmap) + 8)

End Function

Public Function FreeImage_HasICCProfile(ByVal Bitmap As Long) As Boolean

   ' This function is a thin wrapper around FreeImage_GetICCProfile() returning
   ' True, if there is an ICC color profile available for the Bitmap specified or
   ' returns False otherwise.

   FreeImage_HasICCProfile = (FreeImage_GetICCProfileSize(Bitmap) <> 0)

End Function

'ADDED BY TANNER:
Public Function FreeImage_GetPalette_ByTanner(ByVal fiHandle As Long, ByRef dstQuad() As RGBQuad, ByRef numOfColors As Long) As Boolean

    On Error GoTo GetPaletteFailure
    
    FreeImage_GetPalette_ByTanner = False
    
    'Validate handle
    If (fiHandle = 0) Then Exit Function
    
    'Validate color count
    numOfColors = FreeImage_GetColorsUsed(fiHandle)
    If (numOfColors = 0) Then Exit Function
    
    'Validate palette handle
    Dim palHandle As Long
    palHandle = FreeImage_GetPalette(fiHandle)
    If (palHandle = 0) Then Exit Function
    
    'If we're still here, we have what we need to populate the RGB quad array
    FreeImage_GetPalette_ByTanner = True
    ReDim dstQuad(0 To numOfColors - 1) As RGBQuad
    CopyMemoryStrict VarPtr(dstQuad(0)), palHandle, numOfColors * 4
    
    Exit Function
    
GetPaletteFailure:
    PDDebug.LogAction "WARNING!  FreeImage.FreeImage_GetPalette_ByTanner failed unexpectedly.", PDM_External_Lib
    FreeImage_GetPalette_ByTanner = False
    
End Function

' Image color depth conversion wrapper

Public Function FreeImage_GetPaletteEx(ByVal Bitmap As Long) As RGBQuad()

Dim tSA As SAVEARRAY1D
Dim lpSA As Long

   ' This function returns a VB style array of type RGBQUAD, containing
   ' the palette data of the Bitmap. This array provides read and write access
   ' to the actual palette data provided by FreeImage. This is done by
   ' creating a VB array with an own SAFEARRAY descriptor making the
   ' array point to the palette pointer returned by FreeImage_GetPalette().
   
   ' This makes you use code like you would in C/C++:
   
   ' // this code assumes there is a bitmap loaded and
   ' // present in a variable called "dib"
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
   
   
   If (Bitmap) Then
      
      ' create a proper SAVEARRAY descriptor
      With tSA
         .cbElements = 4                              ' size in bytes of RGBQUAD structure
         .cDims = 1                                   ' the array has only 1 dimension
         .cElements = FreeImage_GetColorsUsed(Bitmap) ' the number of elements in the array is
                                                      ' the number of used colors in the Bitmap
         .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE     ' need AUTO and FIXEDSIZE for safety issues,
                                                      ' so the array can not be modified in size
                                                      ' or erased; according to Matthew Curland never
                                                      ' use FIXEDSIZE alone
         .pvData = FreeImage_GetPalette(Bitmap)       ' let the array point to the memory block, the
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

Private Function FreeImage_IsGreyscaleImage(ByVal Bitmap As Long) As Boolean

Dim atRGB() As RGBQuad
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

   Select Case FreeImage_GetBPP(Bitmap)
   
   Case 1, 4, 8
      atRGB = FreeImage_GetPaletteEx(Bitmap)
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

Public Function FreeImage_ConvertColorDepth(ByVal Bitmap As Long, _
                                            ByVal Conversion As FREE_IMAGE_CONVERSION_FLAGS, _
                                   Optional ByVal UnloadSource As Boolean, _
                                   Optional ByVal threshold As Byte = 128, _
                                   Optional ByVal ditherMethod As FREE_IMAGE_DITHER = FID_FS, _
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
   lBPP = FreeImage_GetBPP(Bitmap)

   If (Bitmap) Then
   
      If (Not FreeImage_HasPixels(Bitmap)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to convert a 'header-only' bitmap.")
      End If
   
      Select Case (Conversion And (Not FICF_REORDER_GREYSCALE_PALETTE))
      
      Case FICF_MONOCHROME_THRESHOLD
         If (lBPP > 1) Then
            hDIBNew = FreeImage_Threshold(Bitmap, threshold)
         End If

      Case FICF_MONOCHROME_DITHER
         If (lBPP > 1) Then
            hDIBNew = FreeImage_Dither(Bitmap, ditherMethod)
         End If
      
      Case FICF_GREYSCALE_4BPP
         If (lBPP <> 4) Then
            ' If the color depth is 1 bpp and the we don't have a linear ramp palette
            ' the bitmap is first converted to an 8 bpp greyscale bitmap with a linear
            ' ramp palette and then to 4 bpp.
            If ((lBPP = 1) And (FreeImage_GetColorType(Bitmap) = FIC_PALETTE)) Then
               hDIBTemp = Bitmap
               Bitmap = FreeImage_ConvertToGreyscale(Bitmap)
               Call FreeImage_Unload(hDIBTemp)
            End If
            hDIBNew = FreeImage_ConvertTo4Bits(Bitmap)
         Else
            ' The bitmap is already 4 bpp but may not have a linear ramp.
            ' If we force a linear ramp the bitmap is converted to 8 bpp with a linear ramp
            ' and then back to 4 bpp.
            If (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(Bitmap))) Or _
                (bForceLinearRamp And (FreeImage_GetColorType(Bitmap) = FIC_PALETTE))) Then
               hDIBTemp = FreeImage_ConvertToGreyscale(Bitmap)
               hDIBNew = FreeImage_ConvertTo4Bits(hDIBTemp)
               Call FreeImage_Unload(hDIBTemp)
            End If
         End If
            
      Case FICF_GREYSCALE_8BPP
         ' Convert, if the bitmap is not at 8 bpp or does not have a linear ramp palette.
         If ((lBPP <> 8) Or (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(Bitmap))) Or _
                             (bForceLinearRamp And (FreeImage_GetColorType(Bitmap) = FIC_PALETTE)))) Then
            hDIBNew = FreeImage_ConvertToGreyscale(Bitmap)
         End If
         
      Case FICF_PALLETISED_8BPP
         ' note, that the FreeImage library only quantizes 24 bit images
         ' do not convert any 8 bit images
         If (lBPP <> 8) Then
            ' images with a color depth of 24 bits can directly be
            ' converted with the FreeImage_ColorQuantize function;
            ' other images need to be converted to 24 bits first
            If (lBPP = 24) Then
               hDIBNew = FreeImage_ColorQuantize(Bitmap, QuantizeMethod)
            Else
               hDIBTemp = FreeImage_ConvertTo24Bits(Bitmap)
               hDIBNew = FreeImage_ColorQuantize(hDIBTemp, QuantizeMethod)
               Call FreeImage_Unload(hDIBTemp)
            End If
         End If
         
      Case FICF_RGB_15BPP
         If (lBPP <> 15) Then
            hDIBNew = FreeImage_ConvertTo16Bits555(Bitmap)
         End If
      
      Case FICF_RGB_16BPP
         If (lBPP <> 16) Then
            hDIBNew = FreeImage_ConvertTo16Bits565(Bitmap)
         End If
      
      Case FICF_RGB_24BPP
         If (lBPP <> 24) Then
            hDIBNew = FreeImage_ConvertTo24Bits(Bitmap)
         End If
      
      Case FICF_RGB_32BPP
         If (lBPP <> 32) Then
            hDIBNew = FreeImage_ConvertTo32Bits(Bitmap)
         End If
      
      End Select
      
      If (hDIBNew) Then
         FreeImage_ConvertColorDepth = hDIBNew
         If (UnloadSource) Then
            Call FreeImage_Unload(Bitmap)
         End If
      Else
         FreeImage_ConvertColorDepth = Bitmap
      End If
   
   End If

End Function

Public Function FreeImage_ColorQuantizeEx(ByVal Bitmap As Long, _
                                 Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, _
                                 Optional ByVal UnloadSource As Boolean, _
                                 Optional ByVal paletteSize As Long = 256, _
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

   If (Bitmap) Then
   
      If (Not FreeImage_HasPixels(Bitmap)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to quantize a 'header-only' bitmap.")
      End If
      
      If (FreeImage_GetBPP(Bitmap) <> 24) Then
         hTmp = Bitmap
         Bitmap = FreeImage_ConvertTo24Bits(Bitmap)
         If (UnloadSource) Then
            Call FreeImage_Unload(hTmp)
         End If
         UnloadSource = True
      End If
      
      ' adjust PaletteSize
      If (paletteSize < 2) Then
         paletteSize = 2
      ElseIf (paletteSize > 256) Then
         paletteSize = 256
      End If
      
      lpPalette = pGetMemoryBlockPtrFromVariant(ReservePalette, lBlockSize, lElementSize)
      FreeImage_ColorQuantizeEx = FreeImage_ColorQuantizeExInt(Bitmap, QuantizeMethod, _
            paletteSize, ReserveSize, lpPalette)
      
      If (UnloadSource) Then
         Call FreeImage_Unload(Bitmap)
      End If
   End If

End Function

' Image Rescale wrapper functions

Public Function FreeImage_RescaleEx(ByVal Bitmap As Long, _
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
   
   If (Bitmap) Then
   
      If (Not FreeImage_HasPixels(Bitmap)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to rescale a 'header-only' bitmap.")
      End If
   
      If (Not IsMissing(Width)) Then
         Select Case VarType(Width)
         
         Case vbDouble, vbSingle, vbDecimal, vbCurrency
            lNewWidth = FreeImage_GetWidth(Bitmap) * Width
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
            lNewHeight = FreeImage_GetHeight(Bitmap) * Height
            If (IsPercentValue) Then
               lNewHeight = lNewHeight / 100
            End If
         
         Case Else
            lNewHeight = Height
         
         End Select
      End If
      
      If ((lNewWidth > 0) And (lNewHeight > 0)) Then
         If (ForceCloneCreation) Then
            hDIBNew = FreeImage_Rescale(Bitmap, lNewWidth, lNewHeight, Filter)
         
         ElseIf ((lNewWidth <> FreeImage_GetWidth(Bitmap)) Or _
                 (lNewHeight <> FreeImage_GetHeight(Bitmap))) Then
            hDIBNew = FreeImage_Rescale(Bitmap, lNewWidth, lNewHeight, Filter)
         
         End If
          
      ElseIf (lNewWidth > 0) Then
         If ((lNewWidth <> FreeImage_GetWidth(Bitmap)) Or _
             (ForceCloneCreation)) Then
            lNewHeight = lNewWidth / (FreeImage_GetWidth(Bitmap) / FreeImage_GetHeight(Bitmap))
            hDIBNew = FreeImage_Rescale(Bitmap, lNewWidth, lNewHeight, Filter)
         End If
      
      ElseIf (lNewHeight > 0) Then
         If ((lNewHeight <> FreeImage_GetHeight(Bitmap)) Or _
             (ForceCloneCreation)) Then
            lNewWidth = lNewHeight * (FreeImage_GetWidth(Bitmap) / FreeImage_GetHeight(Bitmap))
            hDIBNew = FreeImage_Rescale(Bitmap, lNewWidth, lNewHeight, Filter)
         End If
      
      End If
      
      If (hDIBNew) Then
         FreeImage_RescaleEx = hDIBNew
         If (UnloadSource) Then
            Call FreeImage_Unload(Bitmap)
         End If
      Else
         FreeImage_RescaleEx = Bitmap
      End If
   End If
                     
End Function

Public Function FreeImage_RescaleByPixel(ByVal Bitmap As Long, _
                                Optional ByVal WidthInPixels As Long, _
                                Optional ByVal HeightInPixels As Long, _
                                Optional ByVal UnloadSource As Boolean, _
                                Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, _
                                Optional ByVal ForceCloneCreation As Boolean) As Long
                                
   ' Thin wrapper for function 'FreeImage_RescaleEx' for removing method
   ' overload fake. This function rescales the image directly to the size
   ' specified by the 'WidthInPixels' and 'HeightInPixels' parameters.

   FreeImage_RescaleByPixel = FreeImage_RescaleEx(Bitmap, WidthInPixels, HeightInPixels, False, _
         UnloadSource, Filter, ForceCloneCreation)

End Function

Public Function FreeImage_CreateFromDC(ByVal hDC As Long, _
                              Optional ByRef hBitmap As Long) As Long

Dim tBM As Bitmap_API
Dim hDIB As Long
Dim lResult As Long
Dim nColors As Long
Dim lpInfo As Long

   ' Creates a FreeImage DIB from a Device Context/Compatible Bitmap. This
   ' function returns a pointer to the DIB as, for instance, 'FreeImage_Load()'
   ' does. So, this could be a real replacement for FreeImage_Load() or
   ' 'FreeImage_CreateFromOlePicture()' when working with DCs and Bitmaps directly
   
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
   ' The caller is responsible to destroy or free the DC and Bitmap if necessary.
   
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
         ' The GetDIBits function clears the biClrUsed and biClrImportant BitmapINFO
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
               ' restore BitmapINFO members
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

Public Function FreeImage_SaveEx(ByVal Bitmap As Long, _
                                 ByVal srcFilename As String, _
                        Optional ByVal imgFormat As FREE_IMAGE_FORMAT = FIF_UNKNOWN, _
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

   
   If (Bitmap) Then
   
      If (Not FreeImage_HasPixels(Bitmap)) Then
         Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                        "Unable to save 'header-only' bitmaps.")
      End If
   
      If ((Not IsMissing(Width)) Or _
          (Not IsMissing(Height))) Then
          
         lBPP = FreeImage_GetBPP(Bitmap)
         hDIBRescale = FreeImage_RescaleEx(Bitmap, Width, Height, InPercent, UnloadSource, Filter)
         bIsNewDIB = (hDIBRescale <> Bitmap)
         Bitmap = hDIBRescale
         bConvertedOnRescale = (lBPP <> FreeImage_GetBPP(Bitmap))
      End If
      
      If (imgFormat = FIF_UNKNOWN) Then
         imgFormat = FreeImage_GetFIFFromFilenameU(StrPtr(srcFilename))
      End If
      If (imgFormat <> FIF_UNKNOWN) Then
         If ((FreeImage_FIFSupportsWriting(imgFormat)) And _
             (FreeImage_FIFSupportsExportType(imgFormat, FIT_BITMAP))) Then
            
            'If (Not FreeImage_IsFilenameValidForFIF(imgFormat, srcFilename)) Then
            '   'Edit by Tanner: don't prevent me from writing whatever file extensions I damn well please!  ;)
            '   'strExtension = "." & FreeImage_GetPrimaryExtensionFromFIF(imgFormat)
            'End If
            
            ' check color depth
            If (colorDepth <> FICD_AUTO) Then
               ' mask out bit 1 (0x02) for the case ColorDepth is FICD_MONOCHROME_DITHER (0x03)
               ' FREE_IMAGE_COLOR_DEPTH values are true bit depths in general except FICD_MONOCHROME_DITHER
               ' by masking out bit 1, 'FreeImage_FIFSupportsExportBPP()' tests for bitdepth 1
               ' what is correct again for dithered images.
               colorDepth = (colorDepth And (Not &H2))
               If (Not FreeImage_FIFSupportsExportBPP(imgFormat, colorDepth)) Then
                  Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                                 "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(imgFormat) & "' " & _
                                 "is unable to write images with a color depth " & _
                                 "of " & colorDepth & " bpp.")
               
               ElseIf (FreeImage_GetBPP(Bitmap) <> colorDepth) Then
               
                  Bitmap = FreeImage_ConvertColorDepth(Bitmap, colorDepth, (UnloadSource Or bIsNewDIB))
                  bIsNewDIB = True
               
               End If
            Else
            
               If (lBPP = 0) Then
                  lBPP = FreeImage_GetBPP(Bitmap)
               End If
               
               If (Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) Then
                  lBPPOrg = lBPP
                  Do
                     lBPP = pGetPreviousColorDepth(lBPP)
                  Loop While ((Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) Or _
                              (lBPP = 0))
                  If (lBPP = 0) Then
                     lBPP = lBPPOrg
                     Do
                        lBPP = pGetNextColorDepth(lBPP)
                     Loop While ((Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) Or _
                                 (lBPP = 0))
                  End If
                  
                  If (lBPP <> 0) Then
                     Bitmap = FreeImage_ConvertColorDepth(Bitmap, lBPP, (UnloadSource Or bIsNewDIB))
                     bIsNewDIB = True
                  End If
               
               ElseIf (bConvertedOnRescale) Then
                  ' restore original color depth
                  ' always unload current DIB here, since 'bIsNewDIB' is True
                  Bitmap = FreeImage_ConvertColorDepth(Bitmap, lBPP, True)
                  
               End If
            End If
            
            FreeImage_SaveEx = FreeImage_Save(imgFormat, Bitmap, srcFilename & strExtension, Options)
            If ((bIsNewDIB) Or (UnloadSource)) Then
               Call FreeImage_Unload(Bitmap)
            End If
         Else
            Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf & _
                           "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(imgFormat) & "' " & _
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
Private Function pGetStringFromPointerA(ByVal ptr As Long) As String
    pGetStringFromPointerA = Strings.StringFromCharPtr(ptr, False)
End Function

Private Function pDeref(ByVal ptr As Long) As Long

   ' This function dereferences a pointer and returns the
   ' contents as it's return value.
   
   ' in C/C++ this would be:
   ' return *(ptr);
   
   Call CopyMemory(pDeref, ByVal ptr, 4)

End Function

Private Function pGetMemoryBlockPtrFromVariant(ByRef Data As Variant, _
                                      Optional ByRef sizeInBytes As Long, _
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
            If (sizeInBytes <= 0) Then
               sizeInBytes = (UBound(Data) + 1)
            
            ElseIf (sizeInBytes > (UBound(Data) + 1)) Then
               sizeInBytes = (UBound(Data) + 1)
            
            End If
         End If
      
      Case vbInteger
         ElementSize = 2
         pGetMemoryBlockPtrFromVariant = pGetArrayPtrFromVariantArray(Data)
         If (pGetMemoryBlockPtrFromVariant) Then
            If (sizeInBytes <= 0) Then
               sizeInBytes = (UBound(Data) + 1) * 2
            
            ElseIf (sizeInBytes > ((UBound(Data) + 1) * 2)) Then
               sizeInBytes = (UBound(Data) + 1) * 2
            
            End If
         End If
      
      Case vbLong
         ElementSize = 4
         pGetMemoryBlockPtrFromVariant = pGetArrayPtrFromVariantArray(Data)
         If (pGetMemoryBlockPtrFromVariant) Then
            If (sizeInBytes <= 0) Then
               sizeInBytes = (UBound(Data) + 1) * 4
            
            ElseIf (sizeInBytes > ((UBound(Data) + 1) * 4)) Then
               sizeInBytes = (UBound(Data) + 1) * 4
            
            End If
         End If
      
      End Select
   Else
      ElementSize = 1
      If ((VarType(Data) = vbLong) And _
          (sizeInBytes >= 0)) Then
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

'Added by Tanner: wrapper to flip a FI DIB by handle
Public Function FreeImage_FlipVertically(ByVal fi_DIB As Long) As Boolean
    FreeImage_FlipVertically = (FreeImage_FlipVertical(fi_DIB) <> 0)
End Function

'Metadata functions
Public Function FreeImage_CreateTagEx(ByVal Model As FREE_IMAGE_MDMODEL, _
                             Optional ByVal Key As String, _
                             Optional ByVal TagType As FREE_IMAGE_MDTYPE = FIDT_NOTYPE, _
                             Optional ByRef Value As Variant, _
                             Optional ByRef Count As Long, _
                             Optional ByVal Id As Long) As FREE_IMAGE_TAG
                             
   ' This function is a wrapper for FreeImage_CreateTag() working with
   ' the VB friendly FREE_IMAGE_TAG structure. So, the return value is
   ' not a pointer to a FITAG structure but a FREE_IMAGE_TAG structure.
   
   ' In contrast to FreeImage's original FreeImage_CreateTag() function, the
   ' parameter 'Model' must be specified, the parameters 'Key', 'TagType',
   ' 'Value', 'Count' and 'Id' my be specified.
   
   ' The 'Model' is needed, since each FREE_IMAGE_TAG structure needs a
   ' valid 'Model' member.
   
   ' All other parameters are optional and enable the caller to specify the tag's
   ' values upon tag creation. Any parameter specified, is set to it's corresponding
   ' member in the FREE_IMAGE_TAG structure.
   
   ' The caller should check the returned FREE_IMAGE_TAG structure's 'TagPtr' member.
   ' If this function succeeded, the 'TagPtr' member is non zero. A value of zero
   ' indicates an error condition sourced from FreeImage_CreateTag().
   
   With FreeImage_CreateTagEx
      .tagPtr = FreeImage_CreateTag()
      If (.tagPtr <> 0) Then
         .Model = Model
         If (LenB(Key) > 0) Then
            .Key = Key
         End If
         .Type = TagType
         .Count = Count
         .Id = Id
         If (Not IsMissing(Value)) Then
            .Value = Value
         End If
         Call pTagToTagPtr(FreeImage_CreateTagEx)
         FreeImage_CreateTagEx = pGetTagFromTagPtr(Model, .tagPtr)
      End If
   End With

End Function

'Smarter tag creator, which allows for specialized types like palettes
Public Function FreeImage_CreateTagTanner(ByVal fiDIB As Long, ByVal Model As FREE_IMAGE_MDMODEL, _
                             Optional ByVal Key As String, _
                             Optional ByVal TagType As FREE_IMAGE_MDTYPE = FIDT_NOTYPE, _
                             Optional ByVal dataPtr As Long, _
                             Optional ByVal dataCount As Long, _
                             Optional ByVal dataLength As Long, _
                             Optional ByVal dataID As Long) As Boolean
                                 
    Dim tagPtr As Long
    tagPtr = FreeImage_CreateTag()
    
    FreeImage_SetTagKey tagPtr, Key
    FreeImage_SetTagID tagPtr, dataID
    FreeImage_SetTagType tagPtr, TagType
    FreeImage_SetTagLength tagPtr, dataLength
    FreeImage_SetTagCount tagPtr, dataCount
    FreeImage_SetTagValue tagPtr, dataPtr
    
    FreeImage_CreateTagTanner = (FreeImage_SetMetadataInt(Model, fiDIB, Key, tagPtr) <> 0)
    FreeImage_DeleteTag tagPtr
    
End Function

Public Function FreeImage_SetMetadataEx(ByVal Bitmap As Long, _
                                        ByRef Tag As FREE_IMAGE_TAG, _
                               Optional ByVal Key As String, _
                               Optional ByVal Model As FREE_IMAGE_MDMODEL = FIMD_NODATA, _
                               Optional ByVal RefreshTag As Boolean) As Boolean
                               
   ' This function is a wrapper for FreeImage_SetMetadata() using the wrapper's
   ' VB friendly FREE_IMAGE_TAG structure as an replacement for the original
   ' function's pointer to a FITAG structure.
   
   ' All parameters 'Bitmap', 'Tag', 'Key' and 'Model' as the function's return value
   ' work according to the FreeImage API documentation expect that Tag is not a
   ' pointer to a FITAG structure but a FREE_IMAGE_TAG structure.
   
   ' As with FreeImage_SetMetadata(), this function sould only be called with
   ' new tags, created with FreeImage_CreateTagEx(), a wrapper function for
   ' FreeImage_CreateTag() working with the VB friendly FREE_IMAGE_TAG structure.
   
   ' Normally, after a newly created tag must be deleted/freed with a call to
   ' FreeImage_DeleteTagEx(), a wrapper function for FreeImage_DeleteTag() working
   ' with the VB friendly FREE_IMAGE_TAG structure (bored already?), after
   ' the tag was appended to an image's metadata model with
   ' FreeImage_SetMetadataEx(). But...
   
   ' There is a wrapper specific additional boolean parameter 'RefreshTag', that
   ' is similar to the parameter 'UnloadSource' found in many wrapper functions.
   ' When 'RefreshTag' is True upon entry, the tag specified in the 'Tag'
   ' parameter is deleted (the underlaying FITAG structure is deleted with
   ' FreeImage_DeteleTag() and all other members of the FREE_IMAGE_TAG structure
   ' are set to null values) and is reassigned with the tag, that is now part
   ' of the image's metadata model. The tag now referenced in the 'Tag'
   ' parameter must NOT be deleted any more by the caller of this function, since
   ' this tag refers to the actual tag data stored with the image. This is like
   ' a FREE_IMAGE_TAG structure obtained from FreeImage_GetMetadata() or
   ' FreeImage_FindFirst/NextMetadata(). Any changes made to this FREE_IMAGE_TAG
   ' structure may be applied to the image with a later call to
   ' FreeImage_UpdateMetadata().
   
   
   With Tag
      If (Model = FIMD_NODATA) Then
         Model = .Model
      End If
      If (LenB(Key) = 0) Then
         Key = .Key
      End If
      If (FreeImage_SetMetadataInt(Model, Bitmap, Key, .tagPtr) <> 0) Then
         FreeImage_SetMetadataEx = True
      End If
      If (RefreshTag) Then
         Call FreeImage_DeleteTagEx(Tag)
         Call FreeImage_GetMetadataEx(Model, Bitmap, Key, Tag)
      End If
   End With

End Function

Public Function FreeImage_GetMetadataEx(ByVal Model As FREE_IMAGE_MDMODEL, _
                                        ByVal Bitmap As Long, _
                                        ByVal Key As String, _
                                        ByRef Tag As FREE_IMAGE_TAG) As Boolean
   
   ' This function is a wrapper for FreeImage_GetMetadata() working with
   ' the VB friendly FREE_IMAGE_TAG structure. All parameters 'Bitmap', 'Tag',
   ' 'Key' and 'Model' as well as the function's return value work according
   ' to the FreeImage API documentation expect that Tag is not a pointer to
   ' a FITAG structure but to a FREE_IMAGE_TAG structure.
   
   ' Tags obtained from FreeImage_GetMetadataEx() must not be deleted with
   ' FreeImage_DeleteTagEx().

   With Tag
      If (FreeImage_GetMetadataInt(Model, Bitmap, Key, .tagPtr) <> 0) Then
         Tag = pGetTagFromTagPtr(Model, .tagPtr)
         FreeImage_GetMetadataEx = True
      End If
   End With

End Function

Public Sub FreeImage_DeleteTagEx(ByRef Tag As FREE_IMAGE_TAG)

   ' This function is a wrapper for FreeImage_DeleteTag() working with
   ' the VB friendly FREE_IMAGE_TAG structure. So, the parameter 'Tag'
   ' is not a pointer to a FITAG structure but a FREE_IMAGE_TAG structure.
   
   ' This function deletes the underlaying FreeImage FITAG structure,
   ' specified the the member 'TagPtr' of the FREE_IMAGE_TAG structure
   ' and also sets all other members of Tag to a null value.
   
   ' Do not get confused with the wrapper functions FreeImage_RemoveTag()
   ' and FreeImage_RemoveTagEx(). These functions remove a tag from an
   ' image's metadata model. This function only deletes of frees (a better
   ' name would be 'FreeImage_FreeTag') a tag created with
   ' FreeImage_CreateTagEx(). Do not delete any tags obtained from any other
   ' function.

   With Tag
      If (.tagPtr <> 0) Then
         Call FreeImage_DeleteTag(.tagPtr)
      End If
      .tagPtr = 0
      .Count = 0
      .Description = vbNullString
      .Id = 0
      .Key = vbNullString
      .Length = 0
      .Model = FIMD_NODATA
      Erase .Palette
      Erase .RationalValue
      .StringValue = vbNullString
      .Type = FIDT_NOTYPE
      .Value = Empty
   End With

End Sub

Private Function pTagToTagPtr(ByRef Tag As FREE_IMAGE_TAG) As Boolean

Dim tTagSave As FITAG
Dim lpTag As Long
Dim abValueBuffer() As Byte
Dim lLength As Long
Dim lCount As Long

   ' This function converts tag data stored in a VB friendly structure
   ' FREE_IMAGE_TAG into a real FreeImage tag pointer (FITAG **tag).
   
   ' This function is called, whenever tag data should be updated for an
   ' image, since the FreeImage's tag pointer remains valid during the
   ' whole lifetime of a DIB. So, changes written to that pointer (or
   ' even better, the FITAG structure at the address, the pointer points
   ' to), are real updates to the image's tag.

   With Tag
   
      lpTag = pDeref(.tagPtr)
      
      ' save current (FITAG) tag for an optional 'undo' operation
      ' invoked on failure
      Call CopyMemory(tTagSave, ByVal lpTag, Len(tTagSave))
      
      ' set tag id
      Call CopyMemory(ByVal lpTag + 8, .Id, 2)
      ' set tag type
      Call CopyMemory(ByVal lpTag + 10, .Type, 2)
      ' set tag key (use native FreeImage function to handle
      ' memory allocation)
      Call FreeImage_SetTagKey(.tagPtr, .Key)
      
      ' here, we update the tag's value
      ' generally, we create a plain byte buffer containing all the
      ' value's data and use FreeImage_SetTagValue() with the
      ' const void *value pointer set to the byte buffer's address.
      
      ' the variable abValueBuffer is our byte buffer that is,
      ' depending on the FreeImage tag data type, allocated and filled
      ' accordingly
      ' The variables 'lLength' and 'lCount' are set up correctly for
      ' each data type and will be filled into the FITAG structure
      ' before calling FreeImage_SetTagValue(); after all, the VB
      ' Tag structure's (FREE_IMAGE_TAG) 'Count' and 'Length' members
      ' are updated with 'lLength' and 'lCount'.
      
      Select Case .Type
      
      Case FIDT_ASCII
         ' use StrConv() to get an ASCII byte array from a VB String (BSTR)
         abValueBuffer = StrConv(.Value, vbFromUnicode)
         ' according to FreeImage's source code, both 'count' and 'length'
         ' must be the length of the string
         lCount = Len(.Value)
         lLength = lCount
         
      Case FIDT_PALETTE
         ' ensure, that there are at least 'Count' entries in the
         ' palette array
         lCount = .Count
         If (UBound(.Palette) + 1 < lCount) Then
            ' if not, adjust Count
            lCount = UBound(.Palette) + 1
         End If
         ' 4 bytes per element
         lLength = lCount * 4
         ' allocate buffer and copy data from Palatte array
         ReDim abValueBuffer(lLength - 1)
         Call CopyMemory(abValueBuffer(0), .Palette(LBound(.Palette)), lLength)
         
      Case FIDT_RATIONAL, _
           FIDT_SRATIONAL
         ' we use a helper function to get a byte buffer for any type of
         ' rational value
         lCount = pGetRationalValueBuffer(.RationalValue, abValueBuffer)
         If (lCount > .Count) Then
            lCount = .Count
         End If
         ' eight bytes per element (2 longs)
         lLength = lCount * 8
         
      Case Else
         ' we use a helper function to get a byte buffer for any other type
         lCount = pGetValueBuffer(.Value, .Type, lLength, abValueBuffer)
         If (lCount > .Count) Then
            lCount = .Count
         End If
         ' lLength was used as an OUT parameter when calling pGetValueBuffer
         ' it now contains the size of one element in bytes so, multiply with
         ' lCount to get the total length
         lLength = lLength * lCount
      
      End Select
      
      ' set tag length
      Call CopyMemory(ByVal lpTag + 16, lLength, 4)
      ' set tag count
      Call CopyMemory(ByVal lpTag + 12, lCount, 4)
       
      If (FreeImage_SetTagValue(.tagPtr, VarPtr(abValueBuffer(0))) <> 0) Then
         
         ' update Tag's members
         ' update Count
         .Count = lCount
         ' update Length
         .Length = lLength
         ' update StringValue
         .StringValue = FreeImage_TagToString(.Model, .tagPtr)
         pTagToTagPtr = True
      Else
      
         ' restore saved (FITAG) tag values on failure
         Call CopyMemory(ByVal lpTag, tTagSave, Len(tTagSave))
      End If
   
   End With

End Function

Public Function FreeImage_TagToString(ByVal Model As Long, _
                                      ByVal Tag As Long, _
                             Optional ByVal Make As String) As String

   ' This function returns the result of the 'FreeImage_TagToString' function
   ' as VB String.
   
   ' All parameters work according to the FreeImage 3 API documentation.

   FreeImage_TagToString = pGetStringFromPointerA(FreeImage_TagToStringInt(Model, Tag, Make))

End Function

Private Function pGetValueBuffer(ByRef Value As Variant, _
                                 ByVal MetaDataVarType As FREE_IMAGE_MDTYPE, _
                                 ByRef ElementSize As Long, _
                                 ByRef Buffer() As Byte) As Long
                            
Dim lElementCount As Long
Dim bIsArray As Boolean
Dim abValueBuffer(7) As Byte
Dim cBytes As Long
Dim i As Long

   ' This function copies any value provided in the Variant 'Value'
   ' parameter into the byte array Buffer. 'Value' may contain an
   ' array as well. The values in the byte buffer are aligned to fit
   ' the FreeImage data type for tag values specified in
   ' 'MetaDataVarType'. For integer values, it does not matter, in
   ' which VB data type the values are provided. For example, it is
   ' possible to transform a provided byte array into a unsigned long
   ' array.
   
   ' The parameter 'ElementSize' is an OUT value, providing the actual
   ' size per element in the byte buffer in bytes to the caller.
   
   ' This function works for the types FIDT_BYTE, FIDT_SHORT, FIDT_LONG,
   ' FIDT_SBYTE , FIDT_SSHORT, FIDT_SLONG, FIDT_FLOAT, FIDT_DOUBLE
   ' and FIDT_IFD
                            
   ElementSize = pGetElementSize(MetaDataVarType)
   If (Not IsArray(Value)) Then
      lElementCount = 1
   Else
      On Error Resume Next
      lElementCount = UBound(Value) - LBound(Value) + 1
      On Error GoTo 0
      bIsArray = True
   End If
   
   If (lElementCount > 0) Then
      ReDim Buffer((lElementCount * ElementSize) - 1)
      
      If (Not bIsArray) Then
         cBytes = pGetVariantAsByteBuffer(Value, abValueBuffer)
         If (cBytes > ElementSize) Then
            cBytes = ElementSize
         End If
         Call CopyMemory(Buffer(0), abValueBuffer(0), cBytes)
      Else
         For i = LBound(Value) To UBound(Value)
            cBytes = pGetVariantAsByteBuffer(Value(i), abValueBuffer)
            If (cBytes > ElementSize) Then
               cBytes = ElementSize
            End If
            Call CopyMemory(Buffer(0 + (i * ElementSize)), abValueBuffer(0), cBytes)
         Next i
      End If
      
      pGetValueBuffer = lElementCount
   End If

End Function



'--------------------------------------------------------------------------------
' Private metadata helper functions
'--------------------------------------------------------------------------------

Private Function pGetTagFromTagPtr(ByVal Model As FREE_IMAGE_MDMODEL, _
                                   ByVal tagPtr As Long) As FREE_IMAGE_TAG

Dim tTag As FITAG
Dim lTemp As Long
Dim i As Long

   ' This function converts data stored in a real FreeImage tag
   ' pointer (FITAG **tag) into a VB friendly structure FREE_IMAGE_TAG.
   
   If (tagPtr <> 0) Then
   
      ' this is like (only like!) tTag tag = (FITAG) TagPtr; in C/C++
      ' we copy Len(tTag) bytes from the address in TagPtr in to a
      ' private FITAG structure tTag so we have easy access to all
      ' FITAG members
      Call CopyMemory(tTag, ByVal pDeref(tagPtr), Len(tTag))
      
      With pGetTagFromTagPtr
      
         ' first fill all members expect 'Value' in our
         ' VB friendly FREE_IMAGE_TAG structure
         
         ' since we use this VB friendly FREE_IMAGE_TAG structure
         ' for later tag modification too, we also need to store the
         ' tag model and the pointer to the actual FreeImage FITAG
         ' structure
         .Model = Model
         .tagPtr = tagPtr
         
         ' although FITAG's 'count' and 'length' members are
         ' unsigned longs, we do not expect values greater
         ' than 2,147,483,647, so we store them in normal VB
         ' signed longs
         .Count = tTag.Count
         .Length = tTag.Length
         
         ' strings are stored as pointers to the actual string
         ' data in FITAG
         .Description = pGetStringFromPointerA(tTag.Description)
         .Key = pGetStringFromPointerA(tTag.Key)
         
         ' FITAG's 'id' and 'type' members are unsigned shorts;
         ' first of all 'id' may exceed the range of a signed
         ' short (Integer data type in VB), so we store them in
         ' signed longs and use CopyMemory for to keep the
         ' unsigned bit layout
         Call CopyMemory(.Id, tTag.Id, 2)
         Call CopyMemory(.Type, tTag.Type, 2)
         
         ' StringValue is the result of FreeImage_TagToString(); we
         ' also store this tag representation in our structure
         .StringValue = FreeImage_TagToString(Model, tagPtr)
         
         ' now comes the hard part, getting the tag's value
         
         Select Case .Type
         
         Case FIDT_BYTE, _
              FIDT_UNDEFINED
            If (.Count > 1) Then
               Dim abBytes() As Byte
               ' for a byte array, just redim a VB Byte array and
               ' copy Count bytes from the pointer
               ReDim abBytes(.Count - 1)
               Call CopyMemory(abBytes(0), ByVal tTag.Value, .Count)
               .Value = abBytes
            Else
               ' copy a single byte into a Long and assign
               ' with CByte()
               Call CopyMemory(lTemp, ByVal tTag.Value, 1)
               .Value = CByte(lTemp)
            End If
         
         Case FIDT_ASCII
            ' for an ASCII string, 'value' is just a pointer to the
            ' string's actual data
            .Value = pGetStringFromPointerA(tTag.Value)
            
         Case FIDT_SHORT
            Dim iTemp As Integer
            If (.Count > 1) Then
               ' for a unsigned long array, first redim Value to
               ' proper size
               ReDim .Value(.Count - 1)
               ' iterate over all items
               For i = 0 To .Count - 1
                  ' copy each value into a Long and
                  ' assign with FreeImage_UnsignedShort() to the
                  ' corresponding item in the (Variant) Value array
                  Call CopyMemory(iTemp, ByVal tTag.Value + i * 2, 2)
                  .Value(i) = FreeImage_UnsignedShort(iTemp)
               Next i
            Else
               ' copy a single byte into a Long and assign
               ' with FreeImgage_UnsignedShort()
               Call CopyMemory(iTemp, ByVal tTag.Value, 2)
               ' this works although FreeImage_UnsignedShort() takes
               ' an Integer parameter since lTemp was 0 before and
               ' we copied only 2 bytes so, VB's implicit conversion
               ' to Integer will never produce an overflow
               .Value = FreeImage_UnsignedShort(iTemp)
            End If
            
         Case FIDT_LONG, _
              FIDT_IFD
            If (.Count > 1) Then
               ' for a unsigned long array, first redim Value to
               ' proper size
               ReDim .Value(.Count - 1)
               ' iterate over all items
               For i = 0 To .Count - 1
                  ' copy each value into a (signed) Long and
                  ' assign with FreeImage_UnsignedLong() to the
                  ' corresponding item in the (Variant) Value array
                  Call CopyMemory(lTemp, ByVal tTag.Value + i * 4, 4)
                  .Value(i) = FreeImage_UnsignedLong(lTemp)
               Next i
            Else
               ' copy a single unsigned long into a (signed) Long and
               ' assign with FreeImage_UnsignedLong()
               Call CopyMemory(lTemp, ByVal tTag.Value, 2)
               .Value = FreeImage_UnsignedLong(lTemp)
            End If
            
         Case FIDT_RATIONAL, _
              FIDT_SRATIONAL
            ' rational values are always stored in the FREE_IMAGE_TAG
            ' structure's FIRATIONAL array 'RationalValue' so, allocate
            ' enough space in both the 'Value' and 'RationalValue'
            ' members to hold 'Count' items
            ReDim .Value(.Count - 1)
            ReDim .RationalValue(.Count - 1)
            For i = 0 To .Count - 1
               ' iterate over all items
               With .RationalValue(i)
                  ' for each item, copy both numerator and denominator
                  ' into a (signed) Long and assign it to the corresponding
                  ' member of the FIRATIONAL structure so, we first assume
                  ' havinge a signed rational (FIDT_SRATIONAL) here
                  Call CopyMemory(lTemp, ByVal tTag.Value + i * 8, 4)
                  .numerator = lTemp
                  Call CopyMemory(lTemp, ByVal tTag.Value + i * 8 + 4, 4)
                  .denominator = lTemp
               End With
               ' if we have an unsigned rational (FIDT_RATIONAL), convert
               ' numerator and denominator
               If (.Type = FIDT_RATIONAL) Then
                  ' convert with FreeImage_UnsignedLong()
                  With .RationalValue(i)
                     .numerator = FreeImage_UnsignedLong(.numerator)
                     .denominator = FreeImage_UnsignedLong(.denominator)
                  End With
                  ' normalze the unsigned rational value
                  Call pNormalizeRational(.RationalValue(i))
               Else
                  ' normalze the signed rational value
                  Call pNormalizeSRational(.RationalValue(i))
               End If
               ' store the current fraction's value (maybe only approximated) in
               ' the 'Value' member of the FREE_IMAGE_TAG structure, if the
               ' denominator is not zero
               If (.RationalValue(i).denominator <> 0) Then
                  .Value(i) = .RationalValue(i).numerator / .RationalValue(i).denominator
               End If
            Next i
            
         Case FIDT_SBYTE
            If (.Count > 1) Then
               ' for a signed byte array, first redim Value to
               ' proper size
               ReDim .Value(.Count - 1)
               ' iterate over all items
               For i = 0 To .Count - 1
                  ' copy each signed byte value into a Long and
                  ' check, whether it is negative (bit 7 set)
                  Call CopyMemory(lTemp, ByVal tTag.Value, 1)
                  If (lTemp And 128) Then
                     ' if negative, calculate the negative value
                     ' and store it in an Integer
                     .Value(i) = CInt(-256 - (Not (lTemp - 1)))
                  Else
                     ' if positive, assign to Value as byte
                     .Value(i) = CByte(lTemp)
                  End If
               Next i
            Else
               ' copy a single signed byte into a Long and
               ' check, whether it is negative (bit 7 set)
               Call CopyMemory(lTemp, ByVal tTag.Value, 1)
               If (lTemp And 128) Then
                  ' if negative, calculate the negative value
                  ' and store it in an Integer
                  .Value = CInt(-256 - (Not (lTemp - 1)))
               Else
                  ' if positive, assign to Value as byte
                  .Value = CByte(lTemp)
               End If
            End If
            
         Case FIDT_SSHORT
            If (.Count > 1) Then
               Dim aiSShorts() As Integer
               ' for a signed short array, just redim a VB Integer array and
               ' copy Count bytes from the pointer
               ReDim aiSShorts(.Count - 1)
               Call CopyMemory(aiSShorts(0), ByVal tTag.Value, .Count * 2)
               .Value = aiSShorts
            Else
               ' copy a single signed short into a Long and assign
               ' with CInt()
               Call CopyMemory(lTemp, ByVal tTag.Value, 2)
               .Value = CInt(lTemp)
            End If
            
         Case FIDT_SLONG
            If (.Count > 1) Then
               Dim alSLongs() As Long
               ' for a signed long array, just redim a VB Long array and
               ' copy Count bytes from the pointer
               ReDim alSLongs(.Count - 1)
               Call CopyMemory(alSLongs(0), ByVal tTag.Value, .Count * 4)
               .Value = alSLongs
            Else
               ' copy a single signed long into a Long and assign
               ' directly
               Call CopyMemory(lTemp, ByVal tTag.Value, 4)
               .Value = lTemp
            End If
            
         Case FIDT_FLOAT
            If (.Count > 1) Then
               Dim asngFloats() As Single
               ' for a float array, just redim a VB Single array and
               ' copy Count bytes from the pointer
               ReDim asngFloats(.Count - 1)
               Call CopyMemory(asngFloats(0), ByVal tTag.Value, .Count * 4)
               .Value = asngFloats
            Else
               Dim sngFloat As Single
               ' copy a single float into a Single and assign
               ' directly
               Call CopyMemory(sngFloat, ByVal tTag.Value, 4)
               .Value = sngFloat
            End If
            
         Case FIDT_DOUBLE
            If (.Count > 1) Then
               Dim adblDoubles() As Double
               ' for a double array, just redim a VB Double array and
               ' copy Count bytes from the pointer
               ReDim adblDoubles(.Count - 1)
               Call CopyMemory(adblDoubles(0), ByVal tTag.Value, .Count * 8)
               .Value = adblDoubles
            Else
               Dim dblDouble As Double
               ' copy a single double into a Double and assign
               ' directly
               Call CopyMemory(dblDouble, ByVal tTag.Value, 8)
               .Value = dblDouble
            End If
            
         Case FIDT_PALETTE
            ' copy 'Count' palette entries (RGBQUAD) form the value
            ' pointer into the proper dimensioned array of RGBQUAD
            ReDim .Palette(.Count - 1)
            For i = 0 To .Count - 1
               Call CopyMemory(.Palette(i), ByVal tTag.Value + i * 4, 4)
            Next i
         
         End Select
      
      End With
   End If

End Function


Private Function pGetRationalValueBuffer(ByRef RationalValues() As FIRATIONAL, _
                                         ByRef Buffer() As Byte) As Long
                                         
Dim lElementCount As Long
Dim abValueBuffer(7) As Byte
Dim cBytes As Long
Dim i As Long

   ' This function copies a number of elements from the FIRATIONAL array
   ' 'RationalValues' into the byte buffer 'Buffer'.
   
   ' From the caller's point of view, this function is the same as
   ' 'pGetValueBuffer', except, it only works for arrays of FIRATIONAL.
   
   ' This function works for the types FIDT_RATIONAL and FIDT_SRATIONAL.
   
   lElementCount = UBound(RationalValues) - LBound(RationalValues) + 1
   ReDim Buffer(lElementCount * 8 + 1)
   
   For i = LBound(RationalValues) To UBound(RationalValues)
      cBytes = pGetVariantAsByteBuffer(RationalValues(i).numerator, abValueBuffer)
      If (cBytes > 4) Then
         cBytes = 4
      End If
      Call CopyMemory(Buffer(0 + (i * 8)), abValueBuffer(0), cBytes)
      
      cBytes = pGetVariantAsByteBuffer(RationalValues(i).denominator, abValueBuffer)
      If (cBytes > 4) Then
         cBytes = 4
      End If
      Call CopyMemory(Buffer(4 + (i * 8)), abValueBuffer(0), cBytes)
   Next i
   
   pGetRationalValueBuffer = lElementCount
                                         
End Function

Private Function pGetVariantAsByteBuffer(ByRef Value As Variant, _
                                         ByRef Buffer() As Byte) As Long

Dim lLength As Long

   ' This function fills a byte buffer 'Buffer' with data taken
   ' from a Variant parameter. Depending on the Variant's type and,
   ' width, it copies N (lLength) bytes into the buffer starting
   ' at the buffer's first byte at Buffer(0). The function returns
   ' the number of bytes copied.
   
   ' It is much easier to assign the Variant to a variable of
   ' the proper native type first, since gathering a Variant's
   ' actual value is a hard job to implement for each subtype.
   
   Select Case VarType(Value)
   
   Case vbByte
      Buffer(0) = Value
      lLength = 1
   
   Case vbInteger
      Dim iInteger As Integer
      iInteger = Value
      lLength = 2
      Call CopyMemory(Buffer(0), iInteger, lLength)
   
   Case vbLong
      Dim lLong As Long
      lLong = Value
      lLength = 4
      Call CopyMemory(Buffer(0), lLong, lLength)
   
   Case vbCurrency
      Dim cCurrency As Currency
      ' since the Currency data type is a so called scaled
      ' integer, we have to divide by 10.000 first to get the
      ' proper bit layout.
      cCurrency = Value / 10000
      lLength = 8
      Call CopyMemory(Buffer(0), cCurrency, lLength)
   
   Case vbSingle
      Dim sSingle As Single
      sSingle = Value
      lLength = 4
      Call CopyMemory(Buffer(0), sSingle, lLength)
   
   Case vbDouble
      Dim dblDouble As Double
      dblDouble = Value
      lLength = 8
      Call CopyMemory(Buffer(0), dblDouble, lLength)
   
   End Select
   
   pGetVariantAsByteBuffer = lLength
                                         
End Function

Private Function pGetElementSize(ByVal vt As FREE_IMAGE_MDTYPE) As Long

   ' This function returns the width in bytes for any of the
   ' FreeImage metadata tag data types.

   Select Case vt
   
   Case FIDT_BYTE, _
        FIDT_SBYTE, _
        FIDT_UNDEFINED, _
        FIDT_ASCII
      pGetElementSize = 1
   
   Case FIDT_SHORT, _
        FIDT_SSHORT
      pGetElementSize = 2

   Case FIDT_LONG, _
        FIDT_SLONG, _
        FIDT_FLOAT, _
        FIDT_PALETTE, _
        FIDT_IFD
      pGetElementSize = 4

   Case Else
      pGetElementSize = 8
      
   End Select

End Function

Private Sub pNormalizeRational(ByRef Value As FIRATIONAL)

Dim vntCommon As Long

   ' This function normalizes an unsigned fraction stored in a FIRATIONAL
   ' structure by cancelling down the fraction. This is commonly done
   ' by dividing both numerator and denominator by their greates
   ' common divisor (gcd).
   ' Does nothing if any of numerator and denominator is 1 or 0.

   With Value
      If ((.numerator <> 1) And (.denominator <> 1) And _
          (.numerator <> 0) And (.denominator <> 0)) Then
         vntCommon = gcd(.numerator, .denominator)
         If (vntCommon <> 1) Then
            ' convert values back to an unsigned long (may
            ' result in a subtype Currency if the range of the
            ' VB Long is insufficient for storing the value!)
            .numerator = FreeImage_UnsignedLong(.numerator / vntCommon)
            .denominator = FreeImage_UnsignedLong(.denominator / vntCommon)
         End If
      End If
   End With

End Sub

Private Sub pNormalizeSRational(ByRef Value As FIRATIONAL)

Dim lCommon As Long

   ' This function normalizes a signed fraction stored in a FIRATIONAL
   ' structure by cancelling down the fraction. This is commonly done
   ' by dividing both numerator and denominator by their greates
   ' common divisor (gcd).
   ' Does nothing if any of numerator and denominator is 1 or 0.
   
   With Value
      If ((.numerator <> 1) And (.denominator <> 1) And _
          (.numerator <> 0) And (.denominator <> 0)) Then
         lCommon = gcd(.numerator, .denominator)
         If (lCommon <> 1) Then
            ' using the CLng() function for not to get
            ' a subtype Double here
            .numerator = CLng(.numerator / lCommon)
            .denominator = CLng(.denominator / lCommon)
         End If
      End If
      
      ' adjust the position of the negative sign if one is present:
      ' it should preceed the numerator, not the denominator
      If (.denominator < 0) Then
         .denominator = -.denominator
         .numerator = -.numerator
      End If
   End With

End Sub

Private Function gcd(ByVal a As Variant, ByVal b As Variant) As Variant

Dim vntTemp As Variant

   ' calculate greatest common divisor

   Do While (b)
      vntTemp = b
      ' calculate b = a % b (modulo)
      ' this could be just:
      ' b = a Mod b
      ' but VB's Mod operator fails for unsigned
      ' long values stored in currency variables
      ' so, we use the mathematical definition of
      ' the modulo operator taken from Wikipedia.
      b = a - floor(a / b) * b
      a = vntTemp
   Loop
   gcd = a

End Function

Private Function floor(ByRef a As Variant) As Variant

   ' This is a VB version of the floor() function.
   If (a < 0) Then
      floor = VBA.Int(a)
   Else
      floor = -VBA.Fix(-a)
   End If

End Function

