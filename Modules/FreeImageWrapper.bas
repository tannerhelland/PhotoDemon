Attribute VB_Name = "Outside_FreeImageV3"
'Note: this file has been heavily modified for use within PhotoDemon.
'
'Much of this module was adapted from the official FreeImage VB6 wrapper by Carsten Klein.
' However, I have made endless changes to improve performance and reliability for PhotoDemon's specific use-case(s).
' I have also removed many wrapper functions that might be useful to you.
'
'Said another way: please avoid using this module in your own projects.  Instead, start from the original,
' official FreeImage version available from this link (good as of August 2020):
' http://freeimage.sourceforge.net/download.html
'
'Thank you to Carsten Klein and the FreeImage team for their commitment to supporting VB6 interfaces.
'
'Original copyright information follows
'
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

Private Type Bitmap_API
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long

'--------------------------------------------------------------------------------
' FreeImage 3 types, constants and enumerations
'--------------------------------------------------------------------------------

' Load / Save flag constants
Public Const FIF_LOAD_NOPIXELS = &H8000&              ' load the image header only (not supported by all plugins)

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
Public Const JPEG_SUBSAMPLING_422 As Long = &H8000&   ' save with low 2x1 chroma subsampling (4:2:2)
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
Public Const TIFF_JPEG As Long = &H8000&             ' save using JPEG compression
Public Const TIFF_LOGLUV As Long = &H10000           ' save using LogLuv compression

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
   Const FILO_LOAD_NOPIXELS = &H8000&
   Const FILO_LOAD_DEFAULT = 0
   Const FILO_GIF_DEFAULT = GIF_DEFAULT
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
   FISO_JXR_LOSSLESS = JXR_LOSSLESS               ' save lossless
   FISP_JXR_PROGRESSIVE = JXR_PROGRESSIVE         ' save as a progressive-JXR (use | to combine with other save flags)
End Enum
#If False Then
    Private Const FISO_SAVE_DEFAULT = 0, FISO_BMP_DEFAULT = BMP_DEFAULT, FISO_BMP_SAVE_RLE = BMP_SAVE_RLE, FISO_JPEG_DEFAULT = JPEG_DEFAULT, FISO_JPEG_QUALITYSUPERB = JPEG_QUALITYSUPERB, FISO_JPEG_QUALITYGOOD = JPEG_QUALITYGOOD, FISO_JPEG_QUALITYNORMAL = JPEG_QUALITYNORMAL, FISO_JPEG_QUALITYAVERAGE = JPEG_QUALITYAVERAGE, FISO_JPEG_QUALITYBAD = JPEG_QUALITYBAD, FISO_JPEG_PROGRESSIVE = JPEG_PROGRESSIVE, FISO_JPEG_SUBSAMPLING_411 = JPEG_SUBSAMPLING_411, FISO_JPEG_SUBSAMPLING_420 = JPEG_SUBSAMPLING_420, FISO_JPEG_SUBSAMPLING_422 = JPEG_SUBSAMPLING_422, FISO_JPEG_SUBSAMPLING_444 = JPEG_SUBSAMPLING_444, FISO_PNM_DEFAULT = PNM_DEFAULT, FISO_PNM_SAVE_RAW = PNM_SAVE_RAW, FISO_PNM_SAVE_ASCII = PNM_SAVE_ASCII, FISO_TARGA_SAVE_RLE = TARGA_SAVE_RLE
    Private Const FISO_TIFF_DEFAULT = TIFF_DEFAULT, FISO_TIFF_CMYK = TIFF_CMYK, FISO_TIFF_PACKBITS = TIFF_PACKBITS, FISO_TIFF_DEFLATE = TIFF_DEFLATE, FISO_TIFF_ADOBE_DEFLATE = TIFF_ADOBE_DEFLATE, FISO_TIFF_NONE = TIFF_NONE, FISO_TIFF_CCITTFAX3 = TIFF_CCITTFAX3, FISO_TIFF_CCITTFAX4 = TIFF_CCITTFAX4, FISO_TIFF_LZW = TIFF_LZW, FISO_TIFF_JPEG = TIFF_JPEG, FISO_JXR_LOSSLESS = JXR_LOSSLESS, FISP_JXR_PROGRESSIVE = JXR_PROGRESSIVE
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
    Private Const FIT_UNKNOWN = 0, FIT_BITMAP = 1, FIT_UINT16 = 2, FIT_INT16 = 3, FIT_UINT32 = 4, FIT_INT32 = 5, FIT_FLOAT = 6, FIT_DOUBLE = 7, FIT_COMPLEX = 8, FIT_RGB16 = 9, FIT_RGBA16 = 10, FIT_RGBF = 11, FIT_RGBAF = 12
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
    Private Const FIC_MINISWHITE = 0, FIC_MINISBLACK = 1, FIC_RGB = 2, FIC_PALETTE = 3, FIC_RGBALPHA = 4, FIC_CMYK = 5
#End If

Public Enum FREE_IMAGE_QUANTIZE
   FIQ_WUQUANT = 0           ' Xiaolin Wu color quantization algorithm
   FIQ_NNQUANT = 1           ' NeuQuant neural-net quantization algorithm by Anthony Dekker
   FIQ_LFPQUANT = 2          ' Lossless Fast Pseudo-Quantization Algorithm by Carsten Klein
End Enum
#If False Then
    Private Const FIQ_WUQUANT = 0, FIQ_NNQUANT = 1, FIQ_LFPQUANT = 2
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
    Private Const FID_FS = 0, FID_BAYER4x4 = 1, FID_BAYER8x8 = 2, FID_CLUSTER6x6 = 3, FID_CLUSTER8x8 = 4, FID_CLUSTER16x16 = 5, FID_BAYER16x16 = 6
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
    Private Const FILTER_BOX = 0, FILTER_BICUBIC = 1, FILTER_BILINEAR = 2, FILTER_BSPLINE = 3, FILTER_CATMULLROM = 4, FILTER_LANCZOS3 = 5
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
    Private Const FICF_MONOCHROME = &H1, FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME, FICF_MONOCHROME_DITHER = &H3, FICF_GREYSCALE_4BPP = &H4, FICF_PALLETISED_8BPP = &H8, FICF_GREYSCALE_8BPP = FICF_PALLETISED_8BPP Or FICF_MONOCHROME, FICF_GREYSCALE = FICF_GREYSCALE_8BPP, FICF_RGB_15BPP = &HF, FICF_RGB_16BPP = &H10, FICF_RGB_24BPP = &H18, FICF_RGB_32BPP = &H20, FICF_RGB_ALPHA = FICF_RGB_32BPP, FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0, FICF_REORDER_GREYSCALE_PALETTE = &H1000
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
    Private Const FICD_AUTO = &H0, FICD_MONOCHROME = &H1, FICD_MONOCHROME_THRESHOLD = FICF_MONOCHROME, FICD_MONOCHROME_DITHER = &H3, FICD_1BPP = FICD_MONOCHROME, FICD_4BPP = &H4, FICD_8BPP = &H8, FICD_15BPP = &HF, FICD_16BPP = &H10, FICD_24BPP = &H18, FICD_32BPP = &H20
#End If

'Note that padding will occur after the Flags member - which is a WORD in FreeImage.h
Private Type FIICCPROFILE
    Flags As Integer
    Size As Long
    Data As Long
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

' Bitmap management functions
Public Declare Function FreeImage_Allocate Lib "FreeImage.dll" Alias "_FreeImage_Allocate@24" ( _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal BitsPerPixel As Long, _
  Optional ByVal RedMask As Long, _
  Optional ByVal GreenMask As Long, _
  Optional ByVal BlueMask As Long) As Long
  
Public Declare Function FreeImage_HasPixelsInt Lib "FreeImage.dll" Alias "_FreeImage_HasPixels@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_LoadUInt Lib "FreeImage.dll" Alias "_FreeImage_LoadU@12" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal srcFilename As Long, _
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveUInt Lib "FreeImage.dll" Alias "_FreeImage_SaveU@16" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal fiBitmap As Long, _
           ByVal srcFilename As Long, _
  Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
  
Public Declare Sub FreeImage_Unload Lib "FreeImage.dll" Alias "_FreeImage_Unload@4" ( _
           ByVal fiBitmap As Long)

' Bitmap information functions
Public Declare Function FreeImage_GetImageType Lib "FreeImage.dll" Alias "_FreeImage_GetImageType@4" ( _
           ByVal fiBitmap As Long) As FREE_IMAGE_TYPE

Private Declare Function FreeImage_GetColorsUsed Lib "FreeImage.dll" Alias "_FreeImage_GetColorsUsed@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetBPP Lib "FreeImage.dll" Alias "_FreeImage_GetBPP@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetWidth Lib "FreeImage.dll" Alias "_FreeImage_GetWidth@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetHeight Lib "FreeImage.dll" Alias "_FreeImage_GetHeight@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetPitch Lib "FreeImage.dll" Alias "_FreeImage_GetPitch@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetPalette Lib "FreeImage" Alias "_FreeImage_GetPalette@4" (ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterX@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterY@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Sub FreeImage_SetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterX@8" ( _
           ByVal fiBitmap As Long, _
           ByVal Resolution As Long)

Public Declare Sub FreeImage_SetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterY@8" ( _
           ByVal fiBitmap As Long, _
           ByVal Resolution As Long)

Public Declare Function FreeImage_GetInfoHeader Lib "FreeImage.dll" Alias "_FreeImage_GetInfoHeader@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetInfo Lib "FreeImage.dll" Alias "_FreeImage_GetInfo@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetColorType Lib "FreeImage.dll" Alias "_FreeImage_GetColorType@4" ( _
           ByVal fiBitmap As Long) As FREE_IMAGE_COLOR_TYPE

Public Declare Function FreeImage_GetTransparencyCount Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyCount@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_Invert Lib "FreeImage.dll" Alias "_FreeImage_Invert@4" (ByVal fiBitmap As Long) As Long

Private Declare Function FreeImage_IsTransparentInt Lib "FreeImage.dll" Alias "_FreeImage_IsTransparent@4" ( _
           ByVal fiBitmap As Long) As Long
           
Public Declare Function FreeImage_GetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_GetTransparentIndex@4" ( _
           ByVal fiBitmap As Long) As Long
           
Public Declare Function FreeImage_SetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_SetTransparentIndex@8" ( _
           ByVal fiBitmap As Long, _
           ByVal Index As Long) As Long
           
'Public Declare Function FreeImage_GetThumbnail Lib "FreeImage.dll" Alias "_FreeImage_GetThumbnail@4" ( _
           ByVal fiBitmap as Long) As Long
           
' Filetype functions
Public Declare Function FreeImage_GetFileTypeU Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeU@8" ( _
           ByVal srcFilename As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_GetFileTypeFromMemory Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromMemory@8" ( _
           ByVal Stream As Long, _
  Optional ByVal Size As Long) As FREE_IMAGE_FORMAT


' Pixel access functions
Public Declare Function FreeImage_GetBits Lib "FreeImage.dll" Alias "_FreeImage_GetBits@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetScanline Lib "FreeImage.dll" Alias "_FreeImage_GetScanLine@8" ( _
           ByVal fiBitmap As Long, _
           ByVal Scanline As Long) As Long
        
        
' Conversion functions
Public Declare Function FreeImage_ConvertTo4Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo4Bits@4" ( _
           ByVal fiBitmap As Long) As Long
           
Public Declare Function FreeImage_ConvertToGreyscale Lib "FreeImage.dll" Alias "_FreeImage_ConvertToGreyscale@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits555@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo16Bits565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits565@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo24Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo24Bits@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertTo32Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo32Bits@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ColorQuantize Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantize@8" ( _
           ByVal fiBitmap As Long, _
           ByVal quantizeMethod As FREE_IMAGE_QUANTIZE) As Long
           
Public Declare Function FreeImage_ColorQuantizeExInt Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantizeEx@20" ( _
           ByVal fiBitmap As Long, _
  Optional ByVal quantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, _
  Optional ByVal paletteSize As Long = 256, _
  Optional ByVal reserveSize As Long = 0, _
  Optional ByVal reservePalettePtr As Long = 0) As Long

Private Declare Function FreeImage_Threshold Lib "FreeImage" Alias "_FreeImage_Threshold@8" ( _
           ByVal fiBitmap As Long, _
           ByVal threshold As Byte) As Long

Public Declare Function FreeImage_Dither Lib "FreeImage" Alias "_FreeImage_Dither@8" ( _
           ByVal fiBitmap As Long, _
           ByVal ditherMethod As FREE_IMAGE_DITHER) As Long

Public Declare Function FreeImage_ConvertToFloat Lib "FreeImage" Alias "_FreeImage_ConvertToFloat@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGBF Lib "FreeImage" Alias "_FreeImage_ConvertToRGBF@4" ( _
           ByVal fiBitmap As Long) As Long

'Manually patched by Tanner:
Public Declare Function FreeImage_ConvertToRGBAF Lib "FreeImage" Alias "_FreeImage_ConvertToRGBAF@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertToUINT16 Lib "FreeImage" Alias "_FreeImage_ConvertToUINT16@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGB16 Lib "FreeImage" Alias "_FreeImage_ConvertToRGB16@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_ConvertToRGBA16 Lib "FreeImage" Alias "_FreeImage_ConvertToRGBA16@4" ( _
           ByVal fiBitmap As Long) As Long
           
Public Declare Function FreeImage_GetRedMask Lib "FreeImage" Alias "_FreeImage_GetRedMask@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Function FreeImage_GetBlueMask Lib "FreeImage" Alias "_FreeImage_GetBlueMask@4" ( _
           ByVal fiBitmap As Long) As Long
           
'Red and blue masks are used to determine RGB vs BGR order.  Green masks aren't used at present.
'Public Declare Function FreeImage_GetGreenMask Lib "FreeImage.dll" Alias "_FreeImage_GetGreenMask@4" ( _
           ByVal fiBitmap as Long) As Long

' Tone mapping operators
Public Declare Function FreeImage_TmoDrago03 Lib "FreeImage" Alias "_FreeImage_TmoDrago03@20" ( _
           ByVal fiBitmap As Long, _
  Optional ByVal gamma As Double = 2.2, _
  Optional ByVal Exposure As Double) As Long

Public Declare Function FreeImage_TmoReinhard05Ex Lib "FreeImage" Alias "_FreeImage_TmoReinhard05Ex@36" ( _
           ByVal fiBitmap As Long, _
  Optional ByVal fIntensity As Double, _
  Optional ByVal fContrast As Double, _
  Optional ByVal fAdaptation As Double = 1#, _
  Optional ByVal fColorCorrection As Double) As Long

' ICC profile functions
Private Declare Function FreeImage_GetICCProfileInt Lib "FreeImage" Alias "_FreeImage_GetICCProfile@4" ( _
           ByVal fiBitmap As Long) As Long

' Plugin functions
Private Declare Function FreeImage_GetFormatFromFIFInt Lib "FreeImage" Alias "_FreeImage_GetFormatFromFIF@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Public Declare Function FreeImage_GetFIFFromFilenameU Lib "FreeImage" Alias "_FreeImage_GetFIFFromFilenameU@4" ( _
           ByVal srcFilename As Long) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_FIFSupportsReadingInt Lib "FreeImage" Alias "_FreeImage_FIFSupportsReading@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsWritingInt Lib "FreeImage" Alias "_FreeImage_FIFSupportsWriting@4" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT) As Long

Private Declare Function FreeImage_FIFSupportsExportTypeInt Lib "FreeImage" Alias "_FreeImage_FIFSupportsExportType@8" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal fiImageType As FREE_IMAGE_TYPE) As Long

Private Declare Function FreeImage_FIFSupportsExportBPPInt Lib "FreeImage" Alias "_FreeImage_FIFSupportsExportBPP@8" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal dstBPP As Long) As Long
           
' Multipage functions
Private Declare Function FreeImage_OpenMultiBitmapInt Lib "FreeImage" Alias "_FreeImage_OpenMultiBitmap@24" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal srcFilename As String, _
           ByVal createNew As Long, _
           ByVal openAsReadOnly As Long, _
           ByVal keepCacheInMemory As Long, _
           ByVal fiFlags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_CloseMultiBitmapInt Lib "FreeImage" Alias "_FreeImage_CloseMultiBitmap@8" ( _
           ByVal fiBitmap As Long, _
  Optional ByVal fiFlags As FREE_IMAGE_SAVE_OPTIONS) As Long

Public Declare Function FreeImage_GetPageCount Lib "FreeImage" Alias "_FreeImage_GetPageCount@4" ( _
           ByVal fiBitmap As Long) As Long

Public Declare Sub FreeImage_AppendPage Lib "FreeImage" Alias "_FreeImage_AppendPage@8" ( _
           ByVal fiBitmap As Long, _
           ByVal pageBitmap As Long)

Public Declare Function FreeImage_LockPage Lib "FreeImage" Alias "_FreeImage_LockPage@8" ( _
           ByVal fiBitmap As Long, _
           ByVal pageNumber As Long) As Long

Private Declare Sub FreeImage_UnlockPageInt Lib "FreeImage" Alias "_FreeImage_UnlockPage@12" ( _
           ByVal fiBitmap As Long, _
           ByVal pageBitmap As Long, _
           ByVal applyChanges As Long)

' Memory I/O streams
Private Declare Function FreeImage_OpenMemoryByPtr Lib "FreeImage" Alias "_FreeImage_OpenMemory@8" ( _
  Optional ByVal dataPtr As Long, _
  Optional ByVal sizeInBytes As Long) As Long

Public Declare Sub FreeImage_CloseMemory Lib "FreeImage" Alias "_FreeImage_CloseMemory@4" ( _
           ByVal hStream As Long)

Private Declare Function FreeImage_LoadFromMemory Lib "FreeImage" Alias "_FreeImage_LoadFromMemory@12" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal hStream As Long, _
  Optional ByVal fiFlags As FREE_IMAGE_LOAD_OPTIONS) As Long

Private Declare Function FreeImage_SaveToMemoryInt Lib "FreeImage" Alias "_FreeImage_SaveToMemory@16" ( _
           ByVal imgFormat As FREE_IMAGE_FORMAT, _
           ByVal fiBitmap As Long, _
           ByVal hStream As Long, _
  Optional ByVal fiFlags As FREE_IMAGE_SAVE_OPTIONS) As Long

Private Declare Function FreeImage_AcquireMemoryInt Lib "FreeImage" Alias "_FreeImage_AcquireMemory@12" ( _
           ByVal hStream As Long, _
           ByRef dataPtr As Long, _
           ByRef sizeInBytes As Long) As Long
           
'--------------------------------------------------------------------------------
' Toolkit functions
'--------------------------------------------------------------------------------

Private Declare Function FreeImage_FlipVertical Lib "FreeImage" Alias "_FreeImage_FlipVertical@4" (ByVal fiBitmap As Long) As Long

' Upsampling and downsampling
Public Declare Function FreeImage_Rescale Lib "FreeImage" Alias "_FreeImage_Rescale@16" ( _
           ByVal fiBitmap As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal Filter As FREE_IMAGE_FILTER) As Long
           
Private Declare Function FreeImage_MakeThumbnailInt Lib "FreeImage" Alias "_FreeImage_MakeThumbnail@12" ( _
           ByVal fiBitmap As Long, _
           ByVal maxPixelSize As Long, _
  Optional ByVal autoConvertHDRto24bpp As Long) As Long

Private Declare Function FreeImage_PreMultiplyWithAlphaInt Lib "FreeImage" Alias "_FreeImage_PreMultiplyWithAlpha@4" ( _
           ByVal fiBitmap As Long) As Long

'This function wraps FreeImage_GetInfoHeader() and returns a populated BITMAPINFOHEADER structure for a given FI-bitmap.
Public Sub FreeImage_GetInfoHeaderEx(ByVal fiBitmap As Long, ByVal ptrToBitmapInfoHeader As Long)
    Dim lpInfoHeader As Long
    lpInfoHeader = FreeImage_GetInfoHeader(fiBitmap)
    If (lpInfoHeader <> 0) Then CopyMemoryStrict ptrToBitmapInfoHeader, lpInfoHeader, 40&
End Sub

'Thin wrapper function returning a real VB Boolean value
Public Function FreeImage_Save(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal fiBitmap As Long, ByVal srcFilename As String, Optional ByVal fiFlags As FREE_IMAGE_SAVE_OPTIONS) As Boolean
    Plugin_FreeImage.InitializeFreeImage True
    FreeImage_Save = (FreeImage_SaveUInt(imgFormat, fiBitmap, StrPtr(srcFilename), fiFlags) = 1)
End Function

'FreeImage internally tracks resolution in "dots per meter"; we must convert to/from DPI accordingly
Public Function FreeImage_GetResolutionX(ByVal fiBitmap As Long) As Double
    FreeImage_GetResolutionX = (0.0254 * FreeImage_GetDotsPerMeterX(fiBitmap))
End Function

Public Sub FreeImage_SetResolutionX(ByVal fiBitmap As Long, ByVal newResolution As Double)
    FreeImage_SetDotsPerMeterX fiBitmap, Int(newResolution / 0.0254 + 0.5)
End Sub

Public Function FreeImage_GetResolutionY(ByVal fiBitmap As Long) As Double
    FreeImage_GetResolutionY = (0.0254 * FreeImage_GetDotsPerMeterY(fiBitmap))
End Function

Public Sub FreeImage_SetResolutionY(ByVal fiBitmap As Long, ByVal newResolution As Double)
    FreeImage_SetDotsPerMeterY fiBitmap, Int(newResolution / 0.0254 + 0.5)
End Sub

Public Function FreeImage_IsTransparent(ByVal fiBitmap As Long) As Boolean
    FreeImage_IsTransparent = (FreeImage_IsTransparentInt(fiBitmap) = 1)
End Function

Public Function FreeImage_FIFSupportsReading(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean
    FreeImage_FIFSupportsReading = (FreeImage_FIFSupportsReadingInt(imgFormat) = 1)
End Function

Public Function FreeImage_FlipVertically(ByVal fiBitmap As Long) As Boolean
    FreeImage_FlipVertically = (FreeImage_FlipVertical(fiBitmap) <> 0)
End Function

Public Function FreeImage_CloseMultiBitmap(ByVal fiBitmap As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean
    FreeImage_CloseMultiBitmap = (FreeImage_CloseMultiBitmapInt(fiBitmap, Flags) = 1)
End Function

Public Function FreeImage_PreMultiplyWithAlpha(ByVal fiBitmap As Long) As Boolean
    FreeImage_PreMultiplyWithAlpha = (FreeImage_PreMultiplyWithAlphaInt(fiBitmap) = 1)
End Function

Public Function FreeImage_OpenMultiBitmap(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal srcFilename As String, Optional ByVal createNew As Boolean = False, Optional ByVal fiFlags As FREE_IMAGE_LOAD_OPTIONS) As Long
    Dim lCreateNew As Long
    If createNew Then lCreateNew = 1 Else lCreateNew = 0
    FreeImage_OpenMultiBitmap = FreeImage_OpenMultiBitmapInt(imgFormat, srcFilename, lCreateNew, 0, 0, fiFlags)
End Function

Public Sub FreeImage_UnlockPage(ByVal fiBitmap As Long, ByVal fiPageBitmap As Long, ByVal applyChanges As Boolean)
    Dim lApplyChanges As Long
    If applyChanges Then lApplyChanges = 1 Else lApplyChanges = 0
    FreeImage_UnlockPageInt fiBitmap, fiPageBitmap, lApplyChanges
End Sub

Public Function FreeImage_MakeThumbnail(ByVal fiBitmap As Long, ByVal maxPixelSize As Long, Optional ByVal autoConvertHDR As Boolean = True) As Long
    Dim lConvert As Long
    If autoConvertHDR Then lConvert = 1 Else lConvert = 0
    FreeImage_MakeThumbnail = FreeImage_MakeThumbnailInt(fiBitmap, maxPixelSize, lConvert)
End Function

Public Sub FreeImage_UnloadEx(ByRef fiBitmap As Long)
    If (fiBitmap <> 0) Then FreeImage_Unload fiBitmap
    fiBitmap = 0
End Sub

'NOTE: modified by Tanner to support loading directly from pointer with known allocation size
Public Function FreeImage_LoadFromMemoryEx(ByVal lDataPtr As Long, ByVal sizeInBytes As Long, Optional ByVal fiFlags As FREE_IMAGE_LOAD_OPTIONS = 0, Optional ByRef imgFormat As FREE_IMAGE_FORMAT = FIF_UNKNOWN) As Long
    
    'Ensure library is available before proceeding
    Plugin_FreeImage.InitializeFreeImage True
    
    'Open a stream on the source pointer
    Dim hStream As Long
    hStream = FreeImage_OpenMemoryByPtr(lDataPtr, sizeInBytes)
    If (hStream <> 0) Then
        
        'On success, detect image type
        If (imgFormat = FIF_UNKNOWN) Then
            imgFormat = FreeImage_GetFileTypeFromMemory(hStream)
            PDDebug.LogAction "FreeImage_LoadFromMemoryEx auto-detected format: " & imgFormat
        End If
        
        'For safety reasons, only load known image types
        If (imgFormat <> FIF_UNKNOWN) Then FreeImage_LoadFromMemoryEx = FreeImage_LoadFromMemory(imgFormat, hStream, fiFlags)
        
        'Close FreeImage stream copy before exiting (other freeing is caller's responsibility)
        FreeImage_CloseMemory hStream
        
    '/open stream on source ptr
    End If

End Function

'This function saves a FreeImage DIB into the destination "dstData()" array.
'
'RETURNS: TRUE on success. FALSE otherwise.
Public Function FreeImage_SaveToMemoryEx(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal fiBitmap As Long, ByRef dstData() As Byte, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, Optional ByRef dstSizeInBytes As Long) As Boolean

    'Ensure library is available before proceeding
    Plugin_FreeImage.InitializeFreeImage True
    
    FreeImage_SaveToMemoryEx = False
    Dim hStream As Long, lpData As Long, lSizeInBytes As Long
    
    'Failsafe checks
    If (fiBitmap = 0) Then Exit Function
    If (Not FreeImage_HasPixels(fiBitmap)) Then Exit Function
    
    'Acquire pointer to new stream
    hStream = FreeImage_OpenMemoryByPtr(0&, 0&)
    If (hStream <> 0) Then
        
        'Perform the save
        FreeImage_SaveToMemoryEx = (FreeImage_SaveToMemoryInt(imgFormat, fiBitmap, hStream, Flags) = 1)
        If FreeImage_SaveToMemoryEx Then
        
            'Get pointer to (and size of) resulting FI stream
            If (FreeImage_AcquireMemoryInt(hStream, lpData, lSizeInBytes) <> 0) Then
            
                'Change by Tanner: return the size in bytes, and only allocate new memory as necessary.
                ' (This allows the caller to reuse allocations that may already exist.)
                dstSizeInBytes = lSizeInBytes
                If Not VBHacks.IsArrayInitialized(dstData) Then
                    ReDim dstData(lSizeInBytes - 1) As Byte
                Else
                    If UBound(dstData) < (lSizeInBytes - 1) Then ReDim dstData(0 To lSizeInBytes - 1) As Byte
                End If
                
                'Copy the contents from the FreeImage stream to the VB array
                CopyMemoryStrict VarPtr(dstData(0)), lpData, lSizeInBytes
                FreeImage_SaveToMemoryEx = True
            
            '/Couldn't acquire pointer to FI stream
            Else
                FreeImage_SaveToMemoryEx = False
            End If
        
        'FreeImage export failed
        Else
            FreeImage_SaveToMemoryEx = False
        End If
        
        'Ensure FreeImage copy of the data is also released
        FreeImage_CloseMemory hStream
        
    '/FreeImage produced a useable destination stream for the save
    End If
    
End Function

Public Function FreeImage_HasICCProfile(ByVal fiBitmap As Long) As Boolean
    FreeImage_HasICCProfile = (FreeImage_GetICCProfileSize(fiBitmap) <> 0)
End Function

Public Function FreeImage_GetICCProfileSize(ByVal fiBitmap As Long) As Long
    Dim tmpProfileHeader As FIICCPROFILE
    CopyMemoryStrict VarPtr(tmpProfileHeader), FreeImage_GetICCProfileInt(fiBitmap), LenB(tmpProfileHeader)
    FreeImage_GetICCProfileSize = tmpProfileHeader.Size
End Function

Public Function FreeImage_GetICCProfileDataPointer(ByVal fiBitmap As Long) As Long
    Dim tmpProfileHeader As FIICCPROFILE
    CopyMemoryStrict VarPtr(tmpProfileHeader), FreeImage_GetICCProfileInt(fiBitmap), LenB(tmpProfileHeader)
    FreeImage_GetICCProfileDataPointer = tmpProfileHeader.Data
End Function

'ADDED BY TANNER:
Public Function FreeImage_GetPalette_ByTanner(ByVal fiBitmap As Long, ByRef dstQuad() As RGBQuad, ByRef numOfColors As Long) As Boolean
    
    FreeImage_GetPalette_ByTanner = False
    
    'Validate handle
    If (fiBitmap <> 0) Then
        
        'Validate color count and palette existence
        Dim hPalette As Long
        hPalette = FreeImage_GetPalette(fiBitmap)
        numOfColors = FreeImage_GetColorsUsed(fiBitmap)
        
        If (numOfColors > 0) And (hPalette <> 0) Then
            FreeImage_GetPalette_ByTanner = True
            ReDim dstQuad(0 To numOfColors - 1) As RGBQuad
            CopyMemoryStrict VarPtr(dstQuad(0)), hPalette, numOfColors * 4
        End If
        
    End If
    
End Function

'Is a given FreeImage bitmap grayscale?
' (Only works on 8-bpp images; HDR images rely on internal FreeImage color mode values).
Private Function FreeImage_IsGreyscaleImage(ByVal fiBitmap As Long) As Boolean
    
    If (FreeImage_GetBPP(fiBitmap) <= 8) Then
        
        Dim imgPalette() As RGBQuad, numColors As Long
        If FreeImage_GetPalette_ByTanner(fiBitmap, imgPalette, numColors) Then
            
            Dim i As Long
            For i = 0 To numColors - 1
                If (imgPalette(i).Red <> imgPalette(i).Green) Or (imgPalette(i).Red <> imgPalette(i).Blue) Then
                    FreeImage_IsGreyscaleImage = False
                    Exit For
                End If
            Next i
            
        End If
    
    End If
    
End Function

'Returns a new FreeImage bitmap in the specified color format.  If unloadSource is TRUE, the passed bitmap handle
' will be freed and forcibly set to 0.
Public Function FreeImage_ConvertColorDepth(ByRef fiBitmap As Long, ByVal conversionType As FREE_IMAGE_CONVERSION_FLAGS, Optional ByVal unloadSource As Boolean = False) As Long
    
    'Perform basic validation before continuing
    Plugin_FreeImage.InitializeFreeImage True
    If (fiBitmap = 0) Then Exit Function
    If (Not FreeImage_HasPixels(fiBitmap)) Then Exit Function
    
    'These optional settings were originally handled as optional params.
    ' PD always uses their default settings, so I have rewritten them as local values.
    Const GRAY_THRESHOLD As Long = 128
    
    Dim DITHER_METHOD As FREE_IMAGE_DITHER
    DITHER_METHOD = FID_FS
    
    Dim QUANTIZE_METHOD As FREE_IMAGE_QUANTIZE
    QUANTIZE_METHOD = FIQ_WUQUANT
    
    'If a new FreeImage bitmap is produced by this function (as the result of a color mode transform),
    ' this value will be non-zero.
    Dim hDIBNew As Long
    
    'Temporary handles may be required for some intermediary transforms
    Dim hDIBTemp As Long
    
    'Current color depth of the source bitmap
    Dim lBPP As Long
    lBPP = FreeImage_GetBPP(fiBitmap)
    
    'Grayscale images are typically forcibly converted to ensure a linear grayscale palette,
    ' but this can be overridden by incoming conversion params.  A corresponding comment from the
    ' FreeImage docs (v3) says this:
    '
    ' NB: here “greyscale” means that the resulting bitmap will have grey colors, but the palette
    ' won’t be a linear greyscale palette. Thus, FreeImage_GetColorType will return FIC_PALETTE."
    '
    'We apply extra checks for this case, and apply manual grayscale conversions as necessary.
    Dim forceGrayPalette As Boolean
    forceGrayPalette = ((conversionType And FICF_REORDER_GREYSCALE_PALETTE) = 0)
    
    'Ignore grayscale reordering when comparing flags
    Select Case (conversionType And (Not FICF_REORDER_GREYSCALE_PALETTE))
      
        Case FICF_MONOCHROME, FICF_MONOCHROME_THRESHOLD
            If (lBPP > 1) Then hDIBNew = FreeImage_Threshold(fiBitmap, GRAY_THRESHOLD)
         
        Case FICF_MONOCHROME_DITHER
            If (lBPP > 1) Then hDIBNew = FreeImage_Dither(fiBitmap, DITHER_METHOD)
        
        'Note the extra branches for the "forceGrayPalette" parameter.  If the source image is already in a
        ' palette-based mode, but *not* explicitly marked as grayscale, we'll forcibly convert it to a
        ' grayscale-specific mode to ensure a fixed [0, 255] linear grayscale palette for the colors.
        '
        'From the FreeImage docs (v3):
        '
        ' "Converts a bitmap to 4 bits. If the bitmap was a high-color bitmap (16, 24 or 32-bit) or if it was
        '  a monochrome or greyscale bitmap (1 or 8-bit), the end result will be a greyscale bitmap,
        '  otherwise (1-bit palletised bitmaps) it will be a palletised bitmap. A clone of the input bitmap is
        '  returned for 4-bit bitmaps."
        Case FICF_GREYSCALE_4BPP
            
            If (lBPP <> 4) Then
                
                'Monochrome mode does not guarantee grayscale values, and the built-in FreeImage
                ' color-depth "upscaler" will simply retain existing colors.  So we must force to
                ' grayscale *before* upscaling the depth.
                If ((lBPP = 1) And (FreeImage_GetColorType(fiBitmap) = FIC_PALETTE)) Then
                    hDIBTemp = FreeImage_ConvertToGreyscale(fiBitmap)
                    hDIBNew = FreeImage_ConvertTo4Bits(hDIBTemp)
                    FreeImage_Unload hDIBTemp
                Else
                    hDIBNew = FreeImage_ConvertTo4Bits(fiBitmap)
                End If
                
            Else
                
                'Check for existing grayscale before converting
                If (((Not forceGrayPalette) And (Not FreeImage_IsGreyscaleImage(fiBitmap))) Or _
                (forceGrayPalette And (FreeImage_GetColorType(fiBitmap) = FIC_PALETTE))) Then
                    hDIBTemp = FreeImage_ConvertToGreyscale(fiBitmap)
                    hDIBNew = FreeImage_ConvertTo4Bits(hDIBTemp)
                    FreeImage_Unload hDIBTemp
                End If
                
            End If
            
        Case FICF_GREYSCALE_8BPP
            
            'Look for 8-bpp gray before converting
            If ((lBPP <> 8) Or _
               ((Not forceGrayPalette) And (Not FreeImage_IsGreyscaleImage(fiBitmap)) Or _
               (forceGrayPalette And (FreeImage_GetColorType(fiBitmap) = FIC_PALETTE)))) Then
                hDIBNew = FreeImage_ConvertToGreyscale(fiBitmap)
            End If
         
        Case FICF_PALLETISED_8BPP
        
            If (lBPP <> 8) Then
                
                '24/32-bpp can be directly quantized; other color-depths must be converted to 24-bpp first.
                If (lBPP = 24) Then
                    hDIBNew = FreeImage_ColorQuantize(fiBitmap, QUANTIZE_METHOD)
                Else
                    hDIBTemp = FreeImage_ConvertTo24Bits(fiBitmap)
                    hDIBNew = FreeImage_ColorQuantize(hDIBTemp, QUANTIZE_METHOD)
                    FreeImage_Unload hDIBTemp
                End If
                
            End If
         
        Case FICF_RGB_15BPP
            If (lBPP <> 15) Then hDIBNew = FreeImage_ConvertTo16Bits555(fiBitmap)
        
        Case FICF_RGB_16BPP
            If (lBPP <> 16) Then hDIBNew = FreeImage_ConvertTo16Bits565(fiBitmap)
         
        Case FICF_RGB_24BPP
            If (lBPP <> 24) Then hDIBNew = FreeImage_ConvertTo24Bits(fiBitmap)
         
        Case FICF_RGB_32BPP
            If (lBPP <> 32) Then hDIBNew = FreeImage_ConvertTo32Bits(fiBitmap)
         
    End Select
      
    'If we had to generate a new image, free the old one now
    If (hDIBNew <> 0) Then
        FreeImage_ConvertColorDepth = hDIBNew
        If unloadSource Then FreeImage_UnloadEx fiBitmap
    Else
        FreeImage_ConvertColorDepth = fiBitmap
    End If
    
End Function

Public Function FreeImage_CreateFromDC(ByVal hDC As Long) As Long
    
    'Ensure library is available before proceeding
    Plugin_FreeImage.InitializeFreeImage True
    FreeImage_CreateFromDC = False
    
    'Retrieve the bitmap (if any) currently selected into the DC
    Dim hBitmap As Long
    Const OBJ_BITMAP As Long = 7
    
    hBitmap = GetCurrentObject(hDC, OBJ_BITMAP)
    If (hBitmap = 0) Then Exit Function
    
    'Retrieve bitmap header
    Dim tBM As Bitmap_API
    If (GetObjectAPI(hBitmap, Len(tBM), tBM) <> 0) Then
        
        'Allocate FreeImage DIB handle
        Dim hFreeImageDIB As Long
        hFreeImageDIB = FreeImage_Allocate(tBM.bmWidth, tBM.bmHeight, tBM.bmBitsPixel)
        If (hFreeImageDIB <> 0) Then
            
            'The GetDIBits function clears the biClrUsed and biClrImportant BitmapINFO members (don't know why).
            ' So save these in case they are needed later (for palletized images only).
            Dim nColors As Long
            nColors = FreeImage_GetColorsUsed(hFreeImageDIB)
            
            Const DIB_RGB_COLORS As Long = 0
            If (GetDIBits(hDC, hBitmap, 0, FreeImage_GetHeight(hFreeImageDIB), FreeImage_GetBits(hFreeImageDIB), FreeImage_GetInfo(hFreeImageDIB), DIB_RGB_COLORS) <> 0) Then
                
                'For palette images, restore number of colors used if relevant
                If (nColors <> 0) Then
                    Dim lpInfo As Long
                    lpInfo = FreeImage_GetInfo(hFreeImageDIB)
                    CopyMemoryStrict lpInfo + 32, VarPtr(nColors), 4
                End If
                
                'Return the FreeImage bitmap handle
                FreeImage_CreateFromDC = hFreeImageDIB
            
            'GetDIBits failed
            Else
                FreeImage_UnloadEx hFreeImageDIB
            End If
      
      '/FreeImage bitmap allocation succeeded
      End If
   
   '/GetObject succeeded
   End If
   
End Function

Public Function FreeImage_SaveEx(ByVal fiBitmap As Long, ByVal dstFilename As String, ByVal imgFormat As FREE_IMAGE_FORMAT, _
                        Optional ByVal saveOptions As FREE_IMAGE_SAVE_OPTIONS, _
                        Optional ByVal colorDepth As FREE_IMAGE_COLOR_DEPTH = FICD_AUTO, _
                        Optional ByVal unloadSource As Boolean = False) As Boolean

    'Ensure library is available before proceeding
    Plugin_FreeImage.InitializeFreeImage True
    FreeImage_SaveEx = False
    
    'Failsafe checks
    If (fiBitmap = 0) Then Exit Function
    If (Not FreeImage_HasPixels(fiBitmap)) Then Exit Function
    If (imgFormat = FIF_UNKNOWN) Then Exit Function
    If (Not FreeImage_FIFSupportsWriting(imgFormat)) Then Exit Function
    
    'If the caller doesn't know what color-depth they want, try to choose a good one for them.
    Dim lBPP As Long, lBppOriginal As Long
    lBPP = FreeImage_GetBPP(fiBitmap)
    lBppOriginal = lBPP
    
    'Custom color-depths may require us to create an intermediary DIB.
    Dim createdNewDIB As Boolean
    createdNewDIB = False
    
    If (colorDepth = FICD_AUTO) Then
        
        'See if the current image color-depth is supported.
        If (Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) Then
            
            'Current color-depth is not supported.  Find the next-highest one that is.
            Do
                lBPP = GetLargerColorDepth(lBPP)
            Loop While (Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) And (lBPP < 32)
            
            'See if the newly selected color-depth is supported.
            If (Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) Then
                
                'Iterate through *smaller* color-depths until we find one that works.
                Do
                    lBPP = GetSmallerColorDepth(lBPP)
                Loop While (Not FreeImage_FIFSupportsExportBPP(imgFormat, lBPP)) And (lBPP > 1)
                
                'Hopefully we found a color-depth that works!
                
            End If
            
            'Attempt to convert to the target color-depth
            If (lBPP >= 1) And (lBPP <= 32) Then
                fiBitmap = FreeImage_ConvertColorDepth(fiBitmap, lBPP, unloadSource)
                createdNewDIB = True
            Else
                PDDebug.LogAction "Couldn't find acceptable auto color-depth for FreeImage_SaveEx"
                Exit Function
            End If
        
        '/current color-depth is supported
        End If
    
    'Caller requested their own color-depth
    Else
        
        'Mask out possible FICD_MONOCHROME_DITHER, which is 0b11
        Dim maskedDepth As Long
        maskedDepth = colorDepth And (Not &H2&)
        
        'Ensure color-depth is supported
        If (Not FreeImage_FIFSupportsExportBPP(imgFormat, maskedDepth)) Then
            PDDebug.LogAction "Incompatible color-depth in FreeImage_SaveEx"
            Exit Function
        End If
        
        'Target color-depth is supported.  Ensure bitmap is in that space.
        If (FreeImage_GetBPP(fiBitmap) <> colorDepth) Then
            fiBitmap = FreeImage_ConvertColorDepth(fiBitmap, colorDepth, unloadSource)
            createdNewDIB = True
        End If
        
    End If
    
    'We now guarantee that fiBitmap is in a compatible color-space for export to the target format.
    
    'Perform the save
    FreeImage_SaveEx = FreeImage_Save(imgFormat, fiBitmap, dstFilename, saveOptions)
    
    'Free FI bitmap handle as requested
    If (createdNewDIB Or unloadSource) Then FreeImage_UnloadEx fiBitmap
    
End Function

'This function returns the result of the 'FreeImage_GetFormatFromFIF' function as a BSTR.
' (The 'Format' parameter works according to the FreeImage 3 API documentation.)
Private Function FreeImage_GetFormatFromFIF(ByVal imgFormat As FREE_IMAGE_FORMAT) As String
   FreeImage_GetFormatFromFIF = Strings.StringFromCharPtr(FreeImage_GetFormatFromFIFInt(imgFormat), False)
End Function

Private Function FreeImage_HasPixels(ByVal fiBitmap As Long) As Boolean
    FreeImage_HasPixels = (FreeImage_HasPixelsInt(fiBitmap) = 1)
End Function

Private Function FreeImage_FIFSupportsWriting(ByVal imgFormat As FREE_IMAGE_FORMAT) As Boolean
    FreeImage_FIFSupportsWriting = (FreeImage_FIFSupportsWritingInt(imgFormat) = 1)
End Function

Private Function FreeImage_FIFSupportsExportType(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal imgType As FREE_IMAGE_TYPE) As Boolean
   FreeImage_FIFSupportsExportType = (FreeImage_FIFSupportsExportTypeInt(imgFormat, imgType) = 1)
End Function

Private Function FreeImage_FIFSupportsExportBPP(ByVal imgFormat As FREE_IMAGE_FORMAT, ByVal imgBPP As Long) As Boolean
    FreeImage_FIFSupportsExportBPP = (FreeImage_FIFSupportsExportBPPInt(imgFormat, imgBPP) = 1)
End Function

'Get the next-smallest color depth (< 32-bpp)
Private Function GetSmallerColorDepth(ByVal srcBPP As Long) As Long
    Select Case srcBPP
        Case 32
            GetSmallerColorDepth = 24
        Case 24
            GetSmallerColorDepth = 16
        Case 16
            GetSmallerColorDepth = 15
        Case 15
            GetSmallerColorDepth = 8
        Case 8
            GetSmallerColorDepth = 4
        Case 4
            GetSmallerColorDepth = 1
        Case Else
            GetSmallerColorDepth = 1
    End Select
End Function

'Get the next-largest color depth (< 32-bpp)
Private Function GetLargerColorDepth(ByVal srcBPP As Long) As Long
    Select Case srcBPP
        Case 1
            GetLargerColorDepth = 4
        Case 4
            GetLargerColorDepth = 8
        Case 8
            GetLargerColorDepth = 15
        Case 15
            GetLargerColorDepth = 16
        Case 16
            GetLargerColorDepth = 24
        Case 24
            GetLargerColorDepth = 32
        Case Else
            GetLargerColorDepth = 32
   End Select
End Function

'Equivalent of C "return *(ptr);"
Private Function pDeref(ByVal ptr As Long) As Long
    CopyMemoryStrict VarPtr(pDeref), ptr, 4
End Function
