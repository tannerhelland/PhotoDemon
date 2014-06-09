Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright ©2012-2014 by Tanner Helland
'Created: 1/September/12
'Last updated: 09/June/14
'Last update: new GDIPlusFillDIBRect function
'
'This interface provides a means for interacting with various GDI+ features.  GDI+ was originally used as a fallback for image loading
' and saving if the FreeImage DLL was not found, but over time it has become more and more integrated into PD.  As of version 6.0, GDI+
' is used for a number of specialized tasks, including viewport rendering of 32bpp images, regional blur of selection masks, antialiased
' lines and circles on various dialogs, and more.
'
'These routines are adapted from the work of a number of other talented VB programmers.  Since GDI+ is not well-documented
' for VB users, I first pieced this module together from the following pieces of code:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP implementation: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Robert Rayment's PaintRR implementation: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
' Many thanks to these individuals for their outstanding work on graphics in VB.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'GDI+ Enums
Public Enum GDIPlusImageFormat
    [ImageBMP] = 0
    [ImageGIF] = 1
    [ImageJPEG] = 2
    [ImagePNG] = 3
    [ImageTIFF] = 4
End Enum

Public Enum GDIPlusStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum

Private Enum EncoderValue
    [EncoderValueColorTypeCMYK] = 0
    [EncoderValueColorTypeYCCK] = 1
    [EncoderValueCompressionLZW] = 2
    [EncoderValueCompressionCCITT3] = 3
    [EncoderValueCompressionCCITT4] = 4
    [EncoderValueCompressionRle] = 5
    [EncoderValueCompressionNone] = 6
    [EncoderValueScanMethodInterlaced] = 7
    [EncoderValueScanMethodNonInterlaced] = 8
    [EncoderValueVersionGif87] = 9
    [EncoderValueVersionGif89] = 10
    [EncoderValueRenderProgressive] = 11
    [EncoderValueRenderNonProgressive] = 12
    [EncoderValueTransformRotate90] = 13
    [EncoderValueTransformRotate180] = 14
    [EncoderValueTransformRotate270] = 15
    [EncoderValueTransformFlipHorizontal] = 16
    [EncoderValueTransformFlipVertical] = 17
    [EncoderValueMultiFrame] = 18
    [EncoderValueLastFrame] = 19
    [EncoderValueFlush] = 20
    [EncoderValueFrameDimensionTime] = 21
    [EncoderValueFrameDimensionResolution] = 22
    [EncoderValueFrameDimensionPage] = 23
    [EncoderValueColorTypeGray] = 24
    [EncoderValueColorTypeRGB] = 25
End Enum

Private Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
    ClassID           As CLSID
    FormatID          As CLSID
    CodecName         As Long
    DllName           As Long
    formatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

Private Type EncoderParameter
    Guid           As CLSID
    NumberOfValues As Long
    encType           As EncoderParameterValueType
    Value          As Long
End Type

Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type

Private Const EncoderCompression      As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const EncoderColorDepth       As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
'Private Const EncoderColorSpace       As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
'Private Const EncoderScanMethod       As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const EncoderVersion          As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
'Private Const EncoderRenderMethod     As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const EncoderQuality          As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
'Private Const EncoderTransformation   As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
'Private Const EncoderLuminanceTable   As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
'Private Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
'Private Const EncoderSaveAsCMYK       As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"
'Private Const EncoderSaveFlag         As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
'Private Const CodecIImageBytes        As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"

'GDI+ recognizes a variety of pixel formats, but we are only concerned with the ones relevant to PhotoDemon:
Private Const PixelFormat24bppRGB = &H21808
Private Const PixelFormat32bppARGB = &H26200A
Public Const PixelFormat32bppPARGB = &HE200B
Private Const PixelFormatAlpha = &H40000
Private Const PixelFormatPremultipliedAlpha = &H80000
Private Const PixelFormat32bppCMYK = &H200F

'Now that PD supports the loading of ICC profiles, we use this constant to retrieve it
Private Const PropertyTagICCProfile As Long = &H8773&

'LockBits constants
Private Const ImageLockModeRead = &H1
Private Const ImageLockModeWrite = &H2
Private Const ImageLockModeUserInputBuf = &H4

'GDI+ supports a variety of different linecaps.  Anchor caps will center the cap at the end of the line.
Public Enum LineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3
   LineCapNoAnchor = &H10
   LineCapSquareAnchor = &H11
   LineCapRoundAnchor = &H12
   LineCapDiamondAnchor = &H13
   LineCapArrowAnchor = &H14
   LineCapCustom = &HFF
   LineCapAnchorMask = &HF0
End Enum

Public Enum DashStyle
   DashStyleSolid = 0
   DashStyleDash = 1
   DashStyleDot = 2
   DashStyleDashDot = 3
   DashStyleDashDotDot = 4
   DashStyleCustom = 5
End Enum

' Dash cap constants
Public Enum DashCap
   DashCapFlat = 0
   DashCapRound = 2
   DashCapTriangle = 3
End Enum

'GDI+ required types
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'GDI+ image properties
Private Type PropertyItem
   propId As Long              ' ID of this property
   propLength As Long              ' Length of the property value, in bytes
   propType As Integer             ' Type of the value, as one of TAG_TYPE_XXX
   propValue As Long               ' property value
End Type

' Image property types
Private Const PropertyTagTypeByte = 1
Private Const PropertyTagTypeASCII = 2
Private Const PropertyTagTypeShort = 3
Private Const PropertyTagTypeLong = 4
Private Const PropertyTagTypeRational = 5
Private Const PropertyTagTypeUndefined = 7
Private Const PropertyTagTypeSLONG = 9
Private Const PropertyTagTypeSRational = 10

'OleCreatePictureIndirect types
Private Type PictDesc
    Size       As Long
    picType       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type

'BITMAP types
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

'Necessary to check for v1.1 of the GDI+ dll
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    
'Start-up and shutdown
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef Token As Long, ByRef inputbuf As GdiplusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GDIPlusStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GDIPlusStatus

'Load image from file, process said file, etc.
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef gpImage As Long) As Long
Private Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal srcFilename As String, ByRef gpImage As Long) As Long
Private Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal gpBitmap As Long, ByRef gpFlags As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal iPixelFormat As Long, ByVal srcBitmap As Long, ByRef dstBitmap As Long) As GDIPlusStatus
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal lStride As Long, ByVal ePixelFormat As Long, ByRef Scan0 As Any, ByRef pBitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gpBitmap As Long, hBmpReturn As Long, ByVal RGBABackground As Long) As GDIPlusStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GDIPlusStatus
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, BITMAP As Long) As GDIPlusStatus
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GDIPlusStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As GDIPlusStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As CLSID, encoderParams As Any) As GDIPlusStatus
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef imgWidth As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef imgHeight As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, ByRef imgWidth As Single, ByRef imgHeight As Single) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, ByRef imgPixelFormat As Long) As Long
Private Declare Function GdipGetDC Lib "gdiplus" (ByVal mGraphics As Long, ByRef hDC As Long) As Long
Private Declare Function GdipReleaseDC Lib "gdiplus" (ByVal mGraphics As Long, ByVal hDC As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal gdipBitmap As Long, gdipRect As RECTL, ByVal gdipFlags As Long, ByVal iPixelFormat As Long, LockedBitmapData As BitmapData) As GDIPlusStatus
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal gdipBitmap As Long, LockedBitmapData As BitmapData) As GDIPlusStatus

'Retrieve properties from an image
'Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal propId As Long, ByVal propSize As Long, ByRef mBuffer As PropertyItem) As Long
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal propId As Long, ByVal propSize As Long, ByRef mBuffer As Long) As Long
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal hImage As Long, ByVal propId As Long, propSize As Long) As Long
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef hResolution As Single) As Long
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef vResolution As Single) As Long

'OleCreatePictureIndirect is used to convert GDI+ images to VB's preferred StdPicture
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

'CLSIDFromString is used to convert a mimetype into a CLSID required by the GDI+ image encoder
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pclsid As CLSID) As Long

'Necessary for converting between ASCII and UNICODE strings
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
'Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

'CopyMemory
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal cb As Long) As Long

'GDI+ calls related to drawing lines and various shapes
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef mGraphics As Long) As Long
'Private Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal srcGraphics As Long, ByRef dstBitmap As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal mSmoothingMode As SmoothingMode) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal x As Single, ByVal y As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal mBrushMode As GDIFillMode, mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal mPath As Long) As Long
'Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal mPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal mPath As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As GpUnit, mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef mEffect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal mEffect As Long, ByRef eParams As Any, ByVal Size As Long) As Long
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal mEffect As Long) As Long
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal x As Single, ByVal y As Single) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal x As Single, ByVal y As Single, ByVal iWidth As Single, ByVal iHeight As Single) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByRef iSource As RECTF, ByVal xForm As Long, ByVal mEffect As Long, ByVal mImageAttributes As Long, ByVal srcUnit As Long) As Long
Private Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef mMatrix As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal mMatrix As Long) As Long
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal mPen As Long, ByVal startCap As LineCap, ByVal endCap As LineCap, ByVal dCap As DashCap) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal mInterpolation As InterpolationMode) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal mCompositingMode As CompositingMode) As Long
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal mGraphics As Long, ByVal mCompositingQuality As CompositingQuality) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef hImageAttr As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImageAttr As Long) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal hImageAttr As Long, ByVal mWrap As WrapMode, ByVal argbConst As Long, ByVal bClamp As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal pixOffsetMode As PixelOffsetMode) As Long

'Transforms
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal mGraphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal mGraphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long

'Helpful GDI functions for moving image data between GDI and GDI+
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Quality mode constants (only supported by certain functions!)
Private Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       'Best performance
   QualityModeHigh = 2      'Best rendering quality
End Enum

'Instead of specifying certain smoothing modes, quality modes (see above) can be used instead.
Private Enum SmoothingMode
   SmoothingModeInvalid = QualityModeInvalid
   SmoothingModeDefault = QualityModeDefault
   SmoothingModeHighSpeed = QualityModeLow
   SmoothingModeHighQuality = QualityModeHigh
   SmoothingModeNone = 3
   SmoothingModeAntiAlias = 4
End Enum

Public Enum InterpolationMode
   InterpolationModeInvalid = QualityModeInvalid
   InterpolationModeDefault = QualityModeDefault
   InterpolationModeLowQuality = QualityModeLow
   InterpolationModeHighQuality = QualityModeHigh
   InterpolationModeBilinear
   InterpolationModeBicubic
   InterpolationModeNearestNeighbor
   InterpolationModeHighQualityBilinear
   InterpolationModeHighQualityBicubic
End Enum

'Alpha compositing options; note that Over will apply alpha blending, while Copy will not
Private Enum CompositingMode
   CompositingModeSourceOver = 0
   CompositingModeSourceCopy = 1
End Enum

'Alpha compositing qualities, which in turn affect how carefully GDI+ will blend the pixels.  Use with caution!
Public Enum CompositingQuality
   CompositingQualityInvalid = QualityModeInvalid
   CompositingQualityDefault = QualityModeDefault
   CompositingQualityHighSpeed = QualityModeLow
   CompositingQualityHighQuality = QualityModeHigh
   CompositingQualityGammaCorrected
   CompositingQualityAssumeLinear
End Enum

'Wrap modes, which control the way GDI+ handles pixels that lie outside image boundaries.  (These are similar to
' the pdFilterSupport class used by many of PhotoDemon's distort filters.)
Public Enum WrapMode
   WrapModeTile = 0
   WrapModeTileFlipX = 1
   WrapModeTileFlipY = 2
   WrapModeTileFlipXY = 3
   WrapModeClamp = 4
End Enum

'PixelOffsetMode controls how GDI+ attempts to antialias objects.  For cheap antialiasing, use PixelOffsetModeHalf.
' This provides a good estimation of AA, without actually applying a full AA operation.
Public Enum PixelOffsetMode
   PixelOffsetModeInvalid = QualityModeInvalid
   PixelOffsetModeDefault = QualityModeDefault
   PixelOffsetModeHighSpeed = QualityModeLow
   PixelOffsetModeHighQuality = QualityModeHigh
   PixelOffsetModeNone = 3
   PixelOffsetModeHalf = 4
End Enum

Private Enum GDIFillMode
   FillModeAlternate = 0
   FillModeWinding = 1
End Enum

Public Enum GpUnit
   UnitWorld = 0
   UnitDisplay = 1
   UnitPixel = 2
   UnitPoint = 3
   UnitInch = 4
   UnitDocument = 5
   UnitMillimeter = 6
End Enum

Private Type BlurParams
  bRadius As Single
  ExpandEdge As Long
End Type

Private Type tmpLong
    lngResult As Long
End Type

Private Type RECTF
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
End Type

Private Type RECTL
    Left        As Long
    Top         As Long
    Width       As Long
    Height      As Long
End Type

' Information about image pixel data
Private Type BitmapData
   Width As Long
   Height As Long
   Stride As Long
   PixelFormat As Long
   Scan0 As Long
   Reserved As Long
End Type

'When GDI+ is initialized, it will assign us a token.  We use this to release GDI+ when the program terminates.
Public g_GDIPlusToken As Long

'GDI+ v1.1 allows for advanced fx work.  When we initialize GDI+, check the availability of version 1.1.
Public g_GDIPlusFXAvailable As Boolean

'Use GDI+ to resize a DIB.  (Technically, to copy a resized portion of a source image into a destination image.)
' The call is formatted similar to StretchBlt, as it used to replace StretchBlt when working with 32bpp data.
' FOR FUTURE REFERENCE: after a bunch of profiling on my Win 7 PC, I can state with 100% confidence that
' the HighQualityBicubic interpolation mode is actually the fastest mode for downsizing 32bpp images.  I have no idea
' why this is, but many, many iterative tests confirmed it.  Stranger still, in descending order after that, the fastest
' algorithms are: HighQualityBilinear, Bilinear, Bicubic.  Regular bicubic interpolation is some 4x slower than the
' high quality mode!!
Public Function GDIPlusResizeDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal interpolationType As InterpolationMode) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    GDIPlusResizeDIB = True

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    If srcDIB.getDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GdipCreateBitmapFromScan0 srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBWidth * 4, PixelFormat32bppPARGB, ByVal srcDIB.getActualDIBBits, tBitmap
        
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.getDIBColorDepth
            .Width = srcDIB.getDIBWidth
            .Height = -srcDIB.getDIBHeight
        End With
        GdipCreateBitmapFromGdiDib imgHeader, ByVal srcDIB.getActualDIBBits, tBitmap
        
    End If
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(iGraphics, interpolationType) = 0 Then
    
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        GdipSetImageAttributesWrapMode imgAttributesHandle, WrapModeTileFlipXY, 0&, 0&
        
        'To improve performance, explicitly request high-speed alpha compositing operation
        GdipSetCompositingQuality iGraphics, CompositingQualityHighSpeed
        
        'PixelOffsetMode doesn't seem to affect rendering speed more than < 5%, but I did notice a slight
        ' improvement from explicitly requesting HighQuality mode - so why not leave it?
        GdipSetPixelOffsetMode iGraphics, PixelOffsetModeHighQuality
    
        'Perform the resize
        If GdipDrawImageRectRectI(iGraphics, tBitmap, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, imgAttributesHandle) <> 0 Then
            GDIPlusResizeDIB = False
        End If
        
        'Release our image attributes object
        GdipDisposeImageAttributes imgAttributesHandle
        
    Else
        GDIPlusResizeDIB = False
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics iGraphics
    GdipDisposeImage tBitmap
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

'Use GDI+ to rotate a DIB.  (Technically, to copy a rotated portion of a source image into a destination image.)
' The function currently expects the rotation to occur around the center point of the source image.  Unlike the various
' size interaction calls in this module, all (x,y) coordinate pairs refer to the CENTER of the image, not the top-left
' corner.  This was a deliberate decision to make copying rotated data easier.
Public Function GDIPlusRotateDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rotationAngle As Single, ByVal interpolationType As InterpolationMode) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    GDIPlusRotateDIB = True

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    If srcDIB.getDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GdipCreateBitmapFromScan0 srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBWidth * 4, PixelFormat32bppPARGB, ByVal srcDIB.getActualDIBBits, tBitmap
        
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.getDIBColorDepth
            .Width = srcDIB.getDIBWidth
            .Height = -srcDIB.getDIBHeight
        End With
        GdipCreateBitmapFromGdiDib imgHeader, ByVal srcDIB.getActualDIBBits, tBitmap
        
    End If
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(iGraphics, interpolationType) = 0 Then
    
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        GdipSetImageAttributesWrapMode imgAttributesHandle, WrapModeTileFlipXY, 0&, 0&
        
        'To improve performance, explicitly request high-speed alpha compositing operation
        GdipSetCompositingQuality iGraphics, CompositingQualityHighSpeed
        
        'PixelOffsetMode doesn't seem to affect rendering speed more than < 5%, but I did notice a slight
        ' improvement from explicitly requesting HighQuality mode - so why not leave it?
        GdipSetPixelOffsetMode iGraphics, PixelOffsetModeHighQuality
    
        'Lock the incoming angle to something in the range [-360, 360]
        'rotationAngle = rotationAngle + 180
        If (rotationAngle <= -360) Or (rotationAngle >= 360) Then rotationAngle = (Int(rotationAngle) Mod 360) + (rotationAngle - Int(rotationAngle))
        
        'Perform the rotation
        
        'Transform the destination world matrix twice: once for the rotation angle, and once again to offset all coordinates.
        ' This allows us to rotate the image around its *center* rather than around its top-left corner.
        If GdipRotateWorldTransform(iGraphics, rotationAngle, 0&) = 0 Then
            If GdipTranslateWorldTransform(iGraphics, dstX + dstWidth / 2, dstY + dstHeight / 2, 1&) = 0 Then
        
                'Render the image onto the destination
                If GdipDrawImageRectRectI(iGraphics, tBitmap, -dstWidth / 2, -dstHeight / 2, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, imgAttributesHandle) <> 0 Then
                    GDIPlusRotateDIB = False
                End If
                
            Else
                GDIPlusRotateDIB = False
            End If
        Else
            GDIPlusRotateDIB = False
        End If
        
        'Release our image attributes object
        GdipDisposeImageAttributes imgAttributesHandle
        
    Else
        GDIPlusRotateDIB = False
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics iGraphics
    GdipDisposeImage tBitmap
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

'Use GDI+ to blur a DIB with variable radius
Public Function GDIPlusBlurDIB(ByRef dstDIB As pdDIB, ByVal blurRadius As Long, ByVal rLeft As Double, ByVal rTop As Double, ByVal rWidth As Double, ByVal rHeight As Double) As Boolean

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
    
    'Next, we need a temporary copy of the image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    
    If dstDIB.getDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GdipCreateBitmapFromScan0 dstDIB.getDIBWidth, dstDIB.getDIBHeight, dstDIB.getDIBWidth * 4, PixelFormat32bppARGB, ByVal dstDIB.getActualDIBBits, tBitmap
    
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = dstDIB.getDIBColorDepth
            .Width = dstDIB.getDIBWidth
            .Height = -dstDIB.getDIBHeight
        End With
        GdipCreateBitmapFromGdiDib imgHeader, ByVal dstDIB.getActualDIBBits, tBitmap
        
    End If
        
    'Create a GDI+ blur effect object
    Dim hEffect As Long
    If GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, hEffect) = 0 Then
        
        'Next, create a compatible set of blur parameters and pass those to the GDI+ blur object
        Dim tmpParams As BlurParams
        tmpParams.bRadius = CSng(blurRadius)
        tmpParams.ExpandEdge = 0
    
        If GdipSetEffectParameters(hEffect, tmpParams, Len(tmpParams)) = 0 Then
    
            'The DrawImageFX call requires a target rect.  Create one now (in GDI+ format, e.g. RECTF)
            Dim tmpRect As RECTF
            tmpRect.Left = rLeft
            tmpRect.Top = rTop
            tmpRect.Width = rWidth
            tmpRect.Height = rHeight
            
            'Create a temporary GDI+ transformation matrix as well
            Dim tmpMatrix As Long
            GdipCreateMatrix2 1&, 0&, 0&, 1&, 0&, 0&, tmpMatrix
            
            'Attempt to render the blur effect
            Dim GDIPlusDebug As Long
            GDIPlusDebug = GdipDrawImageFX(iGraphics, tBitmap, tmpRect, tmpMatrix, hEffect, 0&, UnitPixel)
            If GDIPlusDebug > 0 Then Message "GDI+ failed to render blur effect (Error Code %1).", GDIPlusDebug
            
            'Delete our temporary transformation matrix
            GdipDeleteMatrix tmpMatrix
            
        Else
            Message "GDI+ failed to set effect parameters."
        End If
    
        'Delete our GDI+ blur object
        GdipDeleteEffect hEffect
    
    Else
        Message "GDI+ failed to create blur effect object"
    End If
        
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics iGraphics
    GdipDisposeImage tBitmap
    
End Function

'Use GDI+ to render a series of white-black-white circles, which are preferable for on-canvas controls with good readability
Public Function GDIPlusDrawCanvasCircle(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, Optional ByVal cTransparency As Long = 255) As Boolean

    GDIPlusDrawCircleToDC dstDC, cx, cy, cRadius, RGB(0, 0, 0), cTransparency, 3, True
    GDIPlusDrawCircleToDC dstDC, cx, cy, cRadius, RGB(255, 255, 255), 220, 1, True
    
End Function

'Retrieve a persistent handle to a GDI+-format graphics container.  Optionally, a smoothing mode can be specified so that it does
' not have to be repeatedly specified by a caller function.  (GDI+ sets smoothing mode by graphics container, not by function call.)
Public Function getGDIPlusImageHandleFromDC(ByVal srcDC As Long, Optional ByVal useAA As Boolean = True) As Long

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC srcDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone

    getGDIPlusImageHandleFromDC = iGraphics

End Function

Public Sub releaseGDIPlusImageHandle(ByVal srcHandle As Long)
    GdipDeleteGraphics srcHandle
End Sub

'Return a persistent handle to a GDI+ pen.  This can be useful if many drawing operations are going to be applied with the same pen.
Public Function getGDIPlusPenHandle(ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal customLinecap As LineCap = 0) As Long

    'Create the requested pen
    Dim iPen As Long
    GdipCreatePen1 fillQuadWithVBRGB(eColor, cTransparency), lineWidth, UnitPixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLinecap > 0 Then GdipSetPenLineCap iPen, customLinecap, customLinecap, 0&
    
    'Return the handle
    getGDIPlusPenHandle = iPen

End Function

Public Sub releaseGDIPlusPen(ByVal srcPen As Long)
    GdipDeletePen srcPen
End Sub

'Assuming the client has already obtained a GDI+ graphics handle and a GDI+ pen handle, they can use this function to quickly draw a line using
' the associated objects.
Public Sub GDIPlusDrawLine_Fast(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    'This function is just a thin wrapper to the GdipDrawLine function!
    GdipDrawLine dstGraphics, srcPen, x1, y1, x2, y2

End Sub

'Use GDI+ to render a line, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawLineToDC(ByVal dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLinecap As LineCap = 0) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone
    
    'Create a pen, which will be used to stroke the line
    Dim iPen As Long
    GdipCreatePen1 fillQuadWithVBRGB(eColor, cTransparency), lineWidth, UnitPixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLinecap > 0 Then GdipSetPenLineCap iPen, customLinecap, customLinecap, 0&
    
    'Render the line
    GdipDrawLine iGraphics, iPen, x1, y1, x2, y2
        
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a hollow rectangle, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawRectOutlineToDC(ByVal dstDC As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLinecap As LineCap = 0) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone
    
    'Create a pen, which will be used to stroke the line
    Dim iPen As Long
    GdipCreatePen1 fillQuadWithVBRGB(eColor, cTransparency), lineWidth, UnitPixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLinecap > 0 Then GdipSetPenLineCap iPen, customLinecap, customLinecap, 0&
    
    'Render the rectangle
    GdipDrawLine iGraphics, iPen, rectLeft, rectTop, rectRight, rectTop
    GdipDrawLine iGraphics, iPen, rectRight, rectTop, rectRight, rectBottom
    GdipDrawLine iGraphics, iPen, rectRight, rectBottom, rectLeft, rectBottom
    GdipDrawLine iGraphics, iPen, rectLeft, rectBottom, rectLeft, rectTop
        
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a hollow circle, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawCircleToDC(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal drawRadius As Single = 1, Optional ByVal useAA As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone
    
    'Create a pen, which will be used to stroke the circle
    Dim iPen As Long
    GdipCreatePen1 fillQuadWithVBRGB(eColor, cTransparency), drawRadius, UnitPixel, iPen
    
    'Render the circle
    GdipDrawEllipse iGraphics, iPen, cx - cRadius, cy - cRadius, cRadius * 2, cRadius * 2
        
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a filled ellipse, with optional antialiasing
Public Function GDIPlusDrawEllipseToDC(ByRef dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal useAA As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone
        
    'Create a solid fill brush
    Dim iBrush As Long
    GdipCreateSolidFill fillQuadWithVBRGB(eColor, 255), iBrush
    
    'Fill the ellipse
    GdipFillEllipseI iGraphics, iBrush, x1, y1, xWidth, yHeight
    
    'Release all created objects
    GdipDeleteBrush iBrush
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a rectangle with rounded corners, with optional antialiasing
Public Function GDIPlusDrawRoundRect(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal rRadius As Single, ByVal eColor As Long, Optional ByVal useAA As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, SmoothingModeAntiAlias Else GdipSetSmoothingMode iGraphics, SmoothingModeNone
    
    'GDI+ doesn't have a direct rounded rectangles call, so we have to do it ourselves with a custom path
    Dim rrPath As Long
    GdipCreatePath FillModeWinding, rrPath
        
    'The path will be rendered in two sections: first, filling it.  Second, stroking the path itself to complete the
    ' 1px outside border.
    xWidth = xWidth - 1
    yHeight = yHeight - 1
    
    GdipAddPathArc rrPath, x1 + xWidth - rRadius, y1, rRadius, rRadius, 270, 90
    GdipAddPathArc rrPath, x1 + xWidth - rRadius, y1 + yHeight - rRadius, rRadius, rRadius, 0, 90
    GdipAddPathArc rrPath, x1, y1 + yHeight - rRadius, rRadius, rRadius, 90, 90
    GdipAddPathArc rrPath, x1, y1, rRadius, rRadius, 180, 90
    GdipClosePathFigure rrPath
    
    'Create a solid fill brush
    Dim iBrush As Long
    GdipCreateSolidFill fillQuadWithVBRGB(eColor, 255), iBrush
    
    'Fill the path
    GdipFillPath iGraphics, iBrush, rrPath
    
    'Stroke the path as well (to fill the 1px exterior border)
    Dim iPen As Long
    GdipCreatePen1 fillQuadWithVBRGB(eColor, 255), 1, UnitPixel, iPen
    GdipDrawPath iGraphics, iPen, rrPath
    
    'Release all created objects
    GdipDeletePen iPen
    GdipDeletePath rrPath
    GdipDeleteBrush iBrush
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to fill a DIB with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full DIB alpha
' in a single pass.
Public Function GDIPlusFillDIBRect(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
    
    'Create a solid fill brush
    Dim iBrush As Long
    GdipCreateSolidFill fillQuadWithVBRGB(eColor, eTransparency), iBrush
    
    'Apply the brush
    GdipFillRectangleI iGraphics, iBrush, x1, y1, xWidth, yHeight
    
    'Release all created objects
    GdipDeleteBrush iBrush
    GdipDeleteGraphics iGraphics
    
    GDIPlusFillDIBRect = True

End Function

'GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from a standard RGB()
' long and supplied alpha.  It's not a very efficient conversion, but I need it so infrequently that I don't really care.
Private Function fillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = ExtractR(vbRGB)
    dstQuad.Green = ExtractG(vbRGB)
    dstQuad.Blue = ExtractB(vbRGB)
    dstQuad.Alpha = alphaValue
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    fillQuadWithVBRGB = placeHolder.lngResult
    
End Function

'Use GDI+ to load an image file.  Pretty bare-bones, but should be sufficient for any supported image type.
Public Function GDIPlusLoadPicture(ByVal srcFilename As String, ByRef dstDIB As pdDIB) As Boolean

    'Used to hold the return values of various GDI+ calls
    Dim GDIPlusReturn As Long
      
    'Use GDI+ to load the image
    Dim hImage As Long
    GDIPlusReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusLoadPicture = False
        Exit Function
    End If
    
    'Look for an ICC profile by asking GDI+ to return the ICC profile property's size
    Dim profileSize As Long, hasProfile As Boolean
    GdipGetPropertyItemSize hImage, PropertyTagICCProfile, profileSize
    
    'If the returned size is > 0, this image contains an ICC profile!  Retrieve it now.
    If profileSize > 0 Then
    
        hasProfile = True
    
        Dim iccProfileBuffer() As Byte
        ReDim iccProfileBuffer(0 To profileSize - 1) As Byte
        GdipGetPropertyItem hImage, PropertyTagICCProfile, profileSize, ByVal VarPtr(iccProfileBuffer(0))
        
        dstDIB.ICCProfile.loadICCFromGDIPlus profileSize - 16, VarPtr(iccProfileBuffer(0)) + 16
        
        Erase iccProfileBuffer
        
    End If
        
    'Retrieve the image's size
    Dim imgWidth As Single, imgHeight As Single
    GdipGetImageDimension hImage, imgWidth, imgHeight
    
    'Retrieve the image's horizontal and vertical resolution (if any)
    Dim imgHResolution As Single, imgVResolution As Single
    GdipGetImageHorizontalResolution hImage, imgHResolution
    GdipGetImageVerticalResolution hImage, imgVResolution
    dstDIB.setDPI imgHResolution, imgVResolution
    
    'Retrieve the image's alpha channel data (if any)
    Dim hasAlpha As Boolean
    hasAlpha = False
    
    Dim iPixelFormat As Long
    GdipGetImagePixelFormat hImage, iPixelFormat
    If (iPixelFormat And PixelFormatAlpha) <> 0 Then hasAlpha = True
    If (iPixelFormat And PixelFormatPremultipliedAlpha) <> 0 Then hasAlpha = True
    
    'Check for CMYK images
    Dim isCMYK As Boolean
    If (iPixelFormat = PixelFormat32bppCMYK) Then isCMYK = True
    
    'Create a blank DIB with matching size and alpha channel
    If hasAlpha Then
        dstDIB.createBlank CLng(imgWidth), CLng(imgHeight), 32
    Else
        dstDIB.createBlank CLng(imgWidth), CLng(imgHeight), 24
    End If
    
    Dim copyBitmapData As BitmapData
    Dim tmpRect As RECTL
    Dim iGraphics As Long
    
    'We now copy over image data in one of two ways.  If the image is 24bpp, our job is simple - use BitBlt and an hBitmap.
    ' 32bpp (including CMYK) images require a bit of extra work.
    If hasAlpha Then
        
        'Make sure the image is in 32bpp premultiplied ARGB format
        If iPixelFormat <> PixelFormat32bppPARGB Then GdipCloneBitmapAreaI 0, 0, imgWidth, imgHeight, PixelFormat32bppPARGB, hImage, hImage
        
        'We are now going to copy the image's data directly into our destination DIB by using LockBits.  Very fast, and not much code!
        
        'Start by preparing a BitmapData variable with instructions on where GDI+ should paste the bitmap data
        With copyBitmapData
            .Width = imgWidth
            .Height = imgHeight
            .PixelFormat = PixelFormat32bppPARGB
            .Stride = dstDIB.getDIBArrayWidth
            .Scan0 = dstDIB.getActualDIBBits
        End With
        
        'Next, prepare a clipping rect
        With tmpRect
            .Left = 0
            .Top = 0
            .Width = imgWidth
            .Height = imgHeight
        End With
        
        'Use LockBits to perform the copy for us.
        GdipBitmapLockBits hImage, tmpRect, ImageLockModeUserInputBuf Or ImageLockModeWrite Or ImageLockModeRead, PixelFormat32bppPARGB, copyBitmapData
        GdipBitmapUnlockBits hImage, copyBitmapData
    
    Else
    
        'CMYK is handled separately from regular RGB data, as we want to perform an ICC profile conversion as well.
        ' Note that if a CMYK profile is not present, we allow GDI+ to convert the image to RGB for us.
        If (isCMYK And hasProfile) Then
        
            'Create a blank 32bpp DIB, which will hold the CMYK data
            Dim tmpCMYKDIB As pdDIB
            Set tmpCMYKDIB = New pdDIB
            tmpCMYKDIB.createBlank imgWidth, imgHeight, 32
        
            'Next, prepare a BitmapData variable with instructions on where GDI+ should paste the bitmap data
            With copyBitmapData
                .Width = imgWidth
                .Height = imgHeight
                .PixelFormat = PixelFormat32bppCMYK
                .Stride = tmpCMYKDIB.getDIBArrayWidth
                .Scan0 = tmpCMYKDIB.getActualDIBBits
            End With
            
            'Next, prepare a clipping rect
            With tmpRect
                .Left = 0
                .Top = 0
                .Width = imgWidth
                .Height = imgHeight
            End With
            
            'Use LockBits to perform the copy for us.
            GdipBitmapLockBits hImage, tmpRect, ImageLockModeUserInputBuf Or ImageLockModeWrite Or ImageLockModeRead, PixelFormat32bppCMYK, copyBitmapData
            GdipBitmapUnlockBits hImage, copyBitmapData
                        
            'Apply the transformation using the dedicated CMYK transform handler
            If applyCMYKTransform(dstDIB.ICCProfile.getICCDataPointer, dstDIB.ICCProfile.getICCDataSize, tmpCMYKDIB, dstDIB, dstDIB.ICCProfile.getSourceRenderIntent) Then
            
                Message "Copying newly transformed sRGB data..."
            
                'The transform was successful, and the destination DIB is ready to go!
                dstDIB.ICCProfile.markSuccessfulProfileApplication
                
            'Something went horribly wrong.  Use GDI+ to apply a generic CMYK -> RGB transform.
            Else
            
                Message "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
            
                GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
                GdipDrawImageRect iGraphics, hImage, 0, 0, imgWidth, imgHeight
                GdipDeleteGraphics iGraphics
            
            End If
            
            Set tmpCMYKDIB = Nothing
        
        Else
            
            'Render the GDI+ image directly onto the newly created DIB
            GdipCreateFromHDC dstDIB.getDIBDC, iGraphics
            GdipDrawImageRect iGraphics, hImage, 0, 0, imgWidth, imgHeight
            GdipDeleteGraphics iGraphics
            
        End If
    
    End If
    
    'Release any remaining GDI+ handles and exit
    GdipDisposeImage hImage
    GDIPlusLoadPicture = True
    
End Function

'Save an image using GDI+.  Per the current save spec, ImageID must be specified.
' Additional save options are currently available for JPEGs (save quality, range [1,100]) and TIFFs (compression type).
Public Function GDIPlusSavePicture(ByRef srcPDImage As pdImage, ByVal dstFilename As String, ByVal imgFormat As GDIPlusImageFormat, ByVal outputColorDepth As Long, Optional ByVal JPEGQuality As Long = 92) As Boolean

    On Error GoTo GDIPlusSaveError

    Message "Initializing GDI+..."

    'If the output format is 24bpp (e.g. JPEG) but the input image is 32bpp, composite it against white
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.getCompositedImage tmpDIB, False
    If tmpDIB.getDIBColorDepth <> 24 And imgFormat = [ImageJPEG] Then tmpDIB.compositeBackgroundColor 255, 255, 255

    'Begin by creating a generic bitmap header for the current DIB
    Dim imgHeader As BITMAPINFO
    
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = tmpDIB.getDIBColorDepth
        .Width = tmpDIB.getDIBWidth
        .Height = -tmpDIB.getDIBHeight
    End With

    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hImage As Long
    
    Message "Creating GDI+ compatible image copy..."
        
    'Different GDI+ calls are required for different color depths. GdipCreateBitmapFromGdiDib leads to a blank
    ' alpha channel for 32bpp images, so use GdipCreateBitmapFromScan0 in that case.
    If tmpDIB.getDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GDIPlusReturn = GdipCreateBitmapFromScan0(tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpDIB.getDIBWidth * 4, PixelFormat32bppARGB, ByVal tmpDIB.getActualDIBBits, hImage)
    
    Else
        GDIPlusReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal tmpDIB.getActualDIBBits, hImage)
    End If
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    'Certain image formats require extra parameters, and because the values are passed ByRef, they can't be constants
    Dim GIF_EncoderVersion As Long
    GIF_EncoderVersion = [EncoderValueVersionGif89]
    
    Dim gdipColorDepth As Long
    gdipColorDepth = outputColorDepth
    
    Dim TIFF_Compression As Long
    TIFF_Compression = [EncoderValueCompressionLZW]
    
    'TIFF has some unique constraints on account of its many compression schemes.  Because it only supports a subset
    ' of compression types, we must adjust our code accordingly.
    If imgFormat = ImageTIFF Then
    
        Select Case g_UserPreferences.GetPref_Long("File Formats", "TIFF Compression", 0)
        
            'Default settings (LZW for > 1bpp, CCITT Group 4 fax encoding for 1bpp)
            Case 0
                If gdipColorDepth = 1 Then TIFF_Compression = [EncoderValueCompressionCCITT4] Else TIFF_Compression = [EncoderValueCompressionLZW]
                
            'No compression
            Case 1
                TIFF_Compression = [EncoderValueCompressionNone]
                
            'Macintosh Packbits (RLE)
            Case 2
                TIFF_Compression = [EncoderValueCompressionRle]
            
            'Proper deflate (Adobe-style) - not supported by GDI+
            Case 3
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'Obsolete deflate (PKZIP or zLib-style) - not supported by GDI+
            Case 4
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'LZW
            Case 5
                TIFF_Compression = [EncoderValueCompressionLZW]
                
            'JPEG - not supported by GDI+
            Case 6
                TIFF_Compression = [EncoderValueCompressionLZW]
            
            'Fax Group 3
            Case 7
                gdipColorDepth = 1
                TIFF_Compression = [EncoderValueCompressionCCITT3]
            
            'Fax Group 4
            Case 8
                gdipColorDepth = 1
                TIFF_Compression = [EncoderValueCompressionCCITT4]
                
        End Select
    
    End If
    
    'Request an encoder from GDI+ based on the type passed to this routine
    Dim uEncCLSID As CLSID
    Dim uEncParams As EncoderParameters
    Dim aEncParams() As Byte

    Message "Preparing GDI+ encoder for this filetype..."

    Select Case imgFormat
        
        'BMP export
        Case [ImageBMP]
            pvGetEncoderClsID "image/bmp", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
    
        'GIF export
        Case [ImageGIF]
            pvGetEncoderClsID "image/gif", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderVersion)
                .Value = VarPtr(GIF_EncoderVersion)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
            
        'JPEG export (requires extra work to specify a quality for the encode)
        Case [ImageJPEG]
            pvGetEncoderClsID "image/jpeg", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderQuality)
                .Value = VarPtr(JPEGQuality)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
        
        'PNG export
        Case [ImagePNG]
            pvGetEncoderClsID "image/png", uEncCLSID
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = EncoderParameterValueTypeLong
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
        
        'TIFF export (requires extra work to specify compression and color depth for the encode)
        Case [ImageTIFF]
            pvGetEncoderClsID "image/tiff", uEncCLSID
            uEncParams.Count = 2
            ReDim aEncParams(1 To Len(uEncParams) + Len(uEncParams.Parameter) * 2)
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderCompression)
                .Value = VarPtr(TIFF_Compression)
            End With
            
            CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
            
            With uEncParams.Parameter
                .NumberOfValues = 1
                .encType = [EncoderParameterValueTypeLong]
                .Guid = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(gdipColorDepth)
            End With
            
            CopyMemory aEncParams(Len(uEncParams) + 1), uEncParams.Parameter, Len(uEncParams.Parameter)
    
    End Select

    'With our encoder prepared, we can finally continue with the save
    
    'Check to see if a file already exists at this location
    If FileExist(dstFilename) Then Kill dstFilename
    
    Message "Saving the file..."
    
    'Perform the encode and save
    GDIPlusReturn = GdipSaveImageToFile(hImage, StrConv(dstFilename, vbUnicode), uEncCLSID, aEncParams(1))
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    Message "Releasing all temporary image copies..."
    
    'Release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    Message "Save complete."
    
    GDIPlusSavePicture = True
    Exit Function
    
GDIPlusSaveError:

    GDIPlusSavePicture = False
    
End Function

'Quickly export a DIB to PNG format using GDI+.
Public Function GDIPlusQuickSavePNG(ByVal dstFilename As String, ByRef srcDIB As pdDIB) As Boolean

    On Error GoTo GDIPlusQuickSaveError
    
    'Begin by creating a generic bitmap header for the current DIB
    Dim imgHeader As BITMAPINFO
    
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = srcDIB.getDIBColorDepth
        .Width = srcDIB.getDIBWidth
        .Height = -srcDIB.getDIBHeight
    End With

    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hImage As Long
        
    'Different GDI+ calls are required for different color depths. GdipCreateBitmapFromGdiDib leads to a blank
    ' alpha channel for 32bpp images, so use GdipCreateBitmapFromScan0 in that case.
    If srcDIB.getDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GDIPlusReturn = GdipCreateBitmapFromScan0(srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBWidth * 4, PixelFormat32bppARGB, ByVal srcDIB.getActualDIBBits, hImage)
    
    Else
        GDIPlusReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.getActualDIBBits, hImage)
    End If
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusQuickSavePNG = False
        Exit Function
    End If
        
    'Request a PNG encoder from GDI+
    Dim uEncCLSID As CLSID
    Dim uEncParams As EncoderParameters
    Dim aEncParams() As Byte
        
    pvGetEncoderClsID "image/png", uEncCLSID
    uEncParams.Count = 1
    ReDim aEncParams(1 To Len(uEncParams))
    
    Dim gdipColorDepth As Long
    gdipColorDepth = srcDIB.getDIBColorDepth
    
    With uEncParams.Parameter
        .NumberOfValues = 1
        .encType = EncoderParameterValueTypeLong
        .Guid = pvDEFINE_GUID(EncoderColorDepth)
        .Value = VarPtr(gdipColorDepth)
    End With
    
    CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
    
    'Check to see if a file already exists at this location
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Perform the encode and save
    GDIPlusReturn = GdipSaveImageToFile(hImage, StrConv(dstFilename, vbUnicode), uEncCLSID, aEncParams(1))
    
    If (GDIPlusReturn <> [OK]) Then
        GdipDisposeImage hImage
        GDIPlusQuickSavePNG = False
        Exit Function
    End If
    
    'Release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    GDIPlusQuickSavePNG = True
    Exit Function
    
GDIPlusQuickSaveError:

    GDIPlusQuickSavePNG = False
    
End Function

'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function isGDIPlusAvailable() As Boolean

    Dim gdiCheck As GdiplusStartupInput
    gdiCheck.GdiplusVersion = 1
    
    If (GdiplusStartup(g_GDIPlusToken, gdiCheck) <> [OK]) Then
        isGDIPlusAvailable = False
        g_GDIPlusAvailable = False
        g_GDIPlusFXAvailable = False
    Else
    
        isGDIPlusAvailable = True
        g_GDIPlusAvailable = True
        
        'Next, check to see if v1.1 is available.  This allows for advanced fx work.
        Dim hMod As Long
        hMod = LoadLibrary("gdiplus.dll")
        If hMod Then
            Dim testAddress As Long
            testAddress = GetProcAddress(hMod, "GdipDrawImageFX")
            If testAddress Then g_GDIPlusFXAvailable = True Else g_GDIPlusFXAvailable = False
            FreeLibrary hMod
        End If
        
    End If

End Function

'At shutdown, this function must be called to release our GDI+ instance
Public Function releaseGDIPlus()
    GdiplusShutdown g_GDIPlusToken
End Function

'Thanks to Carles P.V. for providing the following four functions, which are used as part of GDI+ image saving.
' You can download Carles's full project from http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
Private Function pvGetEncoderClsID(strMimeType As String, ClassID As CLSID) As Long

  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function pvDEFINE_GUID(ByVal sGuid As String) As CLSID
'-- Courtesy of: Dana Seaman
'   Helper routine to convert a CLSID(aka GUID) string to a structure
'   Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
    Call CLSIDFromString(StrPtr(sGuid), pvDEFINE_GUID)
End Function

'"Convert" (technically, dereference) an ANSI or Unicode string to the BSTR used by VB
Private Function pvPtrToStrW(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

'Same as above, but in reverse
'Private Function pvPtrToStrA(ByVal lpsz As Long) As String
'
'  Dim sOut As String
'  Dim lLen As Long
'
'    lLen = lstrlenA(lpsz)
'
'    If (lLen > 0) Then
'        sOut = String$(lLen, vbNullChar)
'        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
'        pvPtrToStrA = sOut
'    End If
'End Function

