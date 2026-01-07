Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright 2012-2026 by Tanner Helland
'Created: 1/September/12
'Last updated: 07/April/25
'Last update: add support for rendering from non-system (user-specified) font files
'
'This interface provides a means for interacting with various GDI+ features.  GDI+ was originally
' used as a fallback for image loading and saving if the FreeImage DLL was not found, but over time
' it has become more and more essential to PhotoDemon .  As of version 7.0, GDI+ is deeply embedded
' into PD's rendering pipeline, as it's currently the easiest+fastest way to reasmple 32-bpp pixel
' data regardless of underlying PC features.  It is also used extensively in rendering PD's custom UI.
'
'Jose Roca's convenient GDI+ reference has been a huge help with GDI+ development:
' http://www.jose.it-berater.org/gdiplus/iframe/index.htm
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum GP_Result
    GP_OK = 0
    GP_GenericError = 1
    GP_InvalidParameter = 2
    GP_OutOfMemory = 3
    GP_ObjectBusy = 4
    GP_InsufficientBuffer = 5
    GP_NotImplemented = 6
    GP_Win32Error = 7
    GP_WrongState = 8
    GP_Aborted = 9
    GP_FileNotFound = 10
    GP_ValueOverflow = 11
    GP_AccessDenied = 12
    GP_UnknownImageFormat = 13
    GP_FontFamilyNotFound = 14
    GP_FontStyleNotFound = 15
    GP_NotTrueTypeFont = 16
    GP_UnsupportedGDIPlusVersion = 17
    GP_GDIPlusNotInitialized = 18
    GP_PropertyNotFound = 19
    GP_PropertyNotSupported = 20
End Enum

#If False Then
    Private Const GP_OK = 0, GP_GenericError = 1, GP_InvalidParameter = 2, GP_OutOfMemory = 3, GP_ObjectBusy = 4, GP_InsufficientBuffer = 5, GP_NotImplemented = 6, GP_Win32Error = 7, GP_WrongState = 8, GP_Aborted = 9, GP_FileNotFound = 10, GP_ValueOverflow = 11, GP_AccessDenied = 12, GP_UnknownImageFormat = 13
    Private Const GP_FontFamilyNotFound = 14, GP_FontStyleNotFound = 15, GP_NotTrueTypeFont = 16, GP_UnsupportedGDIPlusVersion = 17, GP_GDIPlusNotInitialized = 18, GP_PropertyNotFound = 19, GP_PropertyNotSupported = 20
#End If

Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'Private Enum GP_DebugEventLevel
'    GP_DebugEventLevelFatal = 0
'    GP_DebugEventLevelWarning = 1
'End Enum
'
'#If False Then
'    Private Const GP_DebugEventLevelFatal = 0, GP_DebugEventLevelWarning = 1
'#End If

'Drawing-related enums

Public Enum GP_QualityMode      'Note that many other settings just wrap these default Quality Mode values
    GP_QM_Invalid = -1
    GP_QM_Default = 0
    GP_QM_Low = 1
    GP_QM_High = 2
End Enum

#If False Then
    Private Const GP_QM_Invalid = -1, GP_QM_Default = 0, GP_QM_Low = 1, GP_QM_High = 2
#End If

Public Enum GP_BitmapLockMode
    GP_BLM_Read = &H1
    GP_BLM_Write = &H2
    GP_BLM_UserInputBuf = &H4
End Enum

#If False Then
    Private Const GP_BLM_Read = &H1, GP_BLM_Write = &H2, GP_BLM_UserInputBuf = &H4
#End If

'Color adjustments are handled internally, at present, so we don't need to expose them to other objects
Private Enum GP_ColorAdjustType
    GP_CAT_Default = 0
    GP_CAT_Bitmap = 1
    GP_CAT_Brush = 2
    GP_CAT_Pen = 3
    GP_CAT_Text = 4
    GP_CAT_Count = 5
    GP_CAT_Any = 6
End Enum

#If False Then
    Private Const GP_CAT_Default = 0, GP_CAT_Bitmap = 1, GP_CAT_Brush = 2, GP_CAT_Pen = 3, GP_CAT_Text = 4, GP_CAT_Count = 5, GP_CAT_Any = 6
#End If

Private Enum GP_ColorMatrixFlags
    GP_CMF_Default = 0
    GP_CMF_SkipGrays = 1
    GP_CMF_AltGray = 2
End Enum

#If False Then
    Private Const GP_CMF_Default = 0, GP_CMF_SkipGrays = 1, GP_CMF_AltGray = 2
#End If

Public Enum GP_CombineMode
    GP_CM_Replace = 0
    GP_CM_Intersect = 1
    GP_CM_Union = 2
    GP_CM_Xor = 3
    GP_CM_Exclude = 4
    GP_CM_Complement = 5
End Enum

#If False Then
    Private Const GP_CM_Replace = 0, GP_CM_Intersect = 1, GP_CM_Union = 2, GP_CM_Xor = 3, GP_CM_Exclude = 4, GP_CM_Complement = 5
#End If

'Compositing mode is the closest GDI+ comes to offering "blend modes".  The default mode alpha-blends the source
' with the destination; "copy" mode overwrites the destination completely.
Public Enum GP_CompositingMode
    GP_CM_SourceOver = 0
    GP_CM_SourceCopy = 1
End Enum

#If False Then
    Private Const GP_CM_SourceOver = 0, GP_CM_SourceCopy = 1
#End If

'Alpha compositing qualities, which affects how GDI+ blends pixels.  Use with caution, as gamma-corrected blending
' yields non-inutitive results!
Public Enum GP_CompositingQuality
    GP_CQ_Invalid = GP_QM_Invalid
    GP_CQ_Default = GP_QM_Default
    GP_CQ_HighSpeed = GP_QM_Low
    GP_CQ_HighQuality = GP_QM_High
    GP_CQ_GammaCorrected = 3&
    GP_CQ_AssumeLinear = 4&
End Enum

#If False Then
    Private Const GP_CQ_Invalid = GP_QM_Invalid, GP_CQ_Default = GP_QM_Default, GP_CQ_HighSpeed = GP_QM_Low, GP_CQ_HighQuality = GP_QM_High, GP_CQ_GammaCorrected = 3&, GP_CQ_AssumeLinear = 4&
#End If

Public Enum GP_DashCap
    GP_DC_Flat = 0
    GP_DC_Square = 0     'This is not a typo; it's supplied as a convenience enum to match supported GP_LineCap values (which differentiate between flat and square, as they should)
    GP_DC_Round = 2
    GP_DC_Triangle = 3
End Enum

#If False Then
    Private Const GP_DC_Flat = 0, GP_DC_Square = 0, GP_DC_Round = 2, GP_DC_Triangle = 3
#End If

Public Enum GP_DashStyle
    GP_DS_Solid = 0&
    GP_DS_Dash = 1&
    GP_DS_Dot = 2&
    GP_DS_DashDot = 3&
    GP_DS_DashDotDot = 4&
    GP_DS_Custom = 5&
End Enum

#If False Then
    Private Const GP_DS_Solid = 0&, GP_DS_Dash = 1&, GP_DS_Dot = 2&, GP_DS_DashDot = 3&, GP_DS_DashDotDot = 4&, GP_DS_Custom = 5&
#End If

Public Enum GP_EncoderValueType
    GP_EVT_Byte = 1
    GP_EVT_ASCII = 2
    GP_EVT_Short = 3
    GP_EVT_Long = 4
    GP_EVT_Rational = 5
    GP_EVT_LongRange = 6
    GP_EVT_Undefined = 7
    GP_EVT_RationalRange = 8
    GP_EVT_Pointer = 9
End Enum

#If False Then
    Private Const GP_EVT_Byte = 1, GP_EVT_ASCII = 2, GP_EVT_Short = 3, GP_EVT_Long = 4, GP_EVT_Rational = 5, GP_EVT_LongRange = 6, GP_EVT_Undefined = 7, GP_EVT_RationalRange = 8, GP_EVT_Pointer = 9
#End If

Public Enum GP_EncoderValue
    GP_EV_ColorTypeCMYK = 0
    GP_EV_ColorTypeYCCK = 1
    GP_EV_CompressionLZW = 2
    GP_EV_CompressionCCITT3 = 3
    GP_EV_CompressionCCITT4 = 4
    GP_EV_CompressionRle = 5
    GP_EV_CompressionNone = 6
    GP_EV_ScanMethodInterlaced = 7
    GP_EV_ScanMethodNonInterlaced = 8
    GP_EV_VersionGif87 = 9
    GP_EV_VersionGif89 = 10
    GP_EV_RenderProgressive = 11
    GP_EV_RenderNonProgressive = 12
    GP_EV_TransformRotate90 = 13
    GP_EV_TransformRotate180 = 14
    GP_EV_TransformRotate270 = 15
    GP_EV_TransformFlipHorizontal = 16
    GP_EV_TransformFlipVertical = 17
    GP_EV_MultiFrame = 18
    GP_EV_LastFrame = 19
    GP_EV_Flush = 20
    GP_EV_FrameDimensionTime = 21
    GP_EV_FrameDimensionResolution = 22
    GP_EV_FrameDimensionPage = 23
    GP_EV_ColorTypeGray = 24
    GP_EV_ColorTypeRGB = 25
End Enum

#If False Then
    Private Const GP_EV_ColorTypeCMYK = 0, GP_EV_ColorTypeYCCK = 1, GP_EV_CompressionLZW = 2, GP_EV_CompressionCCITT3 = 3, GP_EV_CompressionCCITT4 = 4, GP_EV_CompressionRle = 5, GP_EV_CompressionNone = 6, GP_EV_ScanMethodInterlaced = 7, GP_EV_ScanMethodNonInterlaced = 8, GP_EV_VersionGif87 = 9, GP_EV_VersionGif89 = 10
    Private Const GP_EV_RenderProgressive = 11, GP_EV_RenderNonProgressive = 12, GP_EV_TransformRotate90 = 13, GP_EV_TransformRotate180 = 14, GP_EV_TransformRotate270 = 15, GP_EV_TransformFlipHorizontal = 16, GP_EV_TransformFlipVertical = 17, GP_EV_MultiFrame = 18, GP_EV_LastFrame = 19, GP_EV_Flush = 20
    Private Const GP_EV_FrameDimensionTime = 21, GP_EV_FrameDimensionResolution = 22, GP_EV_FrameDimensionPage = 23, GP_EV_ColorTypeGray = 24, GP_EV_ColorTypeRGB = 25
#End If

Public Enum GP_FillMode
    GP_FM_Alternate = 0&
    GP_FM_Winding = 1&
End Enum

#If False Then
    Private Const GP_FM_Alternate = 0&, GP_FM_Winding = 1&
#End If

Public Enum GP_ImageType
    GP_IT_Unknown = 0
    GP_IT_Bitmap = 1
    GP_IT_Metafile = 2
End Enum

#If False Then
    Private Const GP_IT_Unknown = 0, GP_IT_Bitmap = 1, GP_IT_Metafile = 2
#End If

Public Enum GP_InterpolationMode
    GP_IM_Invalid = GP_QM_Invalid
    GP_IM_Default = GP_QM_Default
    GP_IM_LowQuality = GP_QM_Low
    GP_IM_HighQuality = GP_QM_High
    GP_IM_Bilinear = 3
    GP_IM_Bicubic = 4
    GP_IM_NearestNeighbor = 5
    GP_IM_HighQualityBilinear = 6
    GP_IM_HighQualityBicubic = 7
End Enum

#If False Then
    Private Const GP_IM_Invalid = GP_QM_Invalid, GP_IM_Default = GP_QM_Default, GP_IM_LowQuality = GP_QM_Low, GP_IM_HighQuality = GP_QM_High, GP_IM_Bilinear = 3, GP_IM_Bicubic = 4, GP_IM_NearestNeighbor = 5, GP_IM_HighQualityBilinear = 6, GP_IM_HighQualityBicubic = 7
#End If

Public Enum GP_LineCap
    GP_LC_Flat = 0&
    GP_LC_Square = 1&
    GP_LC_Round = 2&
    GP_LC_Triangle = 3&
    GP_LC_NoAnchor = &H10
    GP_LC_SquareAnchor = &H11
    GP_LC_RoundAnchor = &H12
    GP_LC_DiamondAnchor = &H13
    GP_LC_ArrowAnchor = &H14
    GP_LC_Custom = &HFF
End Enum

#If False Then
    Private Const GP_LC_Flat = 0, GP_LC_Square = 1, GP_LC_Round = 2, GP_LC_Triangle = 3, GP_LC_NoAnchor = &H10, GP_LC_SquareAnchor = &H11, GP_LC_RoundAnchor = &H12, GP_LC_DiamondAnchor = &H13, GP_LC_ArrowAnchor = &H14, GP_LC_Custom = &HFF
#End If

Public Enum GP_LineJoin
    GP_LJ_Miter = 0&
    GP_LJ_Bevel = 1&
    GP_LJ_Round = 2&
End Enum

#If False Then
    Private Const GP_LJ_Miter = 0&, GP_LJ_Bevel = 1&, GP_LJ_Round = 2&
#End If

Public Enum GP_MatrixOrder
    GP_MO_Prepend = 0&
    GP_MO_Append = 1&
End Enum

#If False Then
    Private Const GP_MO_Prepend = 0&, GP_MO_Append = 1&
#End If

'EMFs can be converted between various formats.  GDI+ prefers "EMF+", which supports GDI+ primitives as well
Public Enum GP_MetafileType
    GP_MT_Invalid = 0
    GP_MT_Wmf = 1
    GP_MT_WmfPlaceable = 2
    GP_MT_Emf = 3              'Old-style EMF consisting only of GDI commands
    GP_MT_EmfPlus = 4          'New-style EMF+ consisting only of GDI+ commands
    GP_MT_EmfDual = 5          'New-style EMF+ with GDI fallbacks for legacy rendering
End Enum

#If False Then
    Private Const GP_MT_Invalid = 0, GP_MT_Wmf = 1, GP_MT_WmfPlaceable = 2, GP_MT_Emf = 3, GP_MT_EmfPlus = 4, GP_MT_EmfDual = 5
#End If

Public Enum GP_PatternStyle
    GP_PS_Horizontal = 0
    GP_PS_Vertical = 1
    GP_PS_ForwardDiagonal = 2
    GP_PS_BackwardDiagonal = 3
    GP_PS_Cross = 4
    GP_PS_DiagonalCross = 5
    GP_PS_05Percent = 6
    GP_PS_10Percent = 7
    GP_PS_20Percent = 8
    GP_PS_25Percent = 9
    GP_PS_30Percent = 10
    GP_PS_40Percent = 11
    GP_PS_50Percent = 12
    GP_PS_60Percent = 13
    GP_PS_70Percent = 14
    GP_PS_75Percent = 15
    GP_PS_80Percent = 16
    GP_PS_90Percent = 17
    GP_PS_LightDownwardDiagonal = 18
    GP_PS_LightUpwardDiagonal = 19
    GP_PS_DarkDownwardDiagonal = 20
    GP_PS_DarkUpwardDiagonal = 21
    GP_PS_WideDownwardDiagonal = 22
    GP_PS_WideUpwardDiagonal = 23
    GP_PS_LightVertical = 24
    GP_PS_LightHorizontal = 25
    GP_PS_NarrowVertical = 26
    GP_PS_NarrowHorizontal = 27
    GP_PS_DarkVertical = 28
    GP_PS_DarkHorizontal = 29
    GP_PS_DashedDownwardDiagonal = 30
    GP_PS_DashedUpwardDiagonal = 31
    GP_PS_DashedHorizontal = 32
    GP_PS_DashedVertical = 33
    GP_PS_SmallConfetti = 34
    GP_PS_LargeConfetti = 35
    GP_PS_ZigZag = 36
    GP_PS_Wave = 37
    GP_PS_DiagonalBrick = 38
    GP_PS_HorizontalBrick = 39
    GP_PS_Weave = 40
    GP_PS_Plaid = 41
    GP_PS_Divot = 42
    GP_PS_DottedGrid = 43
    GP_PS_DottedDiamond = 44
    GP_PS_Shingle = 45
    GP_PS_Trellis = 46
    GP_PS_Sphere = 47
    GP_PS_SmallGrid = 48
    GP_PS_SmallCheckerBoard = 49
    GP_PS_LargeCheckerBoard = 50
    GP_PS_OutlinedDiamond = 51
    GP_PS_SolidDiamond = 52
End Enum

#If False Then
    Private Const GP_PS_Horizontal = 0, GP_PS_Vertical = 1, GP_PS_ForwardDiagonal = 2, GP_PS_BackwardDiagonal = 3, GP_PS_Cross = 4, GP_PS_DiagonalCross = 5, GP_PS_05Percent = 6, GP_PS_10Percent = 7, GP_PS_20Percent = 8, GP_PS_25Percent = 9, GP_PS_30Percent = 10, GP_PS_40Percent = 11, GP_PS_50Percent = 12, GP_PS_60Percent = 13, GP_PS_70Percent = 14, GP_PS_75Percent = 15, GP_PS_80Percent = 16, GP_PS_90Percent = 17, GP_PS_LightDownwardDiagonal = 18, GP_PS_LightUpwardDiagonal = 19, GP_PS_DarkDownwardDiagonal = 20, GP_PS_DarkUpwardDiagonal = 21, GP_PS_WideDownwardDiagonal = 22, GP_PS_WideUpwardDiagonal = 23, GP_PS_LightVertical = 24, GP_PS_LightHorizontal = 25
    Private Const GP_PS_NarrowVertical = 26, GP_PS_NarrowHorizontal = 27, GP_PS_DarkVertical = 28, GP_PS_DarkHorizontal = 29, GP_PS_DashedDownwardDiagonal = 30, GP_PS_DashedUpwardDiagonal = 31, GP_PS_DashedHorizontal = 32, GP_PS_DashedVertical = 33, GP_PS_SmallConfetti = 34, GP_PS_LargeConfetti = 35, GP_PS_ZigZag = 36, GP_PS_Wave = 37, GP_PS_DiagonalBrick = 38, GP_PS_HorizontalBrick = 39, GP_PS_Weave = 40, GP_PS_Plaid = 41, GP_PS_Divot = 42, GP_PS_DottedGrid = 43, GP_PS_DottedDiamond = 44, GP_PS_Shingle = 45, GP_PS_Trellis = 46, GP_PS_Sphere = 47, GP_PS_SmallGrid = 48, GP_PS_SmallCheckerBoard = 49, GP_PS_LargeCheckerBoard = 50
    Private Const GP_PS_OutlinedDiamond = 51, GP_PS_SolidDiamond = 52
#End If

'GDI+ pixel format IDs use a bitfield system:
' [0, 7] = format index
' [8, 15] = pixel size (in bits)
' [16, 23] = flags
' [24, 31] = reserved (currently unused)

'Note also that pixel format is *not* 100% reliable.  Behavior differs between OSes, even for the "same"
' major GDI+ version.  (See http://stackoverflow.com/questions/5065371/how-to-identify-cmyk-images-in-asp-net-using-c-sharp)
Public Enum GP_PixelFormat
    GP_PF_Indexed = &H10000         'Image uses a palette to define colors
    GP_PF_GDI = &H20000             'Is a format supported by GDI
    GP_PF_Alpha = &H40000           'Alpha channel present
    GP_PF_PreMultAlpha = &H80000    'Alpha is premultiplied (not always correct; manual verification should be used)
    GP_PF_HDR = &H100000            'High bit-depth colors are in use (e.g. 48-bpp or 64-bpp; behavior is unpredictable on XP)
    GP_PF_Canonical = &H200000      'Canonical formats: 32bppARGB, 32bppPARGB, 64bppARGB, 64bppPARGB
    
    GP_PF_32bppCMYK = &H200F        'CMYK is never returned on XP or Vista; ImageFlags can be checked as a failsafe
                                    ' (Conversely, ImageFlags are unreliable on Win 7 - this is the shit we deal with
                                    '  as Windows developers!)
    
    GP_PF_1bppIndexed = &H30101
    GP_PF_4bppIndexed = &H30402
    GP_PF_8bppIndexed = &H30803
    GP_PF_16bppGreyscale = &H101004
    GP_PF_16bppRGB555 = &H21005
    GP_PF_16bppRGB565 = &H21006
    GP_PF_16bppARGB1555 = &H61007
    GP_PF_24bppRGB = &H21808
    GP_PF_32bppRGB = &H22009
    GP_PF_32bppARGB = &H26200A
    GP_PF_32bppPARGB = &HE200B
    GP_PF_48bppRGB = &H10300C
    GP_PF_64bppARGB = &H34400D
    GP_PF_64bppPARGB = &H1C400E
End Enum

#If False Then
    Private Const GP_PF_Indexed = &H10000, GP_PF_GDI = &H20000, GP_PF_Alpha = &H40000, GP_PF_PreMultAlpha = &H80000, GP_PF_HDR = &H100000, GP_PF_Canonical = &H200000, GP_PF_32bppCMYK = &H200F
    Private Const GP_PF_1bppIndexed = &H30101, GP_PF_4bppIndexed = &H30402, GP_PF_8bppIndexed = &H30803, GP_PF_16bppGreyscale = &H101004, GP_PF_16bppRGB555 = &H21005, GP_PF_16bppRGB565 = &H21006
    Private Const GP_PF_16bppARGB1555 = &H61007, GP_PF_24bppRGB = &H21808, GP_PF_32bppRGB = &H22009, GP_PF_32bppARGB = &H26200A, GP_PF_32bppPARGB = &HE200B, GP_PF_48bppRGB = &H10300C, GP_PF_64bppARGB = &H34400D, GP_PF_64bppPARGB = &H1C400E
#End If

'PixelOffsetMode controls how GDI+ calculates positioning.  Normally, each a pixel is treated as a unit square that covers
' the area between [0, 0] and [1, 1].  However, for point-based objects like paths, GDI+ can treat coordinates as if they
' are centered over [0.5, 0.5] offsets within each pixel.  This typically yields prettier path renders, at some consequence
' to rendering performance.  (See http://drilian.com/2008/11/25/understanding-half-pixel-and-half-texel-offsets/)
Public Enum GP_PixelOffsetMode
    GP_POM_Invalid = GP_QM_Invalid
    GP_POM_Default = GP_QM_Default
    GP_POM_HighSpeed = GP_QM_Low
    GP_POM_HighQuality = GP_QM_High
    GP_POM_None = 3&
    GP_POM_Half = 4&
End Enum

#If False Then
    Private Const GP_POM_Invalid = QualityModeInvalid, GP_POM_Default = QualityModeDefault, GP_POM_HighSpeed = QualityModeLow, GP_POM_HighQuality = QualityModeHigh, GP_POM_None = 3, GP_POM_Half = 4
#End If

'Property tags describe image metadata.  Metadata is very complicated to read and/or write, because tags are encoded
' in a variety of ways.  Refer to https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx for details.
' pd2D uses these sparingly; do not expect it to perform full metadata preservation.
Public Enum GP_PropertyTag
'    GP_PT_Artist = &H13B&
'    GP_PT_BitsPerSample = &H102&
'    GP_PT_CellHeight = &H109&
'    GP_PT_CellWidth = &H108&
'    GP_PT_ChrominanceTable = &H5091&
'    GP_PT_ColorMap = &H140&
'    GP_PT_ColorTransferFunction = &H501A&
'    GP_PT_Compression = &H103&
'    GP_PT_Copyright = &H8298&
'    GP_PT_DateTime = &H132&
'    GP_PT_DocumentName = &H10D&
'    GP_PT_DotRange = &H150&
'    GP_PT_EquipMake = &H10F&
'    GP_PT_EquipModel = &H110&
'    GP_PT_ExifAperture = &H9202&
'    GP_PT_ExifBrightness = &H9203&
'    GP_PT_ExifCfaPattern = &HA302&
'    GP_PT_ExifColorSpace = &HA001&
'    GP_PT_ExifCompBPP = &H9102&
'    GP_PT_ExifCompConfig = &H9101&
'    GP_PT_ExifDTDigitized = &H9004&
'    GP_PT_ExifDTDigSS = &H9292&
'    GP_PT_ExifDTOrig = &H9003&
'    GP_PT_ExifDTOrigSS = &H9291&
'    GP_PT_ExifDTSubsec = &H9290&
'    GP_PT_ExifExposureBias = &H9204&
'    GP_PT_ExifExposureIndex = &HA215&
'    GP_PT_ExifExposureProg = &H8822&
'    GP_PT_ExifExposureTime = &H829A&
'    GP_PT_ExifFileSource = &HA300&
'    GP_PT_ExifFlash = &H9209&
'    GP_PT_ExifFlashEnergy = &HA20B&
'    GP_PT_ExifFNumber = &H829D&
'    GP_PT_ExifFocalLength = &H920A&
'    GP_PT_ExifFocalResUnit = &HA210&
'    GP_PT_ExifFocalXRes = &HA20E&
'    GP_PT_ExifFocalYRes = &HA20F&
'    GP_PT_ExifFPXVer = &HA000&
'    GP_PT_ExifIFD = &H8769&
'    GP_PT_ExifInterop = &HA005&
'    GP_PT_ExifISOSpeed = &H8827&
'    GP_PT_ExifLightSource = &H9208&
'    GP_PT_ExifMakerNote = &H927C&
'    GP_PT_ExifMaxAperture = &H9205&
'    GP_PT_ExifMeteringMode = &H9207&
'    GP_PT_ExifOECF = &H8828&
'    GP_PT_ExifPixXDim = &HA002&
'    GP_PT_ExifPixYDim = &HA003&
'    GP_PT_ExifRelatedWav = &HA004&
'    GP_PT_ExifSceneType = &HA301&
'    GP_PT_ExifSensingMethod = &HA217&
'    GP_PT_ExifShutterSpeed = &H9201&
'    GP_PT_ExifSpatialFR = &HA20C&
'    GP_PT_ExifSpectralSense = &H8824&
'    GP_PT_ExifSubjectDist = &H9206&
'    GP_PT_ExifSubjectLoc = &HA214&
'    GP_PT_ExifUserComment = &H9286&
'    GP_PT_ExifVer = &H9000&
'    GP_PT_ExtraSamples = &H152&
'    GP_PT_FillOrder = &H10A&
    GP_PT_FrameDelay = &H5100&
'    GP_PT_FreeByteCounts = &H121&
'    GP_PT_FreeOffset = &H120&
'    GP_PT_Gamma = &H301&
'    GP_PT_GlobalPalette = &H5102&
'    GP_PT_GpsAltitude = &H6&
'    GP_PT_GpsAltitudeRef = &H5&
'    GP_PT_GpsDestBear = &H18&
'    GP_PT_GpsDestBearRef = &H17&
'    GP_PT_GpsDestDist = &H1A&
'    GP_PT_GpsDestDistRef = &H19&
'    GP_PT_GpsDestLat = &H14&
'    GP_PT_GpsDestLatRef = &H13&
'    GP_PT_GpsDestLong = &H16&
'    GP_PT_GpsDestLongRef = &H15&
'    GP_PT_GpsGpsDop = &HB&
'    GP_PT_GpsGpsMeasureMode = &HA&
'    GP_PT_GpsGpsSatellites = &H8&
'    GP_PT_GpsGpsStatus = &H9&
'    GP_PT_GpsGpsTime = &H7&
'    GP_PT_GpsIFD = &H8825&
'    GP_PT_GpsImgDir = &H11&
'    GP_PT_GpsImgDirRef = &H10&
'    GP_PT_GpsLatitude = &H2&
'    GP_PT_GpsLatitudeRef = &H1&
'    GP_PT_GpsLongitude = &H4&
'    GP_PT_GpsLongitudeRef = &H3&
'    GP_PT_GpsMapDatum = &H12&
'    GP_PT_GpsSpeed = &HD&
'    GP_PT_GpsSpeedRef = &HC&
'    GP_PT_GpsTrack = &HF&
'    GP_PT_GpsTrackRef = &HE&
'    GP_PT_GpsVer = &H0&
'    GP_PT_GrayResponseCurve = &H123&
'    GP_PT_GrayResponseUnit = &H122&
'    GP_PT_GridSize = &H5011&
'    GP_PT_HalftoneDegree = &H500C&
'    GP_PT_HalftoneHints = &H141&
'    GP_PT_HalftoneLPI = &H500A&
'    GP_PT_HalftoneLPIUnit = &H500B&
'    GP_PT_HalftoneMisc = &H500E&
'    GP_PT_HalftoneScreen = &H500F&
'    GP_PT_HalftoneShape = &H500D&
'    GP_PT_HostComputer = &H13C&
    GP_PT_ICCProfile = &H8773&
    GP_PT_ICCProfileDescriptor = &H302&
'    GP_PT_ImageDescription = &H10E&
'    GP_PT_ImageHeight = &H101&
'    GP_PT_ImageTitle = &H320&
'    GP_PT_ImageWidth = &H100&
'    GP_PT_IndexBackground = &H5103&
'    GP_PT_IndexTransparent = &H5104&
'    GP_PT_InkNames = &H14D&
'    GP_PT_InkSet = &H14C&
'    GP_PT_JPEGACTables = &H209&
'    GP_PT_JPEGDCTables = &H208&
'    GP_PT_JPEGInterFormat = &H201&
'    GP_PT_JPEGInterLength = &H202&
'    GP_PT_JPEGLosslessPredictors = &H205&
'    GP_PT_JPEGPointTransforms = &H206&
'    GP_PT_JPEGProc = &H200&
'    GP_PT_JPEGQTables = &H207&
'    GP_PT_JPEGQuality = &H5010&
'    GP_PT_JPEGRestartInterval = &H203&
    GP_PT_LoopCount = &H5101&
'    GP_PT_LuminanceTable = &H5090&
'    GP_PT_MaxSampleValue = &H119&
'    GP_PT_MinSampleValue = &H118&
'    GP_PT_NewSubfileType = &HFE&
'    GP_PT_NumberOfInks = &H14E&
    GP_PT_Orientation = &H112&
    GP_PT_PageName = &H11D&
    GP_PT_PageNumber = &H129&
'    GP_PT_PaletteHistogram = &H5113&
'    GP_PT_PhotometricInterp = &H106&
'    GP_PT_PixelPerUnitX = &H5111&
'    GP_PT_PixelPerUnitY = &H5112&
'    GP_PT_PixelUnit = &H5110&
'    GP_PT_PlanarConfig = &H11C&
'    GP_PT_Predictor = &H13D&
'    GP_PT_PrimaryChromaticities = &H13F&
'    GP_PT_PrintFlags = &H5005&
'    GP_PT_PrintFlagsBleedWidth = &H5008&
'    GP_PT_PrintFlagsBleedWidthScale = &H5009&
'    GP_PT_PrintFlagsCrop = &H5007&
'    GP_PT_PrintFlagsVersion = &H5006&
'    GP_PT_REFBlackWhite = &H214&
'    GP_PT_ResolutionUnit = &H128&
'    GP_PT_ResolutionXLengthUnit = &H5003&
'    GP_PT_ResolutionXUnit = &H5001&
'    GP_PT_ResolutionYLengthUnit = &H5004&
'    GP_PT_ResolutionYUnit = &H5002&
'    GP_PT_RowsPerStrip = &H116&
'    GP_PT_SampleFormat = &H153&
'    GP_PT_SamplesPerPixel = &H115&
'    GP_PT_SMaxSampleValue = &H155&
'    GP_PT_SMinSampleValue = &H154&
'    GP_PT_SoftwareUsed = &H131&
'    GP_PT_SRGBRenderingIntent = &H303&
'    GP_PT_StripBytesCount = &H117&
'    GP_PT_StripOffsets = &H111&
'    GP_PT_SubfileType = &HFF&
'    GP_PT_T4Option = &H124&
'    GP_PT_T6Option = &H125&
'    GP_PT_TargetPrinter = &H151&
'    GP_PT_ThreshHolding = &H107&
'    GP_PT_ThumbnailArtist = &H5034&
'    GP_PT_ThumbnailBitsPerSample = &H5022&
'    GP_PT_ThumbnailColorDepth = &H5015&
'    GP_PT_ThumbnailCompressedSize = &H5019&
'    GP_PT_ThumbnailCompression = &H5023&
'    GP_PT_ThumbnailCopyRight = &H503B&
'    GP_PT_ThumbnailData = &H501B&
'    GP_PT_ThumbnailDateTime = &H5033&
'    GP_PT_ThumbnailEquipMake = &H5026&
'    GP_PT_ThumbnailEquipModel = &H5027&
'    GP_PT_ThumbnailFormat = &H5012&
'    GP_PT_ThumbnailHeight = &H5014&
'    GP_PT_ThumbnailImageDescription = &H5025&
'    GP_PT_ThumbnailImageHeight = &H5021&
'    GP_PT_ThumbnailImageWidth = &H5020&
'    GP_PT_ThumbnailOrientation = &H5029&
'    GP_PT_ThumbnailPhotometricInterp = &H5024&
'    GP_PT_ThumbnailPlanarConfig = &H502F&
'    GP_PT_ThumbnailPlanes = &H5016&
'    GP_PT_ThumbnailPrimaryChromaticities = &H5036&
'    GP_PT_ThumbnailRawBytes = &H5017&
'    GP_PT_ThumbnailRefBlackWhite = &H503A&
'    GP_PT_ThumbnailResolutionUnit = &H5030&
'    GP_PT_ThumbnailResolutionX = &H502D&
'    GP_PT_ThumbnailResolutionY = &H502E&
'    GP_PT_ThumbnailRowsPerStrip = &H502B&
'    GP_PT_ThumbnailSamplesPerPixel = &H502A&
'    GP_PT_ThumbnailSize = &H5018&
'    GP_PT_ThumbnailSoftwareUsed = &H5032&
'    GP_PT_ThumbnailStripBytesCount = &H502C&
'    GP_PT_ThumbnailStripOffsets = &H5028&
'    GP_PT_ThumbnailTransferFunction = &H5031&
'    GP_PT_ThumbnailWhitePoint = &H5035&
'    GP_PT_ThumbnailWidth = &H5013&
'    GP_PT_ThumbnailYCbCrCoefficients = &H5037&
'    GP_PT_ThumbnailYCbCrPositioning = &H5039&
'    GP_PT_ThumbnailYCbCrSubsampling = &H5038&
'    GP_PT_TileByteCounts = &H145&
'    GP_PT_TileLength = &H143&
'    GP_PT_TileOffset = &H144&
'    GP_PT_TileWidth = &H142&
'    GP_PT_TransferFunction = &H12D&
'    GP_PT_TransferRange = &H156&
'    GP_PT_WhitePoint = &H13E&
'    GP_PT_XPosition = &H11E&
    GP_PT_XResolution = &H11A&
'    GP_PT_YCbCrCoefficients = &H211&
'    GP_PT_YCbCrPositioning = &H213&
'    GP_PT_YCbCrSubsampling = &H212&
'    GP_PT_YPosition = &H11F&
    GP_PT_YResolution = &H11B&
End Enum

#If False Then
    Private Const GP_PT_Artist = &H13B, GP_PT_BitsPerSample = &H102, GP_PT_CellHeight = &H109, GP_PT_CellWidth = &H108, GP_PT_ChrominanceTable = &H5091, GP_PT_ColorMap = &H140, GP_PT_ColorTransferFunction = &H501A, GP_PT_Compression = &H103, GP_PT_Copyright = &H8298, GP_PT_DateTime = &H132, GP_PT_DocumentName = &H10D, GP_PT_DotRange = &H150, GP_PT_EquipMake = &H10F, GP_PT_EquipModel = &H110, GP_PT_ExifAperture = &H9202, GP_PT_ExifBrightness = &H9203, GP_PT_ExifCfaPattern = &HA302, GP_PT_ExifColorSpace = &HA001
    Private Const GP_PT_ExifCompBPP = &H9102, GP_PT_ExifCompConfig = &H9101, GP_PT_ExifDTDigitized = &H9004, GP_PT_ExifDTDigSS = &H9292, GP_PT_ExifDTOrig = &H9003, GP_PT_ExifDTOrigSS = &H9291, GP_PT_ExifDTSubsec = &H9290, GP_PT_ExifExposureBias = &H9204, GP_PT_ExifExposureIndex = &HA215, GP_PT_ExifExposureProg = &H8822, GP_PT_ExifExposureTime = &H829A, GP_PT_ExifFileSource = &HA300, GP_PT_ExifFlash = &H9209, GP_PT_ExifFlashEnergy = &HA20B, GP_PT_ExifFNumber = &H829D, GP_PT_ExifFocalLength = &H920A
    Private Const GP_PT_ExifFocalResUnit = &HA210, GP_PT_ExifFocalXRes = &HA20E, GP_PT_ExifFocalYRes = &HA20F, GP_PT_ExifFPXVer = &HA000, GP_PT_ExifIFD = &H8769, GP_PT_ExifInterop = &HA005, GP_PT_ExifISOSpeed = &H8827, GP_PT_ExifLightSource = &H9208, GP_PT_ExifMakerNote = &H927C, GP_PT_ExifMaxAperture = &H9205, GP_PT_ExifMeteringMode = &H9207, GP_PT_ExifOECF = &H8828, GP_PT_ExifPixXDim = &HA002, GP_PT_ExifPixYDim = &HA003, GP_PT_ExifRelatedWav = &HA004, GP_PT_ExifSceneType = &HA301
    Private Const GP_PT_ExifSensingMethod = &HA217, GP_PT_ExifShutterSpeed = &H9201, GP_PT_ExifSpatialFR = &HA20C, GP_PT_ExifSpectralSense = &H8824, GP_PT_ExifSubjectDist = &H9206, GP_PT_ExifSubjectLoc = &HA214, GP_PT_ExifUserComment = &H9286, GP_PT_ExifVer = &H9000, GP_PT_ExtraSamples = &H152, GP_PT_FillOrder = &H10A, GP_PT_FrameDelay = &H5100, GP_PT_FreeByteCounts = &H121, GP_PT_FreeOffset = &H120, GP_PT_Gamma = &H301, GP_PT_GlobalPalette = &H5102, GP_PT_GpsAltitude = &H6
    Private Const GP_PT_GpsAltitudeRef = &H5, GP_PT_GpsDestBear = &H18, GP_PT_GpsDestBearRef = &H17, GP_PT_GpsDestDist = &H1A, GP_PT_GpsDestDistRef = &H19, GP_PT_GpsDestLat = &H14, GP_PT_GpsDestLatRef = &H13, GP_PT_GpsDestLong = &H16, GP_PT_GpsDestLongRef = &H15, GP_PT_GpsGpsDop = &HB, GP_PT_GpsGpsMeasureMode = &HA, GP_PT_GpsGpsSatellites = &H8, GP_PT_GpsGpsStatus = &H9, GP_PT_GpsGpsTime = &H7, GP_PT_GpsIFD = &H8825, GP_PT_GpsImgDir = &H11, GP_PT_GpsImgDirRef = &H10, GP_PT_GpsLatitude = &H2
    Private Const GP_PT_GpsLatitudeRef = &H1, GP_PT_GpsLongitude = &H4, GP_PT_GpsLongitudeRef = &H3, GP_PT_GpsMapDatum = &H12, GP_PT_GpsSpeed = &HD, GP_PT_GpsSpeedRef = &HC, GP_PT_GpsTrack = &HF, GP_PT_GpsTrackRef = &HE, GP_PT_GpsVer = &H0, GP_PT_GrayResponseCurve = &H123, GP_PT_GrayResponseUnit = &H122, GP_PT_GridSize = &H5011, GP_PT_HalftoneDegree = &H500C, GP_PT_HalftoneHints = &H141, GP_PT_HalftoneLPI = &H500A, GP_PT_HalftoneLPIUnit = &H500B, GP_PT_HalftoneMisc = &H500E, GP_PT_HalftoneScreen = &H500F
    Private Const GP_PT_HalftoneShape = &H500D, GP_PT_HostComputer = &H13C, GP_PT_ICCProfile = &H8773, GP_PT_ICCProfileDescriptor = &H302, GP_PT_ImageDescription = &H10E, GP_PT_ImageHeight = &H101, GP_PT_ImageTitle = &H320, GP_PT_ImageWidth = &H100, GP_PT_IndexBackground = &H5103, GP_PT_IndexTransparent = &H5104, GP_PT_InkNames = &H14D, GP_PT_InkSet = &H14C, GP_PT_JPEGACTables = &H209, GP_PT_JPEGDCTables = &H208, GP_PT_JPEGInterFormat = &H201, GP_PT_JPEGInterLength = &H202, GP_PT_JPEGLosslessPredictors = &H205
    Private Const GP_PT_JPEGPointTransforms = &H206, GP_PT_JPEGProc = &H200, GP_PT_JPEGQTables = &H207, GP_PT_JPEGQuality = &H5010, GP_PT_JPEGRestartInterval = &H203, GP_PT_LoopCount = &H5101, GP_PT_LuminanceTable = &H5090, GP_PT_MaxSampleValue = &H119, GP_PT_MinSampleValue = &H118, GP_PT_NewSubfileType = &HFE, GP_PT_NumberOfInks = &H14E, GP_PT_Orientation = &H112, GP_PT_PageName = &H11D, GP_PT_PageNumber = &H129, GP_PT_PaletteHistogram = &H5113, GP_PT_PhotometricInterp = &H106, GP_PT_PixelPerUnitX = &H5111
    Private Const GP_PT_PixelPerUnitY = &H5112, GP_PT_PixelUnit = &H5110, GP_PT_PlanarConfig = &H11C, GP_PT_Predictor = &H13D, GP_PT_PrimaryChromaticities = &H13F, GP_PT_PrintFlags = &H5005, GP_PT_PrintFlagsBleedWidth = &H5008, GP_PT_PrintFlagsBleedWidthScale = &H5009, GP_PT_PrintFlagsCrop = &H5007, GP_PT_PrintFlagsVersion = &H5006, GP_PT_REFBlackWhite = &H214, GP_PT_ResolutionUnit = &H128, GP_PT_ResolutionXLengthUnit = &H5003, GP_PT_ResolutionXUnit = &H5001, GP_PT_ResolutionYLengthUnit = &H5004
    Private Const GP_PT_ResolutionYUnit = &H5002, GP_PT_RowsPerStrip = &H116, GP_PT_SampleFormat = &H153, GP_PT_SamplesPerPixel = &H115, GP_PT_SMaxSampleValue = &H155, GP_PT_SMinSampleValue = &H154, GP_PT_SoftwareUsed = &H131, GP_PT_SRGBRenderingIntent = &H303, GP_PT_StripBytesCount = &H117, GP_PT_StripOffsets = &H111, GP_PT_SubfileType = &HFF, GP_PT_T4Option = &H124, GP_PT_T6Option = &H125, GP_PT_TargetPrinter = &H151, GP_PT_ThreshHolding = &H107, GP_PT_ThumbnailArtist = &H5034, GP_PT_ThumbnailBitsPerSample = &H5022
    Private Const GP_PT_ThumbnailColorDepth = &H5015, GP_PT_ThumbnailCompressedSize = &H5019, GP_PT_ThumbnailCompression = &H5023, GP_PT_ThumbnailCopyRight = &H503B, GP_PT_ThumbnailData = &H501B, GP_PT_ThumbnailDateTime = &H5033, GP_PT_ThumbnailEquipMake = &H5026, GP_PT_ThumbnailEquipModel = &H5027, GP_PT_ThumbnailFormat = &H5012, GP_PT_ThumbnailHeight = &H5014, GP_PT_ThumbnailImageDescription = &H5025, GP_PT_ThumbnailImageHeight = &H5021, GP_PT_ThumbnailImageWidth = &H5020, GP_PT_ThumbnailOrientation = &H5029, GP_PT_ThumbnailPhotometricInterp = &H5024
    Private Const GP_PT_ThumbnailPlanarConfig = &H502F, GP_PT_ThumbnailPlanes = &H5016, GP_PT_ThumbnailPrimaryChromaticities = &H5036, GP_PT_ThumbnailRawBytes = &H5017, GP_PT_ThumbnailRefBlackWhite = &H503A, GP_PT_ThumbnailResolutionUnit = &H5030, GP_PT_ThumbnailResolutionX = &H502D, GP_PT_ThumbnailResolutionY = &H502E, GP_PT_ThumbnailRowsPerStrip = &H502B, GP_PT_ThumbnailSamplesPerPixel = &H502A, GP_PT_ThumbnailSize = &H5018, GP_PT_ThumbnailSoftwareUsed = &H5032, GP_PT_ThumbnailStripBytesCount = &H502C, GP_PT_ThumbnailStripOffsets = &H5028
    Private Const GP_PT_ThumbnailTransferFunction = &H5031, GP_PT_ThumbnailWhitePoint = &H5035, GP_PT_ThumbnailWidth = &H5013, GP_PT_ThumbnailYCbCrCoefficients = &H5037, GP_PT_ThumbnailYCbCrPositioning = &H5039, GP_PT_ThumbnailYCbCrSubsampling = &H5038, GP_PT_TileByteCounts = &H145, GP_PT_TileLength = &H143, GP_PT_TileOffset = &H144, GP_PT_TileWidth = &H142, GP_PT_TransferFunction = &H12D, GP_PT_TransferRange = &H156, GP_PT_WhitePoint = &H13E, GP_PT_XPosition = &H11E, GP_PT_XResolution = &H11A, GP_PT_YCbCrCoefficients = &H211
    Private Const GP_PT_YCbCrPositioning = &H213, GP_PT_YCbCrSubsampling = &H212, GP_PT_YPosition = &H11F, GP_PT_YResolution = &H11B
#End If

'Private Enum GP_PropertyTagType
'    GP_PTT_Byte = 1
'    GP_PTT_ASCII = 2
'    GP_PTT_Short = 3
'    GP_PTT_Long = 4
'    GP_PTT_Rational = 5
'    GP_PTT_Undefined = 7
'    GP_PTT_SLONG = 9
'    GP_PTT_SRational = 10
'End Enum
'
'#If False Then
'    Private Const GP_PTT_Byte = 1, GP_PTT_ASCII = 2, GP_PTT_Short = 3, GP_PTT_Long = 4, GP_PTT_Rational = 5, GP_PTT_Undefined = 7, GP_PTT_SLONG = 9, GP_PTT_SRational = 10
'#End If

Public Enum GP_RotateFlip
    GP_RF_NoneFlipNone = 0
    GP_RF_90FlipNone = 1
    GP_RF_180FlipNone = 2
    GP_RF_270FlipNone = 3
    GP_RF_NoneFlipX = 4
    GP_RF_90FlipX = 5
    GP_RF_180FlipX = 6
    GP_RF_270FlipX = 7
    GP_RF_NoneFlipY = GP_RF_180FlipX
    GP_RF_90FlipY = GP_RF_270FlipX
    GP_RF_180FlipY = GP_RF_NoneFlipX
    GP_RF_270FlipY = GP_RF_90FlipX
    GP_RF_NoneFlipXY = GP_RF_180FlipNone
    GP_RF_90FlipXY = GP_RF_270FlipNone
    GP_RF_180FlipXY = GP_RF_NoneFlipNone
    GP_RF_270FlipXY = GP_RF_90FlipNone
End Enum

#If False Then
    Private Const GP_RF_NoneFlipNone = 0, GP_RF_90FlipNone = 1, GP_RF_180FlipNone = 2, GP_RF_270FlipNone = 3, GP_RF_NoneFlipX = 4, GP_RF_90FlipX = 5, GP_RF_180FlipX = 6, GP_RF_270FlipX = 7, GP_RF_NoneFlipY = GP_RF_180FlipX
    Private Const GP_RF_90FlipY = GP_RF_270FlipX, GP_RF_180FlipY = GP_RF_NoneFlipX, GP_RF_270FlipY = GP_RF_90FlipX, GP_RF_NoneFlipXY = GP_RF_180FlipNone, GP_RF_90FlipXY = GP_RF_270FlipNone, GP_RF_180FlipXY = GP_RF_NoneFlipNone, GP_RF_270FlipXY = GP_RF_90FlipNone
#End If

Public Enum GP_SmoothingMode
    GP_SM_Invalid = GP_QM_Invalid
    GP_SM_Default = GP_QM_Default
    GP_SM_HighSpeed = GP_QM_Low
    GP_SM_HighQuality = GP_QM_High
    GP_SM_None = 3&
    GP_SM_Antialias = 4&
End Enum

#If False Then
    Private Const GP_SM_Invalid = GP_QM_Invalid, GP_SM_Default = GP_QM_Default, GP_SM_HighSpeed = GP_QM_Low, GP_SM_HighQuality = GP_QM_High, GP_SM_None = 3, GP_SM_Antialias = 4
#End If

'GDI+ string format settings.  Note that "near" and "far" monikers are used instead of left/right;
' this allows handling RTL languages in a more natural way.
Public Enum GP_StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
    
    'IMPORTANT NOTE: GDI+ cannot render justified text.  This setting *ONLY* works with PhotoDemon's advanced text renderer.
    StringAlignmentJustify = 3
End Enum

#If False Then
    Private Const StringAlignmentNear = 0, StringAlignmentCenter = 1, StringAlignmentFar = 2, StringAlignmentJustify = 3
#End If

Public Enum GP_Unit
    GP_U_World = 0&
    GP_U_Display = 1&
    GP_U_Pixel = 2&
    GP_U_Point = 3&
    GP_U_Inch = 4&
    GP_U_Document = 5&
    GP_U_Millimeter = 6&
End Enum

#If False Then
    Private Const GP_U_World = 0, GP_U_Display = 1, GP_U_Pixel = 2, GP_U_Point = 3, GP_U_Inch = 4, GP_U_Document = 5, GP_U_Millimeter = 6
#End If

Public Enum GP_WrapMode
    GP_WM_Tile = 0
    GP_WM_TileFlipX = 1
    GP_WM_TileFlipY = 2
    GP_WM_TileFlipXY = 3
    GP_WM_Clamp = 4
End Enum

#If False Then
    Private Const GP_WM_Tile = 0, GP_WM_TileFlipX = 1, GP_WM_TileFlipY = 2, GP_WM_TileFlipXY = 3, GP_WM_Clamp = 4
#End If

'GDI+ uses a modified bitmap struct when performing things like raster format conversions
Public Type GP_BitmapData
    BD_Width As Long
    BD_Height As Long
    BD_Stride As Long
    BD_PixelFormat As GP_PixelFormat
    BD_Scan0 As Long
    BD_Reserved As Long
End Type

'GDI interop is made easier by declaring a few GDI-specific structs
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

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQuad
End Type

'This (stupid) type is used so we can take advantage of LSet when performing some conversions
Private Type tmpLong
    lngResult As Long
End Type

'On GDI+ v1.1 or later, certain effects can be rendered via GDI+.  Note that these are buggy and *not* well-tested,
' so we avoid them in PD except for curiosity and testing purposes.
Private Type GP_BlurParams
    BP_Radius As Single
    BP_ExpandEdge As Long
End Type

'Exporting images via GDI+ is a big headache.  A number of convoluted structs are required if the user
' wants to custom-set any image properties.
Private Type GP_EncoderParameter
    EP_GUID(0 To 15) As Byte
    EP_NumOfValues As Long
    EP_ValueType As GP_EncoderValueType
    EP_ValuePtr As Long
End Type

Private Type GP_EncoderParameters
    EP_Count As Long
    EP_Parameter As GP_EncoderParameter
End Type

Private Type GP_ImageCodecInfo
    IC_ClassID(0 To 15) As Byte
    IC_FormatID(0 To 15) As Byte
    IC_CodecName As Long
    IC_DllName As Long
    IC_FormatDescription As Long
    IC_FilenameExtension As Long
    IC_MimeType As Long
    IC_Flags As Long
    IC_Version As Long
    IC_SigCount As Long
    IC_SigSize As Long
    IC_SigPattern As Long
    IC_SigMask As Long
End Type

'Helper structs for metafile headers.  IMPORTANT NOTE!  There are probably struct alignment issues with these structs,
' as they are legacy structs that intermix 16- and 32-bit datatypes.  I do not need these at present (I only need them
' as part of an unused union in a GDI+ metafile type), so I have not tested them thoroughly.  Use at your own risk.
Private Type GDI_SizeL
    cx As Long
    cy As Long
End Type

Private Type GDI_MetaHeader
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

Private Type GDIP_EnhMetaHeader3
    itype As Long
    nSize As Long
    rclBounds As RectL
    rclFrame As RectL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As GDI_SizeL
    szlMillimeters As GDI_SizeL
End Type

Private Type GP_MetafileHeader_UNION
    muWmfHeader As GDI_MetaHeader
    muEmfHeader As GDIP_EnhMetaHeader3
End Type

'Want additional information on a metafile-type Image object?  This struct contains basic header data.
' IMPORTANT NOTE: please see the previous comment on struct alignment.  I can't guarantee that anything past
' the mfOrigHeader union is aligned correctly; use those at your own risk.
Private Type GP_MetafileHeader
    mfType As GP_MetafileType
    mfSize As Long
    mfVersion As Long
    mfEmfPlusFlags As Long
    mfDpiX As Single
    mfDpiY As Single
    mfBoundsX As Long
    mfBoundsY As Long
    mfBoundsWidth As Long
    mfBoundsHeight As Long
    mfOrigHeader As GP_MetafileHeader_UNION
    mfEmfPlusHeaderSize As Long
    mfLogicalDpiX As Long
    mfLogicalDpiY As Long
End Type

'GDI+ image properties
Public Type GP_PropertyItem
    propID As GP_PropertyTag    'Tag identifier
    propLength As Long          'Length of the property value, in bytes
    propType As Integer         'Type of tag value (one of GP_PropertyTagType)
    ignorePadding As Integer
    propValue As Long           'Property value or pointer to property value, contingent on propType, above
End Type

'Like image formats, export encoder properties are also defined by GUID.  These values come from
' the Win 8.1 version of gdiplusimaging.h.  Note that some are restricted to GDI+ v1.1.
'Private Const GP_EP_ChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const GP_EP_ColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const GP_EP_Compression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
'Private Const GP_EP_LuminanceTable As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const GP_EP_Quality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
'Private Const GP_EP_RenderMethod As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
'Private Const GP_EP_SaveFlag As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
'Private Const GP_EP_ScanMethod As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
'Private Const GP_EP_Transformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const GP_EP_Version As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"

'THESE ENCODER PROPERTIES REQUIRE GDI+ v1.1 OR LATER!
'Private Const GP_EP_ColorSpace As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
'Private Const GP_EP_SaveAsCMYK As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"

'Multi-frame (GIF) and multi-page (TIFF) files support retrieval of individual pages via
' something Microsoft confusingly calls "frame dimensions".  Frame retrieval functions
' require you to specify which kind of frame you want to retrieve; these GUIDs control that.
Private Const GP_FD_Page As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
'Private Const GP_FD_Resolution As String = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"    'used for ICONs; PD parses those manually
Private Const GP_FD_Time As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"

'GDI+ uses GUIDs to define image formats.  VB6 doesn't let us predeclare byte arrays (at least not easily),
' so we save ourselves the trouble and just use string versions.
Private Const GP_FF_GUID_Undefined = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_MemoryBMP = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_EMF = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_WMF = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
'Private Const GP_FF_GUID_EXIF = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"   'Unused
Private Const GP_FF_GUID_Icon = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"

'Core GDI+ functions:
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef gdipToken As Long, ByRef startupStruct As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GP_Result
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal gdipToken As Long) As GP_Result

'Object creation/destruction/property functions
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcRect As RectL, ByVal lockMode As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat, ByRef srcBitmapData As GP_BitmapData) As GP_Result
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcBitmapData As GP_BitmapData) As GP_Result

Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal newPixelFormat As GP_PixelFormat, ByVal hSrcBitmap As Long, ByRef hDstBitmap As Long) As GP_Result

Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef origGDIBitmapInfo As BITMAPINFO, ByVal ptrToPixels As Long, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal bmpWidth As Long, ByVal bmpHeight As Long, ByVal bmpStride As Long, ByVal bmpPixelFormat As GP_PixelFormat, ByVal ptrToPixels As Long, ByRef dstGdipBitmap As Long) As GP_Result
'Private Declare Function GdipCreateCachedBitmap Lib "gdiplus" (ByVal hBitmap As Long, ByVal hGraphics As Long, ByRef dstCachedBitmap As Long) As GP_Result
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef dstGraphics As Long) As GP_Result
Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, ByRef dstGraphics As Long) As GP_Result
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef dstImageAttributes As Long) As GP_Result
'Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal brushWrapMode As GP_WrapMode, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RectF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As GP_WrapMode, ByRef dstLineGradientBrush As Long) As GP_Result
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As GP_Result
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal srcColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal hImage As Long, ByVal textureWrapMode As GP_WrapMode, ByRef dstTexture As Long) As GP_Result

Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GP_Result
'Private Declare Function GdipDeleteCachedBitmap Lib "gdiplus" (ByVal hCachedBitmap As Long) As GP_Result
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GP_Result
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImageAttributes As Long) As GP_Result

Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As GP_Result
'Private Declare Function GdipDrawCachedBitmap Lib "gdiplus" (ByVal hGraphics As Long, ByVal hCachedBitmap As Long, ByVal x As Long, ByVal y As Long) As GP_Result
Private Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal y As Single) As GP_Result
Private Declare Function GdipDrawImageI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal y As Long) As GP_Result
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal y As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As GP_Result
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal y As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As GP_Result
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointFloats As Long, ByVal dstPtCount As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointInts As Long, ByVal dstPtCount As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GP_Result
Private Declare Function GdipDrawLines Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result

Private Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipFillPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipFillRegion Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hRegion As Long) As GP_Result

'Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal hImage As Long, ByRef dstRectF As RectF, ByRef dstUnit As GP_Unit) As GP_Result
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Single, ByRef dstHeight As Single) As GP_Result
'Private Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
'Private Declare Function GdipGetImageDecodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef dstHeight As Long) As GP_Result
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstHResolution As Single) As GP_Result
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, ByRef dstPixelFormat As GP_PixelFormat) As GP_Result
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDstGuid As Long) As GP_Result
Private Declare Function GdipGetImageType Lib "gdiplus" (ByVal srcImage As Long, ByRef dstImageType As GP_ImageType) As GP_Result
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstVResolution As Single) As GP_Result
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Long) As GP_Result
Private Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal hMetafile As Long, ByRef dstHeader As GP_MetafileHeader) As GP_Result
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByVal srcPropertySize As Long, ByVal ptrToDstBuffer As Long) As GP_Result
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As GP_PropertyTag, ByRef dstPropertySize As Long) As GP_Result

Private Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDimensionGuid As Long, ByRef dstCount As Long) As GP_Result
Private Declare Function GdipImageGetFrameDimensionsCount Lib "gdiplus" (ByVal hImage As Long, ByRef dstCount As Long) As GP_Result
Private Declare Function GdipImageGetFrameDimensionsList Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDimensionGuids As Long, ByVal srcCount As Long) As GP_Result
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rotateFlipType As GP_RotateFlip) As GP_Result
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDimensionGuid As Long, ByVal frameIndex As Long) As GP_Result

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal ptrSrcFilename As Long, ByRef dstGdipImage As Long) As GP_Result
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal srcIStream As Long, ByRef dstGdipImage As Long) As GP_Result

Private Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (ByRef srcFontCollection As Long) As GP_Result
Private Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (ByRef dstFontCollection As Long) As GP_Result
Private Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal dstFontCollection As Long, ByVal lSrcFilename As Long) As GP_Result

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToFilename As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal hImage As Long, ByVal dstIStream As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result

Private Declare Function GdipSetClipRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingMode As GP_CompositingMode) As GP_Result
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal newWrapMode As GP_WrapMode, ByVal argbOfClampMode As Long, ByVal bClampMustBeZero As Long) As GP_Result
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal typeOfAdjustment As GP_ColorAdjustType, ByVal enableSeparateAdjustmentFlag As Long, ByVal ptrToColorMatrix As Long, ByVal ptrToGrayscaleMatrix As Long, ByVal extraColorMatrixFlags As GP_ColorMatrixFlags) As GP_Result
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipSetLineGammaCorrection Lib "gdiplus" (ByVal hBrush As Long, ByVal useGammaCorrection As Long) As GP_Result
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal hMetafile As Long, ByVal metafileRasterizationLimitDpi As Long) As GP_Result
Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoints As PointFloat) As GP_Result
Private Declare Function GdipSetPathGradientGammaCorrection Lib "gdiplus" (ByVal hBrush As Long, ByVal useGammaCorrection As Long) As GP_Result
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal hBrush As Long, ByVal hMatrix As Long) As GP_Result

'Some GDI+ functions are *only* supported on GDI+ 1.1, which first shipped with Vista (but requires explicit activation
' via manifest, and as such, is unavailable to PD until Win 7).  Take care to confirm the availability of these functions
' before using them.
Private Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result
'Private Declare Function GdipConvertToEmfPlusToFile Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal filenamePointer As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef dstEffect As Long) As GP_Result
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal hEffect As Long) As GP_Result
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByRef drawRect As RectF, ByVal hTransformMatrix As Long, ByVal hEffect As Long, ByVal hImageAttributes As Long, ByVal srcUnit As GP_Unit) As GP_Result
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal hEffect As Long, ByRef srcParams As Any, ByVal srcParamSize As Long) As GP_Result

'Non-GDI+ helper functions:
Private Declare Function CLSIDFromString Lib "ole32" (ByVal ptrToGuidString As Long, ByVal ptrToByteArray As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByVal ptrToDstStream As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal srcIStream As Long, ByRef dstHGlobal As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long
Private Declare Function StringFromCLSID Lib "ole32" (ByVal ptrToGuid As Long, ByRef ptrToDstString As Long) As Long

'Internally cached values:

'Startup values
Private m_GDIPlusToken As Long, m_GDIPlus11Available As Boolean

'Some GDI+ functions require world transformation data.  This dummy graphics container is used to host
' any such transformations. It is created when GDI+ is initialized, and destroyed when GDI+ is released.
' To be a good citizen, please undo any world transforms before a function releases.  This ensures that
' subsequent functions don't get messed up.
Private m_TransformDIB As pdDIB, m_TransformGraphics As Long

'To modify opacity in GDI+, an image attributes matrix is used.  Rather than recreating one every time
' an alpha operation is required, we simply create a default identity matrix at initialization,
' then re-use it as necessary.
Private m_AttributesMatrix() As Single

'When loading multi-page (TIFF) or multi-frame (GIF) images, we first perform a default load operation
' on the first page/frame. (This gives the user something to work with if subsequent pages fail,
' which is worryingly possible especially given the complexities of loading TIFFs.)  PD's central load
' function can then do whatever it needs to - e.g. prompt the user for desired load behavior - and
' notify GDI+ of the result.  If the user wants more pages/frames from the file, we don't have to load
' it again; instead, we can just activate subsequent pages in turn, carrying on where we first left off.
Private m_hMultiPageImage As Long, m_OriginalFIF As PD_IMAGE_FORMAT

'When loading GIFs, we need to cache some extra GIF-related metadata (e.g. frame times).
' This metadata gets embedded into a parent pdImage object, and reused at export time as relevant.
Private m_FrameTimes() As Long, m_FrameCount As Long

'If the user adds custom fonts at run-time, we need to maintain them in a persistent FontCollection object.
Private m_UserFontCollection As Long

'Use GDI+ to resize a DIB.  (Technically, to copy a resized portion of a source image into a destination image.)
' The call is formatted similar to StretchBlt, as it used to replace StretchBlt when working with 32bpp data.
' FOR FUTURE REFERENCE: after a bunch of profiling on my Win 7 PC, I can state with 100% confidence that
' the HighQualityBicubic interpolation mode is actually the fastest mode for downsizing 32bpp images.  I have no idea
' why this is, but many, many iterative tests confirmed it.  Stranger still, in descending order after that, the fastest
' algorithms are: HighQualityBilinear, Bilinear, Bicubic.  Regular bicubic interpolation is some 4x slower than the
' high quality mode!!
Public Function GDIPlusResizeDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal interpolationType As GP_InterpolationMode, Optional ByVal pixelOffsetMode As PD_2D_PixelOffset = P2_PO_Normal) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    GDIPlusResizeDIB = True

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim hGdipGraphics As Long, hGdipBitmap As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, hGdipGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB hGdipBitmap, srcDIB
    
    'hGdipGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(hGdipGraphics, interpolationType) = GP_OK Then
    
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        GdipSetImageAttributesWrapMode imgAttributesHandle, GP_WM_TileFlipXY, 0&, 0&
        GdipSetCompositingQuality hGdipGraphics, GP_CQ_AssumeLinear
        If (pixelOffsetMode = P2_PO_Normal) Then GdipSetPixelOffsetMode hGdipGraphics, GP_POM_HighSpeed Else GdipSetPixelOffsetMode hGdipGraphics, GP_POM_HighQuality
        
        'Perform the resize
        If (GdipDrawImageRectRectI(hGdipGraphics, hGdipBitmap, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle) <> 0) Then
            GDIPlusResizeDIB = False
        End If
        
        'Release our image attributes object
        If (imgAttributesHandle <> 0) Then GdipDisposeImageAttributes imgAttributesHandle
        
    Else
        GDIPlusResizeDIB = False
    End If
    
    'Release both the destination graphics object and the source bitmap object
    If (hGdipGraphics <> 0) Then GdipDeleteGraphics hGdipGraphics
    If (hGdipBitmap <> 0) Then GdipDisposeImage hGdipBitmap
    
    'GDI+ draw functions always result in a premultiplied image
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'Free the destination DIB from its DC, as it may not be required again for some time
    dstDIB.FreeFromDC
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format$((Timer - profileTime) * 1000, "0000.00")
    
End Function

'Simpler rotate/flip function, and limited to the constants specified by the enum.
Public Function GDIPlusRotateFlipDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotationType As GP_RotateFlip) As Boolean
    
    'Wrap a GDI+ bitmap handle around the source DIB
    Dim hGdipBitmap As Long
    If GetGdipBitmapHandleFromDIB(hGdipBitmap, srcDIB) Then
        
        'Apply the rotation
        If (GdipImageRotateFlip(hGdipBitmap, rotationType) = GP_OK) Then
            
            'Resize the target DIB to match
            Dim newWidth As Long, newHeight As Long
            GdipGetImageWidth hGdipBitmap, newWidth
            GdipGetImageHeight hGdipBitmap, newHeight
    
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            If (dstDIB.GetDIBWidth <> newWidth) Or (dstDIB.GetDIBHeight <> newHeight) Or (dstDIB.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
                dstDIB.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth, 0
            End If
            
            'The destination DIB will *always* have premultiplied alpha, since we're using a
            ' GdipDraw* function to perform the fast rotate.
            dstDIB.SetInitialAlphaPremultiplicationState True
            
            'Copy the rotated source into the destination DIB
            Dim hGraphics As Long
            If (GdipCreateFromHDC(dstDIB.GetDIBDC, hGraphics) = GP_OK) Then
        
                'For performance reasons, allow the renderer to copy pixels instead of blending them (as the target image
                ' is guaranteed to be empty).  Note that we don't care if this fails; the result will be correct
                ' either way.
                GdipSetCompositingMode hGraphics, GP_CM_SourceCopy
        
                'Render the rotated image
                GDIPlusRotateFlipDIB = GdipDrawImageI(hGraphics, hGdipBitmap, 0, 0) = GP_OK
                
                'Release both the destination graphics object and the source bitmap object
                GdipDeleteGraphics hGraphics
                
            End If
            
        End If
        
        GdipDisposeImage hGdipBitmap
        
    End If
    
End Function

'In-place rotate/flip function, which reduces the need for extra allocations
Public Function GDIPlusRotateFlip_InPlace(ByRef srcDIB As pdDIB, ByVal rotationType As GP_RotateFlip) As Boolean
    Dim hGdipBitmap As Long
    If GetGdipBitmapHandleFromDIB(hGdipBitmap, srcDIB) Then
        GDIPlusRotateFlip_InPlace = (GdipImageRotateFlip(hGdipBitmap, rotationType) = GP_OK)
        GdipDisposeImage hGdipBitmap
    End If
End Function

'Use GDI+ to blur a DIB with variable radius
Public Function GDIPlusBlurDIB(ByRef dstDIB As pdDIB, ByVal blurRadius As Long, ByVal rLeft As Double, ByVal rTop As Double, ByVal rWidth As Double, ByVal rHeight As Double) As Boolean

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim hGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    
    'Next, we need a temporary copy of the image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, dstDIB
    
    'Create a GDI+ blur effect object
    Dim hEffect As Long
    If (GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, hEffect) = 0) Then
        
        'Next, create a compatible set of blur parameters and pass those to the GDI+ blur object
        Dim tmpParams As GP_BlurParams
        tmpParams.BP_Radius = CSng(blurRadius)
        tmpParams.BP_ExpandEdge = 0
    
        If GdipSetEffectParameters(hEffect, tmpParams, Len(tmpParams)) = 0 Then
    
            'The DrawImageFX call requires a target rect.  Create one now (in GDI+ format, e.g. RECTF)
            Dim tmpRect As RectF
            tmpRect.Left = rLeft
            tmpRect.Top = rTop
            tmpRect.Width = rWidth
            tmpRect.Height = rHeight
            
            'Create a temporary GDI+ transformation matrix as well
            Dim tmpMatrix As pd2DTransform
            Set tmpMatrix = New pd2DTransform
            
            'Attempt to render the blur effect
            Dim GDIPlusDebug As GP_Result
            GDIPlusDebug = GdipDrawImageFX(hGraphics, tBitmap, tmpRect, tmpMatrix.GetHandle(True), hEffect, 0&, GP_U_Pixel)
            
            GDIPlusBlurDIB = (GDIPlusDebug = GP_OK)
            If (Not GDIPlusBlurDIB) Then PDDebug.LogAction "GDI+ failed to render blur effect (Error Code %1).", GDIPlusDebug
            
            'Delete our temporary transformation matrix
            Set tmpMatrix = Nothing
            
        Else
            GDIPlusBlurDIB = False
            PDDebug.LogAction "GDI+ failed to set effect parameters."
        End If
    
        'Delete our GDI+ blur object
        GdipDeleteEffect hEffect
    
    Else
        GDIPlusBlurDIB = False
        PDDebug.LogAction "GDI+ failed to create blur effect object"
    End If
        
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics hGraphics
    GdipDisposeImage tBitmap
    
End Function

'Use GDI+ to fill a DIB with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full DIB alpha
' in a single pass.
Public Function GDIPlusFillDIBRect(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal cOpacity As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    
    If useAA Then
        GdipSetSmoothingMode hGraphics, GP_SM_Antialias
    Else
        GdipSetSmoothingMode hGraphics, GP_SM_None
    End If
    
    GdipSetCompositingMode hGraphics, dstFillMode
    
    'Create a solid fill brush from the source image
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, cOpacity)
    
    If (hBrush <> 0) Then
        GdipFillRectangle hGraphics, hBrush, x1, y1, xWidth, yHeight
        ReleaseGDIPlusBrush hBrush
    Else
        Debug.Print "WARNING!  GDIPlusFillDIBRect failed because hBrush was null."
    End If
    
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillDIBRect = True

End Function

'Given a source DIB, fill it with the alpha checkerboard pattern.  32bpp images can then be alpha blended onto it.
' Note that - by design - this function assumes a COPY operation, not a traditional PAINT operation.  Copying is faster,
' and there should never be a need to alpha-blend the checkerboard pattern atop something.
Public Function GDIPlusFillDIBRect_Pattern(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal bltWidth As Single, ByVal bltHeight As Single, ByRef srcDIB As pdDIB, Optional ByVal useThisDCInstead As Long = 0, Optional ByVal fixBoundaryPainting As Boolean = False, Optional ByVal noAntialiasing As Boolean = False) As Boolean
    
    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    
    If (useThisDCInstead <> 0) Then
        GdipCreateFromHDC useThisDCInstead, hGraphics
    Else
        GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    End If
    
    If noAntialiasing Then GdipSetSmoothingMode hGraphics, GP_SM_None Else GdipSetSmoothingMode hGraphics, GP_SM_Antialias
    GdipSetCompositingQuality hGraphics, GP_CQ_AssumeLinear
    GdipSetPixelOffsetMode hGraphics, GP_POM_HighSpeed
    GdipSetCompositingMode hGraphics, GP_CM_SourceCopy
    GdipSetClipRect hGraphics, x1, y1, bltWidth, bltHeight, GP_CM_Replace
    
    'Create a texture fill brush from the source image
    Dim srcBitmap As Long, hBrush As Long
    GetGdipBitmapHandleFromDIB srcBitmap, srcDIB
    GdipCreateTexture srcBitmap, GP_WM_Tile, hBrush
    
    'Because pattern fills are prone to boundary overflow when used with transparent overlays, the caller can
    ' have us restrict painting to the interior integer region only.)
    If fixBoundaryPainting Then
        
        Dim xDif As Single, yDif As Single
        xDif = x1 - Int(x1)
        yDif = y1 - Int(y1)
        bltWidth = Int(bltWidth - xDif - 0.5)
        bltHeight = Int(bltHeight - yDif - 0.5)
        
    End If
    
    'Apply the brush
    GdipFillRectangle hGraphics, hBrush, x1, y1, bltWidth, bltHeight
    
    'Release all created objects
    ReleaseGDIPlusBrush hBrush
    GdipDisposeImage srcBitmap
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillDIBRect_Pattern = True
    
End Function

'Use GDI+ to fill an arbitrary DC with an arbitrary GDI+ brush
Public Function GDIPlusFillDC_Brush(ByRef dstDC As Long, ByVal srcBrushHandle As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    If useAA Then
        GdipSetSmoothingMode hGraphics, GP_SM_Antialias
    Else
        GdipSetSmoothingMode hGraphics, GP_SM_None
    End If
    
    GdipSetCompositingMode hGraphics, dstFillMode
    
    'Apply the brush
    GdipFillRectangle hGraphics, srcBrushHandle, x1, y1, xWidth, yHeight
    
    'Release all created objects
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillDC_Brush = True

End Function

'Use GDI+ to quickly convert a 24bpp DIB to 32bpp with solid alpha channel
Public Sub GDIPlusConvertDIB24to32(ByRef dstDIB As pdDIB)
    
    If (dstDIB.GetDIBColorDepth = 32) Then Exit Sub
    
    Dim dstBitmap As Long, srcBitmap As Long
    
    'Create a temporary source DIB to hold the intermediate copy of the image
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB dstDIB
    
    'We know the source DIB is 24bpp, so use GdipCreateBitmapFromGdiDib to obtain a handle
    Dim imgHeader As BITMAPINFO
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = srcDIB.GetDIBColorDepth
        .Width = srcDIB.GetDIBWidth
        .Height = -srcDIB.GetDIBHeight
    End With
    
    GdipCreateBitmapFromGdiDib imgHeader, srcDIB.GetDIBPointer, srcBitmap
    
    'Next, recreate the destination DIB as 32bpp
    dstDIB.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, , 255
    
    'Clone the bitmap area from source to destination, while converting format as necessary
    Dim gdipReturn As Long
    gdipReturn = GdipCloneBitmapAreaI(0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, GP_PF_32bppPARGB, srcBitmap, dstBitmap)
    GdipDisposeImage srcBitmap
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim hGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    
    'Paint the converted image to the destination
    GdipDrawImage hGraphics, dstBitmap, 0, 0
    
    'The target image will always have premultiplied alpha (not really relevant, as the source is 24-bpp, but this
    ' lets us use various accelerated codepaths throughout the project).
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'Release our bitmap copies and GDI+ instances
    GdipDisposeImage dstBitmap
    GdipDeleteGraphics hGraphics
 
End Sub

'Use GDI+ to load an image file.  Pretty bare-bones, but should be sufficient for any supported image type.
Public Function GDIPlusLoadPicture(ByVal srcFilename As String, ByRef dstDIB As pdDIB, Optional ByRef dstImage As pdImage = Nothing, Optional ByRef numOfPages As Long = 1, Optional ByVal nonInteractiveMode As Boolean = False, Optional ByVal overrideParameters As String = vbNullString) As Boolean

    'Used to hold the return values of various GDI+ calls
    Dim GDIPlusReturn As GP_Result
    
    'Use GDI+ to load the image
    Dim hImage As Long
    GDIPlusReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    If (GDIPlusReturn <> GP_OK) Then
        If (hImage <> 0) Then GdipDisposeImage hImage
        InternalGDIPlusError , "GDIPlusLoadPictureFailure", GDIPlusReturn
        GDIPlusLoadPicture = False
        Exit Function
    End If
    
    'If we're still here, the image (probably) loaded successfully.  Create a destination DIB as necessary.
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Retrieve the image's format as a GUID
    Dim formatGUID(0 To 15) As Byte
    GdipGetImageRawFormat hImage, VarPtr(formatGUID(0))
    
    'Convert the GUID into a string
    Dim imgStringPointer As Long, imgFormatGuidString As String
    StringFromCLSID VarPtr(formatGUID(0)), imgStringPointer
    imgFormatGuidString = Strings.StringFromCharPtr(imgStringPointer, True)
    
    'And finally, convert the string into an FIF long
    Dim imgFormatFIF As PD_IMAGE_FORMAT
    imgFormatFIF = GetFIFFromGUID(imgFormatGuidString)
    m_OriginalFIF = imgFormatFIF
    
    'Metafiles require special consideration; set that flag in advance
    Dim isMetafile As Boolean
    isMetafile = (imgFormatFIF = PDIF_EMF) Or (imgFormatFIF = PDIF_WMF)
    
    'Multi-page TIFFs also require special consideration; set another flag in advance.
    ' (Note that an obnoxious reliance on GUIDs forces us to cache a persistent object identifying frame parameters;
    '  we will reuse this later in the function, as necessary.)
    Dim frameDimensionID(0 To 15) As Byte
    CLSIDFromString StrPtr(GP_FD_Page), VarPtr(frameDimensionID(0))
    
    Dim isMultiPage As Boolean, imgPageCount As Long
    
    'TIFFs use a simple frame count
    If (imgFormatFIF = PDIF_TIFF) Then
        If (GdipImageGetFrameCount(hImage, VarPtr(frameDimensionID(0)), imgPageCount) = GP_OK) Then isMultiPage = (imgPageCount > 1)
    
    'GIFs are more complicated to retrieve frame count
    ElseIf (imgFormatFIF = PDIF_GIF) Then
        
        'First, retrieve frame dimension count
        Dim numFrameDimensions As Long
        If (GdipImageGetFrameDimensionsCount(hImage, numFrameDimensions) = GP_OK) Then
            
            'Frame dimension count should be at least 1
            If (numFrameDimensions > 0) Then
            
                'Retrieve GUIDs for frame dimensions, then pull the frame count using that
                Dim lstGuids() As Byte
                ReDim lstGuids(0 To 16 * numFrameDimensions - 1) As Byte
                If (GdipImageGetFrameDimensionsList(hImage, VarPtr(lstGuids(0)), numFrameDimensions) = GP_OK) Then
                    If (GdipImageGetFrameCount(hImage, VarPtr(lstGuids(0)), imgPageCount) = GP_OK) Then isMultiPage = (imgPageCount > 1)
                End If
            
            End If
            
        End If
        
    End If
    
    'We're now going to retrieve various image properties using standard GDI+ property retrieval functions.
    Dim tmpPropHeader As GP_PropertyItem, tmpPropBuffer() As Byte
    
    'If this is an animated GIF, load all frame times and cache them; if the rest of the load process
    ' is successful, we'll embed these inside the parent pdImage object.
    If isMultiPage And (imgFormatFIF = PDIF_GIF) Then
    
        m_FrameCount = imgPageCount
        ReDim m_FrameTimes(0) As Long
        
        If GDIPlus_ImageGetProperty(hImage, GP_PT_FrameDelay, tmpPropHeader, tmpPropBuffer) Then
            
            'Ensure the retrieved data matches our expected size
            If (UBound(tmpPropBuffer) = (m_FrameCount * 4) - 1) Then
                
                'Copy the frame times into our buffer, and because they're in hundredths of a second
                ' (WTF GIF?), convert them to more useful ms measurements.
                ReDim m_FrameTimes(0 To m_FrameCount - 1) As Long
                CopyMemoryStrict VarPtr(m_FrameTimes(0)), VarPtr(tmpPropBuffer(0)), m_FrameCount * 4
                Erase tmpPropBuffer
                
                Dim cFrame As Long
                For cFrame = 0 To m_FrameCount - 1
                    m_FrameTimes(cFrame) = m_FrameTimes(cFrame) * 10
                Next cFrame
                
            End If
            
        End If
    
    End If
    
    'Look for an ICC profile and cache the result; if the image *does* have an embedded profile, we will use
    ' it in a subsequent function as part of transforming pixel bytes to a standard 32-bit buffer.
    Dim imgHasIccProfile As Boolean, embeddedProfile As pdICCProfile
    imgHasIccProfile = GDIPlus_ImageGetProperty(hImage, GP_PT_ICCProfile, tmpPropHeader, tmpPropBuffer)
    
    If imgHasIccProfile Then
        
        'Create a temporary profile, and add it to PD's central color management cache
        Set embeddedProfile = New pdICCProfile
        embeddedProfile.LoadICCFromPtr tmpPropHeader.propLength, tmpPropHeader.propValue
        Erase tmpPropBuffer
        
        Dim colorProfileHash As String
        colorProfileHash = ColorManagement.AddProfileToCache(embeddedProfile)
        If (Not dstImage Is Nothing) Then dstImage.SetColorProfile_Original colorProfileHash
        
    End If
    
    'Next, pull an orientation flag, if any.  This is most relevant for JPEGs coming from a digital camera, but other
    ' formats (like TIFF) can also supply it.
    If UserPrefs.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then
        If AutoCorrectImageOrientation(hImage) Then PDDebug.LogAction "Image contains orientation data, and it was successfully handled."
    End If
    
    'Metafiles can contain brushes and other objects stored at extremely high DPIs.
    ' Limit these to 300 dpi to prevent OOM errors later on.
    If isMetafile Then GdipSetMetafileDownLevelRasterizationLimit hImage, 300
    
    'Retrieve the image's size
    ' RANDOM FACT! GdipGetImageDimension works fine on bitmaps.  On metafiles, it returns bizarre values
    ' that can be astronomically large.  My assumption is that these image dimensions are not necessarily
    ' returned in pixels (though pixels are the default for bitmaps), or perhaps they are meant to be
    ' adjusted at run-time by system DPI.  Regardless, the old floating-point width/height as now stored
    ' in the -F suffixed variables, while the original values now store Long-type copies of the image's
    ' initial dimensions.  (Also, these copies are good for both bitmaps and metafiles.)
    Dim imgWidth As Long, imgHeight As Long
    GdipGetImageWidth hImage, imgWidth
    GdipGetImageHeight hImage, imgHeight
    
    'Metafiles may want to use floating-point values, instead
    Dim imgWidthF As Single, imgHeightF As Single
    GdipGetImageDimension hImage, imgWidthF, imgHeightF
    
    'Retrieve the image's horizontal and vertical resolution (if any)
    Dim imgHResolution As Single, imgVResolution As Single
    GdipGetImageHorizontalResolution hImage, imgHResolution
    GdipGetImageVerticalResolution hImage, imgVResolution
    dstDIB.SetDPI imgHResolution, imgVResolution
    
    'pdDebug.LogAction "GDI+ image dimensions reported as: " & imgWidth & "x" & imgHeight & ", (" & imgWidthF & ", " & imgHeightF & ")"
    'pdDebug.LogAction "GDI+ image resolution reported as: " & imgHResolution & "x" & imgVResolution
    
    'Metafile containers (EMF, WMF) require special handling.
    Dim emfPlusAvailable As Boolean, metafileWasUpsampled As Boolean
    emfPlusAvailable = False
    metafileWasUpsampled = False
    
    If isMetafile Then
        
        'In a perfect world, we might do something like GIMP, and display an import dialog for metafiles.  This would allow the user to
        ' set an initial size for metafiles, taking advantage of their lossless rescalability before forcibly rasterizing them.
        
        'I don't want to implement this just yet, so instead, I'm simply aiming to report the same default size as MS Paint and Irfanview
        ' (which are the only programs I have that reliably load WMF and EMF files).
        
        'EMF dimensions are already reported identical to those programs, but WMF files are not.  The following code will make WMF sizes
        ' align with other software.
        If (imgFormatFIF = PDIF_WMF) Then
        
            'I assume 96 is used because it's the default DPI value in Windows.  I have not tested if different system DPI values affect
            ' the way GDI+ reports metafile size.
            If (imgHResolution <> 0!) Then imgWidth = imgWidth * CSng(96! / imgHResolution) Else imgHResolution = 96
            If (imgVResolution <> 0!) Then imgHeight = imgHeight * CSng(96! / imgVResolution) Else imgVResolution = 96
            
        End If
        
        'See if the incoming metafile already contains EMF+ records.  If it does, we don't want to up-convert it.
        ' (The GDI+ conversion function straight-up fails if the original contains mixed EMF/EMF+ data, unfortunately.)
        Dim tmpHeader As GP_MetafileHeader
        If (GdipGetMetafileHeaderFromMetafile(hImage, tmpHeader) = GP_OK) Then
            emfPlusAvailable = (tmpHeader.mfType = GP_MT_EmfPlus) Or (tmpHeader.mfType = GP_MT_EmfDual)
            If (tmpHeader.mfType = GP_MT_EmfPlus) Then PDDebug.LogAction "Note: incoming metafile is in EMF+ format.  Up-sampling will be skipped."
            If (tmpHeader.mfType = GP_MT_EmfDual) Then PDDebug.LogAction "Note: incoming metafile is in dual-EMF format.  EMF+ data will be preferentially used."
        End If
        
        'If GDI+ v1.1 is available, we can translate old-style EMFs and WMFs into the new GDI+ EMF+ format, which supports antialiasing
        ' and alpha channels (among other things).
        If ((Not emfPlusAvailable) And GDI_Plus.IsGDIPlusV11Available) Then
            
            PDDebug.LogAction "Incoming metafile detected.  Attempting to upsample to EMF+ format..."
            
            'Create a temporary GDI+ graphics object, whose properties will be used to control the render state of the EMF
            Dim tmpSettingsDIB As pdDIB
            Set tmpSettingsDIB = New pdDIB
            tmpSettingsDIB.CreateBlank 8, 8, 32, 0, 0
            
            Dim tmpGraphics As Long
            If (GdipCreateFromHDC(tmpSettingsDIB.GetDIBDC, tmpGraphics) = GP_OK) Then
                
                'Set high-quality antialiasing and interpolation
                GdipSetSmoothingMode tmpGraphics, GP_SM_Antialias
                GdipSetInterpolationMode tmpGraphics, GP_IM_HighQualityBicubic
                
                'Attempt to convert the EMF to EMF+ format
                Dim mfHandleDst As Long, convSuccess As Long
                
                'For reference: to write EMF+ data to file, use code like the following:
                'Dim newEmfPlusFileAndPath As String
                'newEmfPlusFileAndPath = Files.FileGetPath(srcFilename) & Files.FileGetName(srcFilename, True) & " (EMFPlus).emf"
                'If GdipConvertToEmfPlusToFile(tmpGraphics, hImage, convSuccess, StrPtr(newEmfPlusFileAndPath), EmfTypeEmfPlusOnly, 0, mfHandleDst) = 0 Then
                '
                'In PD, however, we want to perform the whole thing in-memory so we can immediately rasterize
                ' the result.
                Dim emfConvertResult As GP_Result
                emfConvertResult = GdipConvertToEmfPlus(tmpGraphics, hImage, convSuccess, GP_MT_EmfPlus, 0&, mfHandleDst)
                If (emfConvertResult = GP_OK) Then
                    
                    PDDebug.LogAction "EMF+ conversion successful!  Continuing with load..."
                    
                    'Conversion successful!  Replace our current image handle with the EMF+ copy
                    metafileWasUpsampled = True
                    emfPlusAvailable = True
                    GdipDisposeImage hImage
                    hImage = mfHandleDst

                Else
                    PDDebug.LogAction "EMF+ conversion failed (#" & CStr(emfConvertResult) & ", " & CStr(convSuccess) & ").  Original EMF data will be used."
                End If
                
                'Release our temporary graphics container
                GdipDeleteGraphics tmpGraphics
                
            End If
            
            'Release our temporary settings DIB
            Set tmpSettingsDIB = Nothing
        
        Else
            metafileWasUpsampled = emfPlusAvailable
        End If
        
    End If
    
    'Look for an alpha channel
    Dim imgHasAlpha As Boolean
    imgHasAlpha = False
    
    Dim imgPixelFormat As GP_PixelFormat
    GdipGetImagePixelFormat hImage, imgPixelFormat
    imgHasAlpha = ((imgPixelFormat And GP_PF_Alpha) <> 0)
    If (Not imgHasAlpha) Then imgHasAlpha = ((imgPixelFormat And GP_PF_PreMultAlpha) <> 0)
    
    'Make a note of the image's specific color depth, as relevant to PD
    Dim imgColorDepth As Long
    imgColorDepth = GetColorDepthFromPixelFormat(imgPixelFormat)
    
    'Check for CMYK images
    Dim isCMYK As Boolean
    isCMYK = ((imgPixelFormat And GP_PF_32bppCMYK) = GP_PF_32bppCMYK)
    If isCMYK Then PDDebug.LogAction "CMYK image found."
    
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, cTransform As pdLCMSTransform
    Dim hGraphics As Long, copyBitmapData As GP_BitmapData, tmpRect As RectL
    
    'We now want to split handling for metafiles vs bitmaps.
    If isMetafile Then
        
        'NOTE: this call *may* result in new imgWidth/imgHeight values!
        ' (The user can override embedded values via an import-time prompt.)
        GDIPlusLoadPicture = RasterizeMetafile(dstDIB, dstImage, imgWidth, imgHeight, metafileWasUpsampled, hImage, nonInteractiveMode, overrideParameters)
        
        'The user may have received a prompt to provide custom width/height values.  They are free to cancel
        ' this prompt.  When this happens, we need to provide safe cleanup.
        If (Not GDIPlusLoadPicture) Then
            GdipDisposeImage hImage
            Exit Function
        End If
        
    'This is a raster-type image
    Else
    
        'Create the destination DIB for this image.  Color-depth must be handled manually.
        If isCMYK Then
            dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 24
        Else
            dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 32, 0, 0
        End If
    
        'We now copy over image data in one of two ways.  Additional transforms may be required if
        ' the source image is in an unexpected color format (e.g. CMYK).
        If imgHasAlpha Then
            
            'We are now going to copy the image's data directly into our destination DIB by using LockBits.
            ' Very fast, and not much code!
            
            'Start by preparing a BitmapData variable with instructions on where GDI+ should paste the bitmap data
            With copyBitmapData
                .BD_Width = imgWidth
                .BD_Height = imgHeight
                .BD_PixelFormat = GP_PF_32bppPARGB
                .BD_Stride = dstDIB.GetDIBStride
                .BD_Scan0 = dstDIB.GetDIBPointer
            End With
            
            'Next, prepare a clipping rect
            With tmpRect
                .Left = 0
                .Top = 0
                .Right = imgWidth
                .Bottom = imgHeight
            End With
            
            'Use LockBits to perform the copy for us.
            GdipBitmapLockBits hImage, tmpRect, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppPARGB, copyBitmapData
            GdipBitmapUnlockBits hImage, copyBitmapData
        
        'The image does *not* have alpha
        Else
            
            'CMYK is handled separately from regular RGB data, as we want to perform an ICC profile conversion as well.
            ' Note that if a CMYK profile is not present, we allow GDI+ to convert the image to RGB for us.
            If (isCMYK And imgHasIccProfile) Then
            
                'Create a blank 32bpp DIB, which will hold the CMYK data
                Dim tmpCMYKDIB As pdDIB
                Set tmpCMYKDIB = New pdDIB
                tmpCMYKDIB.CreateBlank imgWidth, imgHeight, 32
            
                'Next, prepare a BitmapData variable with instructions on where GDI+ should paste the bitmap data
                With copyBitmapData
                    .BD_Width = imgWidth
                    .BD_Height = imgHeight
                    .BD_PixelFormat = GP_PF_32bppCMYK
                    .BD_Stride = tmpCMYKDIB.GetDIBStride
                    .BD_Scan0 = tmpCMYKDIB.GetDIBPointer
                End With
                
                'Next, prepare a clipping rect
                With tmpRect
                    .Left = 0
                    .Top = 0
                    .Right = imgWidth
                    .Bottom = imgHeight
                End With
                
                'Use LockBits to perform the copy for us.
                GdipBitmapLockBits hImage, tmpRect, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppCMYK, copyBitmapData
                GdipBitmapUnlockBits hImage, copyBitmapData
                
                'We now need to apply the CMYK transform.  This is a multistep process that has been condensed here due to
                ' its rarity in the actual processing chain.
                Dim cmSuccessful As Boolean
                cmSuccessful = False
                
                Set srcProfile = New pdLCMSProfile
                Set dstProfile = New pdLCMSProfile
                
                If srcProfile.CreateFromPDICCObject(embeddedProfile) Then
                    If dstProfile.CreateSRGBProfile() Then
                        
                        Set cTransform = New pdLCMSTransform
                        If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, TYPE_CMYK_8, TYPE_BGR_8, INTENT_PERCEPTUAL) Then
                            
                            Set srcProfile = Nothing: Set dstProfile = Nothing
                            cmSuccessful = cTransform.ApplyTransformToArbitraryMemory(copyBitmapData.BD_Scan0, dstDIB.GetDIBScanline(0), copyBitmapData.BD_Stride, dstDIB.GetDIBStride, dstDIB.GetDIBHeight, dstDIB.GetDIBWidth, False)
                                
                            If cmSuccessful Then
                                PDDebug.LogAction "Copying newly transformed sRGB data..."
                                dstDIB.SetColorManagementState cms_ProfileConverted
                                dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                                dstDIB.SetInitialAlphaPremultiplicationState True
                            End If
                            
                            Set cTransform = Nothing
                            
                        End If
                    End If
                End If
                
                'Check for potential failure states, and fall back to a naive CMYK transform as necessary.
                If (Not cmSuccessful) Then
                    
                    PDDebug.LogAction "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
                    
                    GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
                    If (hGraphics <> 0) Then
                        GdipDrawImageRect hGraphics, hImage, 0, 0, imgWidth, imgHeight
                        GdipDeleteGraphics hGraphics
                    End If
                    
                End If
                
                Set tmpCMYKDIB = Nothing
            
            Else
                
                'Render the GDI+ image directly onto the newly created DIB
                GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
                If (hGraphics <> 0) Then
                    GdipDrawImageRect hGraphics, hImage, 0, 0, imgWidth, imgHeight
                    GdipDeleteGraphics hGraphics
                End If
                
            End If
        
        'End raster w/ alpha vs raster w/out alpha check
        End If
    
    'End metafile vs raster check
    End If
    
    'Note some original file settings inside the DIB
    dstDIB.SetOriginalFormat imgFormatFIF
    dstDIB.SetOriginalColorDepth imgColorDepth
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'Release any remaining GDI+ handles.  IMPORTANTLY, if this is a multipage TIFF, we *skip* this step,
    ' as we'll be reusing the image handle on subsequent pages.
    If isMultiPage Then
        m_hMultiPageImage = hImage
        numOfPages = imgPageCount
    Else
        If (hImage <> 0) Then GdipDisposeImage hImage
        hImage = 0
    End If
    
    'Before exiting, check for an embedded color profile.  If the image had one, we want to apply it to the
    ' destination image now, if we haven't already.  (Only CMYK images will have been processed already.)
    If (Not isCMYK) And imgHasIccProfile And ColorManagement.UseEmbeddedICCProfiles() Then
        
        PDDebug.LogAction "Applying color management to GDI+ image..."
        
        Set srcProfile = New pdLCMSProfile
        Set dstProfile = New pdLCMSProfile
        
        Dim srcFormat As LCMS_PIXEL_FORMAT
        
        'Test the embedded profile for basic validity
        If srcProfile.CreateFromPDICCObject(embeddedProfile) Then
            
            'Look for grayscale profiles.  If the source image was grayscale, it will have already been converted
            ' to RGB/A - but if it contained an ICC profile in grayscale mode (alongside the gray data), we need
            ' to apply special handling, because LittleCMS does not support applying gray-mode profiles to RGB pixels.
            If srcProfile.IsGrayProfile Then
                
                'Ensure alpha channel existence
                If (dstDIB.GetDIBColorDepth = 24) Then dstDIB.ConvertTo32bpp
                dstDIB.SetAlphaPremultiplication False
                
                'Make a temporary copy of the base image in 8-bpp grayscale
                Dim tmpBytes() As Byte
                DIBs.GetDIBGrayscaleMap dstDIB, tmpBytes, False
                
                'Create a destination linear gray profile
                If dstProfile.CreateGenericGrayscaleProfile() Then
                    
                    'Convert source gray bytes to linear gray bytes
                    srcFormat = TYPE_GRAY_8
                    
                    Set cTransform = New pdLCMSTransform
                    If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcFormat, srcFormat, INTENT_PERCEPTUAL) Then
                        
                        cmSuccessful = cTransform.ApplyTransformToArbitraryMemory(VarPtr(tmpBytes(0, 0)), VarPtr(tmpBytes(0, 0)), dstDIB.GetDIBWidth, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, dstDIB.GetDIBWidth, False)
                        
                        'On a successful transform, copy the newly transformed gray data back into a standard RGBA object
                        If cmSuccessful Then
                            
                            PDDebug.LogAction "Copying newly transformed gray data..."
                            
                            'Alpha is already premultiplied because we assume a fully opaque image (JPEGs don't support alpha)
                            DIBs.CreateDIBFromGrayscaleMap dstDIB, tmpBytes, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight
                            dstDIB.SetColorManagementState cms_ProfileConverted
                            dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                            
                        End If
                        
                        Set cTransform = Nothing
                    
                    '/Failed to create gray transform; source transform is likely bugged in some way
                    End If
                    
                End If
                
                'Free local grayscale copy of image
                Erase tmpBytes
                
            'Source profile is non-grayscale, so we can proceed normally
            Else
                
                'Ensure we can create a successful destination sRGB profile (failsafe check only)
                If dstProfile.CreateSRGBProfile() Then
                    
                    If (dstDIB.GetDIBColorDepth = 24) Then srcFormat = TYPE_BGR_8 Else srcFormat = TYPE_BGRA_8
                    
                    Set cTransform = New pdLCMSTransform
                    If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcFormat, srcFormat, INTENT_PERCEPTUAL) Then
                        
                        '32-bpp images need to be unpremultiplied, but if the source image didn't contain any alpha
                        ' (e.g. JPEGs), then this step is pointless as the target image has all-255 for its alpha channel.
                        If (dstDIB.GetAlphaPremultiplication And imgHasAlpha) Then dstDIB.SetAlphaPremultiplication False
                        
                        Set srcProfile = Nothing: Set dstProfile = Nothing
                        cmSuccessful = cTransform.ApplyTransformToPDDib(dstDIB)
                        
                        If cmSuccessful Then
                            PDDebug.LogAction "Copying newly transformed sRGB data..."
                            dstDIB.SetColorManagementState cms_ProfileConverted
                            dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                            If (Not dstDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
                        End If
                        
                        Set cTransform = Nothing
                        
                    Else
                        PDDebug.LogAction "WARNING!  Image could not be color-managed; color space mismatch is a likely explanation."
                    End If
                
                '/created dst color profile successfully
                End If
            
            '/source profile is gray vs color
            End If
        
        '/source profile object created successfully
        End If
    
    '/source image has embedded color profile
    End If
    
    'Return success!
    GDIPlusLoadPicture = True
    
End Function

'Used for rendering previews in the "import metafile and optionally choose custom dimensions" dialog
Public Function PaintMetafileToArbitraryDIB(ByRef dstDIB As pdDIB, ByVal hGdipImage As Long, Optional ByVal dstX As Single = 0!, Optional ByVal dstY As Single = 0!) As Boolean
    
    PaintMetafileToArbitraryDIB = False
    Dim hGraphics As Long
    
    If (Not dstDIB Is Nothing) And (hGdipImage <> 0) Then
        GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
        GdipDrawImageRect hGraphics, hGdipImage, dstX, dstY, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight
        GdipDeleteGraphics hGraphics
        PaintMetafileToArbitraryDIB = True
    End If
    
End Function

'Rasterize a passed metadata handle, with optional support for prompting the user for custom dimensions
Private Function RasterizeMetafile(ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, ByRef intWidth As Long, ByRef intHeight As Long, ByVal metafileIsEmfPlus As Boolean, ByVal hImage As Long, Optional ByVal nonInteractiveMode As Boolean = False, Optional ByVal overrideParameters As String = vbNullString) As Boolean
    
    RasterizeMetafile = False
    
    'Check for non-interactive mode.  The user can specify this manually, *or* we can auto-infer it
    ' from PD's global batch state tracker.
    If (Not nonInteractiveMode) Then
        nonInteractiveMode = (Macros.GetMacroStatus() = MacroBATCH) Or (Macros.GetMacroStatus() = MacroPLAYBACK)
    End If
    
    'We now know the size the metafile is *supposed* to be, but the most useful thing about vector graphics
    ' is the ability to losslessly (ish) resize them to arbitrary sizes.  If the user allows, raise a prompt
    ' to ask the user what size they want us to use for this image.
    Dim userWidth As Long, userHeight As Long
    
    'In non-interactive mode, rely on the embedded SVG size parameter, *or* any optional overrides supplied via
    ' optional param string.
    If nonInteractiveMode Then
        
        'Default to embedded size
        userWidth = intWidth
        userHeight = intHeight
        
        'Look for user overrides
        Dim cOverrideParams As pdSerialize
        Set cOverrideParams = New pdSerialize
        cOverrideParams.SetParamString overrideParameters
        
        'Override default size with user-supplied values
        If Not cOverrideParams.GetBool("vector-size-use-default", True, True) Then
            userWidth = cOverrideParams.GetLong("vector-size-x", 0, True)
            If (userWidth <= 0) Or (userWidth > 32000) Then userWidth = intWidth
            userHeight = cOverrideParams.GetLong("vector-size-y", 0, True)
            If (userHeight <= 0) Or (userHeight > 32000) Then userHeight = intHeight
        End If
        
        RasterizeMetafile = True
        
    'UI prompt allowed
    Else
        
        Dim userInput As VbMsgBoxResult, userDPI As Long
        userInput = Dialogs.PromptImportEMF(hImage, intWidth, intHeight, userWidth, userHeight, userDPI)
        
        If (userInput = vbOK) Then
            
            'Validate user width/height
            If (userWidth < 1) Then userWidth = intWidth
            If (userHeight < 1) Then userHeight = intHeight
            
            'Cache DPI inside the parent pdImage object
            If (Not dstImage Is Nothing) Then dstImage.SetDPI userDPI, userDPI
            
            RasterizeMetafile = True
            
        Else
            'Do nothing; case is handled by default
        End If
    
    End If
    
    If RasterizeMetafile Then
        
        'Replace incoming width/height with the user's overrides (if any)
        intWidth = userWidth
        intHeight = userHeight
        
        'Metafiles can be rendered to transparent surfaces, but they're not always predictable.  GDI commands in particular
        ' can sometimes result in portions of the metafile being transparent for no obvious reason.  (I've got a number of
        ' problematic examples in the "images from testers/metafiles" folder.)
        '
        'In the future, a dialog could be presented, but right now, we use very simple heuristics to determine how to proceed.
        ' 1) If we converted the metafile to EMF+ ourselves, the image is 99+% guaranteed to be okay with transparency.
        '    Paint it to a 32-bpp surface.
        ' 2) If the EMF+ data came from somewhere else, the data may handle transparency unpredictably.
        '    Paint to a 24-bpp base just to be safe.
        If metafileIsEmfPlus Then
            dstDIB.CreateBlank intWidth, intHeight, 32, 0, 0
        Else
            dstDIB.CreateBlank intWidth, intHeight, 24, vbWhite
        End If
        
        'Render the GDI+ image directly onto the newly created DIB
        Dim hGraphics As Long
        GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
        GdipDrawImageRect hGraphics, hImage, 0!, 0!, intWidth, intHeight
        GdipDeleteGraphics hGraphics
        
    End If
        
End Function

'Returns TRUE if the source image...
' 1) contains orientation data, and...
' 2) said orientation data is *not* "standard orientation", and...
' 3) we successfully apply said orientation data to the underlying image
'
'If you don't want orientation data applied, *don't call this function*!
Private Function AutoCorrectImageOrientation(ByVal hImage As Long) As Boolean
    
    Dim tmpPropHeader As GP_PropertyItem, tmpPropBuffer() As Byte
    If GDIPlus_ImageGetProperty(hImage, GP_PT_Orientation, tmpPropHeader, tmpPropBuffer) Then
        
        'The returned buffer should only ever be two bytes, as this property is an integer.
        If (tmpPropHeader.propLength = 2) Then
            
            'Select based on the MSB
            Select Case tmpPropBuffer(0)
                
                'Standard orientation - ignore!
                Case 1
            
                'The 0th row is at the visual top of the image, and the 0th column is the visual right-hand side
                Case 2
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_NoneFlipX) = GP_OK)
                    
                'The 0th row is at the visual bottom of the image, and the 0th column is the visual right-hand side
                Case 3
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_180FlipNone) = GP_OK)
                
                'The 0th row is at the visual bottom of the image, and the 0th column is the visual left-hand side
                Case 4
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_NoneFlipY) = GP_OK)
                
                'The 0th row is the visual left-hand side of of the image, and the 0th column is the visual top
                Case 5
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_270FlipY) = GP_OK)
                
                'The 0th row is the visual right -hand side of of the image, and the 0th column is the visual top
                Case 6
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_90FlipNone) = GP_OK)
                    
                'The 0th row is the visual right -hand side of of the image, and the 0th column is the visual bottom
                Case 7
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_90FlipY) = GP_OK)
                    
                'The 0th row is the visual left-hand side of of the image, and the 0th column is the visual bottom
                Case 8
                    AutoCorrectImageOrientation = (GdipImageRotateFlip(hImage, GP_RF_270FlipNone) = GP_OK)
            
            End Select
        
        End If
    
    End If
        
End Function

'After calling GDIPlusLoadPicture and discovering that your file is multi-frame (GIF) or multi-page (TIFF),
' you can call this function to continue loading subsequent frames/pages into the active image.  Note that
' PD will *always* call this function on multi-page images, and this is important because the saved image
' handle would leak if you encountered a multipage image but *didn't* call this function after.
Public Function ContinueLoadingMultipageImage(ByRef srcFilename As String, ByRef dstDIB As pdDIB, Optional ByVal numOfPages As Long = 0, Optional ByVal showMessages As Boolean = True, Optional ByRef targetImage As pdImage = Nothing, Optional ByVal suppressDebugData As Boolean = False, Optional ByVal suggestedFilename As String = vbNullString) As Boolean
    
    ContinueLoadingMultipageImage = False
    
    'Ensure we maintained a handle to the image in question
    If (m_hMultiPageImage <> 0) Then
        
        'Failsafe check to ensure the destination DIB exists
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        
        'GDI+ uses GUIDs to access frame parameters; for TIFF, these are defined by "page", for GIF, by "time".
        ' To simplify our code, we'll store the GUID we need in a format-agnostic struct.
        Dim frameDimensionID(0 To 15) As Byte
        If (m_OriginalFIF = PDIF_TIFF) Then
            CLSIDFromString StrPtr(GP_FD_Page), VarPtr(frameDimensionID(0))
        ElseIf (m_OriginalFIF = PDIF_GIF) Then
            CLSIDFromString StrPtr(GP_FD_Time), VarPtr(frameDimensionID(0))
        End If
        
        'Ensure the passed page count is accurate.  (The verification code varies by file type.)
        Dim imgPageCount As Long
        If (m_OriginalFIF = PDIF_TIFF) Then
            If (GdipImageGetFrameCount(m_hMultiPageImage, VarPtr(frameDimensionID(0)), imgPageCount) <> GP_OK) Then InternalGDIPlusError "ContinueLoadingMultipageImage failed to retrieve page count."
        ElseIf (m_OriginalFIF = PDIF_GIF) Then
            Dim numFrameDimensions As Long
            If (GdipImageGetFrameDimensionsCount(m_hMultiPageImage, numFrameDimensions) = GP_OK) Then
                If (numFrameDimensions > 0) Then
                    Dim lstGuids() As Byte
                    ReDim lstGuids(0 To 16 * numFrameDimensions - 1) As Byte
                    If (GdipImageGetFrameDimensionsList(m_hMultiPageImage, VarPtr(lstGuids(0)), numFrameDimensions) = GP_OK) Then
                        If (GdipImageGetFrameCount(m_hMultiPageImage, VarPtr(lstGuids(0)), imgPageCount) <> GP_OK) Then InternalGDIPlusError "ContinueLoadingMultipageImage failed to retrieve page count."
                    End If
                End If
            End If
        End If
        
        If (imgPageCount <> numOfPages) Then
            InternalGDIPlusError "ContinueLoadingMultipageImage passed bad page numbers", "reported page count differs (" & CStr(numOfPages) & " vs " & CStr(imgPageCount) & ")"
            Exit Function
        End If
        
        'To correctly assemble a correct image (particularly in the case of animated GIFs), we need to cache
        ' a *lot* of per-frame data.  TIFFs have their own complications because each frame can be in a different
        ' color format, including ones that require color-management like CMYK.  PD attempts to cover all
        ' possible cases in a proper color-managed fashion, regardless of incoming format.
        Dim imgWidth As Long, imgHeight As Long, imgHResolution As Single, imgVResolution As Single
        Dim imgPixelFormat As GP_PixelFormat, imgHasAlpha As Boolean, isCMYK As Boolean
        Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, cTransform As pdLCMSTransform
        Dim hGraphics As Long, copyBitmapData As GP_BitmapData, tmpRect As RectL
        Dim tmpPropHeader As GP_PropertyItem, tmpPropBuffer() As Byte
        Dim imgHasIccProfile As Boolean, embeddedProfile As pdICCProfile
        
        'If the multipage handle is valid, and the image is a GIF, retrieve some animation-specific metadata
        ' (like loop count) and cache it inside the parent object.
        If (m_OriginalFIF = PDIF_GIF) Then
            If GDIPlus_ImageGetProperty(m_hMultiPageImage, GP_PT_LoopCount, tmpPropHeader, tmpPropBuffer) Then
            
                'The returned value is a ushort; copy it into a signed long
                Dim tmpLoopCount As Long
                CopyMemoryStrict VarPtr(tmpLoopCount), VarPtr(tmpPropBuffer(0)), 2
                If (Not targetImage Is Nothing) Then targetImage.ImgStorage.AddEntry "animation-loop-count", Trim$(Str$(tmpLoopCount))
                
            End If
        End If
        
        'TIFFs are messy because individual frames can have rotation tags that require a post-load transform.
        ' Grab the user's preference for this in advance, so we don't have to travel out to the preferences
        ' file on each frame.
        Dim autoRotatePref As Boolean
        autoRotatePref = UserPrefs.GetPref_Boolean("Loading", "ExifAutoRotate", True)
        
        'Time to start iterating pages.  The first page was already loaded in a previous step, so we start
        ' at index 1 (indices are always 0-based).
        Dim pageToLoad As Long
        For pageToLoad = 1 To numOfPages - 1
            
            'If the image is large, it's nice to provide status updates to the user - we try to do this every 8-ish frames
            Message "Loading page %1 of %2...", CStr(pageToLoad + 1), numOfPages, "DONOTLOG"
            If ((pageToLoad And 7) = 0) Then VBHacks.DoEvents_SingleHwnd FormMain.hWnd
            
            'Notify GDI+ of the frame/page we want to select.  Note that performance is better when
            ' accessing frames in sequence, as GIF frame appearance can depend on the contents of previous
            ' frames (so accessing out-of-order requires more work on GDI+'s part).
            If (GdipImageSelectActiveFrame(m_hMultiPageImage, VarPtr(frameDimensionID(0)), pageToLoad) = GP_OK) Then
                
                'Throughout this process, we'll be notifying the destination DIB of various page parameters.
                dstDIB.SetOriginalFormat m_OriginalFIF
                
                'Retrieve this frame's size
                GdipGetImageWidth m_hMultiPageImage, imgWidth
                GdipGetImageHeight m_hMultiPageImage, imgHeight
                
                'PDDebug.LogAction "Loading page with dimensions (" & CStr(imgWidth) & "x" & CStr(imgHeight) & ")"
                
                'Retrieve this frame's horizontal and vertical resolution (if any)
                GdipGetImageHorizontalResolution m_hMultiPageImage, imgHResolution
                GdipGetImageVerticalResolution m_hMultiPageImage, imgVResolution
                dstDIB.SetDPI imgHResolution, imgVResolution
                
                'Look for an alpha channel
                GdipGetImagePixelFormat m_hMultiPageImage, imgPixelFormat
                imgHasAlpha = ((imgPixelFormat And GP_PF_Alpha) <> 0)
                If (Not imgHasAlpha) Then imgHasAlpha = ((imgPixelFormat And GP_PF_PreMultAlpha) <> 0)
                
                'Check for CMYK pages
                isCMYK = ((imgPixelFormat And GP_PF_32bppCMYK) = GP_PF_32bppCMYK)
                If isCMYK Then PDDebug.LogAction "CMYK page found."
                
                'Make a note of the image's specific color depth, as relevant to PD
                dstDIB.SetOriginalColorDepth GetColorDepthFromPixelFormat(imgPixelFormat)
                
                'Look for an ICC profile and cache the result.  (Again, this is a TIFF issue; each page
                ' could theoretically have its own ICC profile, ugh.)
                imgHasIccProfile = GDIPlus_ImageGetProperty(m_hMultiPageImage, GP_PT_ICCProfile, tmpPropHeader, tmpPropBuffer)
                If imgHasIccProfile Then
                    Set embeddedProfile = New pdICCProfile
                    embeddedProfile.LoadICCFromPtr tmpPropHeader.propLength, tmpPropHeader.propValue
                    Erase tmpPropBuffer
                    
                    'TODO: when we eventually get around to full color-management, we'll want to convert subsequent
                    ' layers to the same color space as the base layer.  For now, however, we just convert to sRGB.
                    
                End If
                
                'Next, pull an orientation flag, if any, and apply it to the underlying page.
                ' (Note that we can skip this step for GIFs; they don't support rotation tags.)
                If autoRotatePref And (m_OriginalFIF <> PDIF_GIF) Then
                    If AutoCorrectImageOrientation(m_hMultiPageImage) Then PDDebug.LogAction "Image contains orientation data, and it was successfully handled."
                End If
                
                'We now need to copy the relevant image bytes into the destination DIB.  This is complicated, unfortunately.
                
                'Create the destination DIB for this image.  CMYK requires special handling.
                If isCMYK Then
                    dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 24
                Else
                    dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 32, 0, 0
                End If
                
                'We now copy over image data in one of two ways.  Additional transforms may be required if
                ' the source image is in an unexpected color format (e.g. CMYK).
                If imgHasAlpha Then
                    
                    'We are now going to copy the image's data directly into our destination DIB by using LockBits.
                    ' Very fast, and not much code!
                    
                    'Start by preparing a BitmapData variable with instructions on where GDI+ should paste the bitmap data
                    With copyBitmapData
                        .BD_Width = imgWidth
                        .BD_Height = imgHeight
                        .BD_PixelFormat = GP_PF_32bppPARGB
                        .BD_Stride = dstDIB.GetDIBStride
                        .BD_Scan0 = dstDIB.GetDIBPointer
                    End With
                    
                    'Next, prepare a clipping rect
                    With tmpRect
                        .Left = 0
                        .Top = 0
                        .Right = imgWidth
                        .Bottom = imgHeight
                    End With
                    
                    'Use LockBits to perform the copy for us.
                    GdipBitmapLockBits m_hMultiPageImage, tmpRect, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppPARGB, copyBitmapData
                    GdipBitmapUnlockBits m_hMultiPageImage, copyBitmapData
                    
                'Image does *not* have alpha
                Else
                    
                    'CMYK is handled separately from regular RGB data, as we want to perform an ICC profile conversion as well.
                    ' Note that if a CMYK profile is not present, we allow GDI+ to convert the image to RGB for us.
                    If (isCMYK And imgHasIccProfile) Then
                    
                        'Create a blank 32bpp DIB, which will hold the CMYK data
                        Dim tmpCMYKDIB As pdDIB
                        Set tmpCMYKDIB = New pdDIB
                        tmpCMYKDIB.CreateBlank imgWidth, imgHeight, 32
                    
                        'Next, prepare a BitmapData variable with instructions on where GDI+ should paste the bitmap data
                        With copyBitmapData
                            .BD_Width = imgWidth
                            .BD_Height = imgHeight
                            .BD_PixelFormat = GP_PF_32bppCMYK
                            .BD_Stride = tmpCMYKDIB.GetDIBStride
                            .BD_Scan0 = tmpCMYKDIB.GetDIBPointer
                        End With
                        
                        'Next, prepare a clipping rect
                        With tmpRect
                            .Left = 0
                            .Top = 0
                            .Right = imgWidth
                            .Bottom = imgHeight
                        End With
                        
                        'Use LockBits to perform the copy for us.
                        GdipBitmapLockBits m_hMultiPageImage, tmpRect, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppCMYK, copyBitmapData
                        GdipBitmapUnlockBits m_hMultiPageImage, copyBitmapData
                        
                        'We now need to apply the CMYK transform.  This is a multistep process that has been condensed here due to
                        ' its rarity in the actual processing chain.
                        Dim cmSuccessful As Boolean
                        cmSuccessful = False

                        Set srcProfile = New pdLCMSProfile
                        Set dstProfile = New pdLCMSProfile

                        If srcProfile.CreateFromPDICCObject(embeddedProfile) Then
                            If dstProfile.CreateSRGBProfile() Then

                                Set cTransform = New pdLCMSTransform
                                If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, TYPE_CMYK_8, TYPE_BGR_8, INTENT_PERCEPTUAL) Then

                                    Set srcProfile = Nothing: Set dstProfile = Nothing
                                    cmSuccessful = cTransform.ApplyTransformToArbitraryMemory(copyBitmapData.BD_Scan0, dstDIB.GetDIBScanline(0), copyBitmapData.BD_Stride, dstDIB.GetDIBStride, dstDIB.GetDIBHeight, dstDIB.GetDIBWidth, False)

                                    If cmSuccessful Then
                                        PDDebug.LogAction "Copying newly transformed sRGB data..."
                                        dstDIB.SetColorManagementState cms_ProfileConverted
                                        dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                                        dstDIB.SetInitialAlphaPremultiplicationState True
                                    End If

                                    Set cTransform = Nothing

                                End If
                            End If
                        End If
                        
                        'Check for potential failure states, and fall back to a naive CMYK transform as necessary.
                        If (Not cmSuccessful) Then
                            
                            PDDebug.LogAction "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
                            
                            GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
                            If (hGraphics <> 0) Then
                                GdipDrawImageRect hGraphics, m_hMultiPageImage, 0, 0, imgWidth, imgHeight
                                GdipDeleteGraphics hGraphics
                            End If
                            
                        End If
                        
                        Set tmpCMYKDIB = Nothing
                    
                    Else
                        
                        'Render the GDI+ image directly onto the newly created DIB
                        GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
                        GdipSetCompositingMode hGraphics, GP_CM_SourceCopy
                        If (hGraphics <> 0) Then
                            GdipDrawImageRect hGraphics, m_hMultiPageImage, 0, 0, imgWidth, imgHeight
                            GdipDeleteGraphics hGraphics
                        End If
                        
                    End If
                
                'End alpha vs no-alpha check
                End If
                
                'The destination DIB now contains a full copy of the page data!  Convert it to 32-bpp as necessary.
                dstDIB.SetInitialAlphaPremultiplicationState True
                ImageImporter.ForceTo32bppMode dstDIB
                
                'If this page has an embedded color profile, we want to apply it to the destination image now,
                ' if we haven't already.  (Only CMYK images will have been processed already.)
                If (Not isCMYK) And imgHasIccProfile Then
                    
                    PDDebug.LogAction "Applying color management to current page..."
                    
                    Set srcProfile = New pdLCMSProfile
                    Set dstProfile = New pdLCMSProfile
                    
                    If srcProfile.CreateFromPDICCObject(embeddedProfile) Then
                        If dstProfile.CreateSRGBProfile() Then
                            
                            Dim srcFormat As LCMS_PIXEL_FORMAT
                            If (dstDIB.GetDIBColorDepth = 24) Then srcFormat = TYPE_BGR_8 Else srcFormat = TYPE_BGRA_8
                            
                            Set cTransform = New pdLCMSTransform
                            If cTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcFormat, srcFormat, INTENT_PERCEPTUAL) Then
                                
                                Set srcProfile = Nothing: Set dstProfile = Nothing
                                If dstDIB.GetAlphaPremultiplication Then dstDIB.SetAlphaPremultiplication False
                                cmSuccessful = cTransform.ApplyTransformToPDDib(dstDIB)
                                
                                If cmSuccessful Then
                                    PDDebug.LogAction "Copying newly transformed sRGB data..."
                                    dstDIB.SetColorManagementState cms_ProfileConverted
                                    dstDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                                    dstDIB.SetAlphaPremultiplication True
                                End If
                                
                                Set cTransform = Nothing
                                
                            Else
                                PDDebug.LogAction "WARNING!  Image could not be color-managed; color space mismatch is a likely explanation."
                            End If
                        End If
                    End If
                    
                End If
                
                'Create a blank layer in the receiving image, then copy our finished DIB into it
                Dim newLayerID As Long, newLayerName As String
                newLayerID = targetImage.CreateBlankLayer
                newLayerName = Layers.GenerateInitialLayerName(vbNullString, suggestedFilename, True, targetImage, dstDIB, pageToLoad)
                targetImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, dstDIB, True
            
            '/bad select frame call
            Else
                PDDebug.LogAction "WARNING!  Failed to set active page #" & pageToLoad
            End If
        
        Next pageToLoad
        
        'For animated GIFs, we now want to assign frame times to each frame.  We store the frame time as
        ' an absolute value inside each layer, and we also append frame time to each layer's name.
        If (m_OriginalFIF = PDIF_GIF) And (m_FrameCount > 0) Then
            
            'Failsafe check for success on original frame time retrieval
            If (UBound(m_FrameTimes) = m_FrameCount - 1) Then
                
                'Add frame count to the target image (this is not currently used)
                targetImage.ImgStorage.AddEntry "agif-frame-count", m_FrameCount
                
                Dim cFrame As Long
                For cFrame = 0 To m_FrameCount - 1
                    targetImage.GetLayerByIndex(cFrame).SetLayerFrameTimeInMS m_FrameTimes(cFrame)
                    targetImage.GetLayerByIndex(cFrame).SetLayerName targetImage.GetLayerByIndex(cFrame).GetLayerName & " (" & CStr(m_FrameTimes(cFrame)) & "ms)"
                Next cFrame
                
            End If
            
        End If
        
        'Before exiting, make sure we free the parent image handle
        GDI_Plus.ReleaseGDIPlusImage m_hMultiPageImage
        
        ContinueLoadingMultipageImage = True
        
    Else
        InternalGDIPlusError "ContinueLoadingMultipageImage failed", "multipage handle was prematurely closed"
    End If
    
End Function

'Given a GDI+ pixel format value, return a numeric color depth (e.g. 24, 32, etc)
Private Function GetColorDepthFromPixelFormat(ByVal gdipPixelFormat As GP_PixelFormat) As Long

    If (gdipPixelFormat = GP_PF_1bppIndexed) Then
        GetColorDepthFromPixelFormat = 1
    ElseIf (gdipPixelFormat = GP_PF_4bppIndexed) Then
        GetColorDepthFromPixelFormat = 4
    ElseIf (gdipPixelFormat = GP_PF_8bppIndexed) Then
        GetColorDepthFromPixelFormat = 8
    ElseIf (gdipPixelFormat = GP_PF_16bppGreyscale) Or (gdipPixelFormat = GP_PF_16bppRGB555) Or (gdipPixelFormat = GP_PF_16bppRGB565) Or (gdipPixelFormat = GP_PF_16bppARGB1555) Then
        GetColorDepthFromPixelFormat = 16
    ElseIf (gdipPixelFormat = GP_PF_24bppRGB) Or (gdipPixelFormat = GP_PF_32bppRGB) Then
        GetColorDepthFromPixelFormat = 24
    ElseIf (gdipPixelFormat = GP_PF_32bppARGB) Or (gdipPixelFormat = GP_PF_32bppPARGB) Then
        GetColorDepthFromPixelFormat = 32
    ElseIf (gdipPixelFormat = GP_PF_48bppRGB) Then
        GetColorDepthFromPixelFormat = 48
    ElseIf (gdipPixelFormat = GP_PF_64bppARGB) Or (gdipPixelFormat = GP_PF_64bppPARGB) Then
        GetColorDepthFromPixelFormat = 64
    Else
        GetColorDepthFromPixelFormat = 24
    End If

End Function

'Save an image using GDI+.  Per the current save spec, ImageID must be specified.
' Additional save options are currently available for JPEGs (save quality, range [1,100]) and TIFFs (compression type).
Public Function GDIPlusSavePicture(ByRef srcPDImage As pdImage, ByVal dstFilename As String, ByVal imgFormat As PD_2D_FileFormatExport, ByVal outputColorDepth As Long, Optional ByVal jpegQuality As Long = 92) As Boolean
    
    Message "Saving..."
    
    On Error GoTo GDIPlusSaveError
    PDDebug.LogAction "Prepping image for GDI+ export..."
    
    'If the output format is 24bpp (e.g. JPEG) but the input image is 32bpp, composite it against white
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    If (tmpDIB.GetDIBColorDepth <> 24) And imgFormat = P2_FFE_JPEG Then tmpDIB.CompositeBackgroundColor 255, 255, 255
    
    'Create a GDI+ bitmap handle for the source image
    PDDebug.LogAction "Creating GDI+ image copy..."
        
    Dim hImage As Long
    If (Not GetGdipBitmapHandleFromDIB(hImage, tmpDIB)) Then
        GDI_Plus.ReleaseGDIPlusImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    'Certain image formats require extra parameters, and because the values are passed ByRef, they can't be constants
    Dim gdipColorDepth As Long
    gdipColorDepth = outputColorDepth
    
    Dim tiff_Compression As GP_EncoderValue
    tiff_Compression = GP_EV_CompressionLZW
    
    'TIFF has some unique constraints on account of its many compression schemes.  Because it only supports a subset
    ' of compression types, we must adjust our code accordingly.
    If (imgFormat = P2_FFE_TIFF) Then
    
        Select Case UserPrefs.GetPref_Long("File Formats", "TIFF Compression", 0)
        
            'Default settings (LZW for > 1bpp, CCITT Group 4 fax encoding for 1bpp)
            Case 0
                If (gdipColorDepth = 1) Then tiff_Compression = GP_EV_CompressionCCITT4 Else tiff_Compression = GP_EV_CompressionLZW
                
            'No compression
            Case 1
                tiff_Compression = GP_EV_CompressionNone
                
            'Macintosh Packbits (RLE)
            Case 2
                tiff_Compression = GP_EV_CompressionRle
            
            'Proper deflate (Adobe-style) - not supported by GDI+
            Case 3
                tiff_Compression = GP_EV_CompressionLZW
            
            'Obsolete deflate - not supported by GDI+
            Case 4
                tiff_Compression = GP_EV_CompressionLZW
            
            'LZW
            Case 5
                tiff_Compression = GP_EV_CompressionLZW
                
            'JPEG - not supported by GDI+
            Case 6
                tiff_Compression = GP_EV_CompressionLZW
            
            'Fax Group 3
            Case 7
                gdipColorDepth = 1
                tiff_Compression = GP_EV_CompressionCCITT3
            
            'Fax Group 4
            Case 8
                gdipColorDepth = 1
                tiff_Compression = GP_EV_CompressionCCITT4
                
        End Select
    
    End If
    
    'Request an encoder from GDI+ based on the type passed to this routine
    Dim exportGuid(0 To 15) As Byte
    
    'GDI+ takes encoder parameters in a very particular sequential format:
    ' 4 byte long: number of encoder parameters
    ' (n) * LenB(GP_EncoderParameter): actual encoder parameters
    '
    'There's no easy way to create a variable-length struct like this in VB6, so instead, we create
    ' an array of encoder params, and as the final step before calling GDI+, we copy everything into
    ' a temporary byte array formatted per GDI+'s requirements.
    Dim numExportParams As Long
    Dim exportParams() As GP_EncoderParameter
    
    PDDebug.LogAction "Preparing GDI+ encoder..."
    
    'Get the clsID for this encoder
    GetEncoderGUIDForPd2dFormat imgFormat, VarPtr(exportGuid(0))
    
    Select Case imgFormat
        
        'BMP export
        Case P2_FFE_BMP
            
            numExportParams = 1
            ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
            
            With exportParams(0)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(gdipColorDepth)
            End With
            
        'GIF export
        Case P2_FFE_GIF
            
            Dim gif_EncoderVersion As GP_EncoderValue
            gif_EncoderVersion = GP_EV_VersionGif89
            
            numExportParams = 1
            ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
            
            With exportParams(0)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_Version), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(gif_EncoderVersion)
            End With
            
        'JPEG export (requires extra work to specify a quality for the encode)
        Case P2_FFE_JPEG
        
            numExportParams = 1
            ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
            
            With exportParams(0)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_Quality), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(jpegQuality)
            End With
            
        'PNG export
        Case P2_FFE_PNG
            
            numExportParams = 1
            ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
            
            With exportParams(0)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(gdipColorDepth)
            End With
            
        'TIFF export (requires extra work to specify compression and color depth for the encode)
        Case P2_FFE_TIFF
            
            numExportParams = 2
            ReDim exportParams(0 To numExportParams - 1) As GP_EncoderParameter
            
            With exportParams(0)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_Compression), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(tiff_Compression)
            End With
            
            With exportParams(1)
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(gdipColorDepth)
            End With
            
    End Select
    
    'Convert our list of params to a format GDI+ understands.
    Dim tmpEncodeParams() As Byte, tmpEncodeParamSize As Long
    If (numExportParams > 0) Then
        tmpEncodeParamSize = 4 + LenB(exportParams(0)) * numExportParams
    Else
        tmpEncodeParamSize = 4
    End If
    
    'First comes the number of parameters
    ReDim tmpEncodeParams(0 To tmpEncodeParamSize - 1) As Byte
    CopyMemoryStrict VarPtr(tmpEncodeParams(0)), VarPtr(numExportParams), 4&
    
    '...followed by each parameter in turn
    If (numExportParams > 0) Then
        Dim i As Long
        For i = 0 To numExportParams - 1
            CopyMemoryStrict VarPtr(tmpEncodeParams(4)) + (i * LenB(exportParams(0))), VarPtr(exportParams(i)), LenB(exportParams(0))
        Next i
    End If
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFilename) Then
        Do
            tmpFilename = dstFilename & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFilename
    End If
    
    'Pass all completed structs to GDI+ and let it handle everything from here
    Dim gpReturn As GP_Result
    gpReturn = GdipSaveImageToFile(hImage, StrPtr(tmpFilename), VarPtr(exportGuid(0)), VarPtr(tmpEncodeParams(0)))
    
    If (gpReturn = GP_OK) Then
        
       'Safe saving: if the destination file already existed, attempt to replace it now
       If Strings.StringsNotEqual(dstFilename, tmpFilename) Then
           Dim overwriteOK As Boolean
           overwriteOK = (Files.FileReplace(dstFilename, tmpFilename) = FPR_SUCCESS)
           If (Not overwriteOK) Then
               Files.FileDelete tmpFilename
               PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
           End If
       End If
                
    Else
        InternalGDIPlusError "GdipSaveImageToFile", "GDI+ failure", gpReturn
        GDI_Plus.ReleaseGDIPlusImage hImage
        GDIPlusSavePicture = False
        Exit Function
    End If
    
    'Release the GDI+ copy of the image
    gpReturn = GdipDisposeImage(hImage)
    If (gpReturn <> GP_OK) Then InternalGDIPlusError "GdipDisposeImage", "GDI+ failure", gpReturn
    
    Message "Save complete."
    
    GDIPlusSavePicture = True
    Exit Function
    
GDIPlusSaveError:
    GDIPlusSavePicture = False
    
End Function

'Quickly export a DIB to PNG format using GDI+.
Public Function GDIPlusQuickSavePNG(ByVal dstFilename As String, ByRef srcDIB As pdDIB) As Boolean

    On Error GoTo GDIPlusQuickSaveError
    
    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hGdipBitmap As Long
    If GetGdipBitmapHandleFromDIB(hGdipBitmap, srcDIB) Then
    
        'Request a PNG encoder from GDI+
        Dim exportGuid(0 To 15) As Byte
        Dim uEncParams As GP_EncoderParameters
        Dim aEncParams() As Byte
            
        GetEncoderGUIDForPd2dFormat P2_FFE_PNG, VarPtr(exportGuid(0))
        uEncParams.EP_Count = 1
        ReDim aEncParams(1 To Len(uEncParams))
        
        Dim gdipColorDepth As Long
        gdipColorDepth = srcDIB.GetDIBColorDepth
        
        With uEncParams.EP_Parameter
            .EP_NumOfValues = 1
            .EP_ValueType = GP_EVT_Long
            CLSIDFromString StrPtr(GP_EP_ColorDepth), VarPtr(.EP_GUID(0))
            .EP_ValuePtr = VarPtr(gdipColorDepth)
        End With
        
        CopyMemoryStrict VarPtr(aEncParams(1)), VarPtr(uEncParams), Len(uEncParams)
        
        'Check to see if a file already exists at this location
        Files.FileDeleteIfExists dstFilename
        
        'Perform the encode and save
        GDIPlusReturn = GdipSaveImageToFile(hGdipBitmap, StrPtr(dstFilename), VarPtr(exportGuid(0)), aEncParams(1))
        GDIPlusQuickSavePNG = (GDIPlusReturn = GP_OK)
        
        'Release the GDI+ copy of the image
        GDIPlusReturn = GdipDisposeImage(hGdipBitmap)
        
    Else
        PDDebug.LogAction "WARNING!  GDI+ QuickSavePNG failed to create a valid gdipBitmap handle."
    End If
    
    Exit Function
    
GDIPlusQuickSaveError:
    GDIPlusQuickSavePNG = False
    
End Function

'I'm not sure whether a pure GDI+ solution or a manual solution is faster, but because the manual solution
' guarantees the smallest possible rect (unlike GDI+), I'm going with it for now.
Public Function IntersectRectF(ByRef dstRect As RectF, ByRef srcRect1 As RectF, ByRef srcRect2 As RectF) As Boolean

    With dstRect
    
        .Left = Max2Float_Single(srcRect1.Left, srcRect2.Left)
        .Width = Min2Float_Single(srcRect1.Left + srcRect1.Width, srcRect2.Left + srcRect2.Width)
        .Top = Max2Float_Single(srcRect1.Top, srcRect2.Top)
        .Height = Min2Float_Single(srcRect1.Top + srcRect1.Height, srcRect2.Top + srcRect2.Height)

        If (.Width >= .Left) And (.Height >= .Top) Then
            .Width = .Width - .Left
            .Height = .Height - .Top
            IntersectRectF = True
        Else
            IntersectRectF = False
        End If
    
    End With
    
End Function

'Nearly identical to StretchBlt, but using GDI+ so we can:
' 1) support fractional source/dest/width/height
' 2) apply variable opacity
' 3) control stretch mode directly inside the call
Public Sub GDIPlus_StretchBlt(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal newAlpha As Single = 1!, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal useThisDestinationDCInstead As Long = 0, Optional ByVal disableEdgeFix As Boolean = False, Optional ByVal isZoomedIn As Boolean = False, Optional ByVal dstCopyIsOkay As Boolean = False)
    
    If (dstDIB Is Nothing) And (useThisDestinationDCInstead = 0) Then Exit Sub
    
    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Currency
    'VBHacks.GetHighResTime profileTime
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim hGraphics As Long, hBitmap As Long
    If (useThisDestinationDCInstead <> 0) Then
        GdipCreateFromHDC useThisDestinationDCInstead, hGraphics
    Else
        GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    End If
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB hBitmap, srcDIB
    
    'hGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If (GdipSetInterpolationMode(hGraphics, interpolationType) = GP_OK) Then
        
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        
        'To improve performance, explicitly request high-speed (aka linear) alpha compositing operation, and standard
        ' pixel offsets (on pixel borders, instead of center points)
        If (Not disableEdgeFix) Then GdipSetImageAttributesWrapMode imgAttributesHandle, GP_WM_TileFlipXY, 0, 0
        GdipSetCompositingQuality hGraphics, GP_CQ_AssumeLinear
        If isZoomedIn Then GdipSetPixelOffsetMode hGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode hGraphics, GP_POM_HighSpeed
        
        'If modified alpha is requested, pass the new value to this image container
        If (newAlpha < 1!) Then
            m_AttributesMatrix(3, 3) = newAlpha
            GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1, VarPtr(m_AttributesMatrix(0, 0)), 0, GP_CMF_Default
        End If
        
        'If the caller doesn't care about source blending (e.g. they're painting to a known transparent destination),
        ' copy mode can improve performance.
        If dstCopyIsOkay Then GdipSetCompositingMode hGraphics, GP_CM_SourceCopy
        
        'Because the resize step is the most cumbersome one, it can be helpful to track it
        'Dim resizeTime As Currency
        'VBHacks.GetHighResTime resizeTime
        
        'Perform the resize
        GdipDrawImageRectRect hGraphics, hBitmap, x1, y1, dstWidth, dstHeight, x2, y2, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&
        
        'Report resize time here
        'Debug.Print "GDI+ resize time: " & Format$(VBHacks.GetTimerDifferenceNow(resizeTime) * 1000, "0000.00") & " ms"
        
        'Release our image attributes object
        GdipDisposeImageAttributes imgAttributesHandle
        
        'Reset alpha in the (reusable) identity matrix
        If (newAlpha < 1!) Then m_AttributesMatrix(3, 3) = 1!
        
        'Update premultiplication status in the target
        If (Not dstDIB Is Nothing) Then dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
        
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDisposeImage hBitmap
    GdipDeleteGraphics hGraphics
    
    'To keep resources low, free both DIBs from their DCs
    If (Not srcDIB Is Nothing) Then srcDIB.FreeFromDC
    If (Not dstDIB Is Nothing) Then dstDIB.FreeFromDC
    
    'Uncomment the line below to receive timing reports
    'Debug.Print "GDI+ wrapper time: " & Format(CStr(VBHacks.GetTimerDifferenceNow(profileTime) * 1000), "0000.00") & " ms"
    
End Sub

'Similar function to GDIPlus_StretchBlt, above, but using a destination parallelogram instead of a rect.
'
'Note that the supplied plgPoints array *MUST HAVE THREE POINTS* in it, in the specific order: top-left, top-right, bottom-left.
' The fourth point is inferred from the other three.
Public Sub GDIPlus_PlgBlt(ByRef dstDIB As pdDIB, ByRef plgPoints() As PointFloat, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal newAlpha As Single = 1!, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal useHQOffsets As Boolean = True, Optional ByVal fixEdges As Boolean = False)

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim hGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, hGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, srcDIB
    
    'hGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(hGraphics, interpolationType) = GP_OK Then
        
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        If (newAlpha <> 1!) Or fixEdges Then GdipCreateImageAttributes imgAttributesHandle Else imgAttributesHandle = 0
        
        'Certain blt operations can cause GDI+ to fuck up edges.  Fix this by asking it to tile the image upon
        ' exceeding boundaries (but only the caller specifies as much).
        If fixEdges Then GdipSetImageAttributesWrapMode imgAttributesHandle, GP_WM_TileFlipXY, 0, 0
        
        'To improve performance and quality, explicitly request high-speed (aka linear) alpha compositing operation, and high-quality
        ' pixel offsets (treat pixels as if they fall on pixel borders, instead of center points - this provides rudimentary edge
        ' antialiasing, which is the best we can do without murdering performance)
        GdipSetCompositingQuality hGraphics, GP_CQ_AssumeLinear
        If useHQOffsets Then GdipSetPixelOffsetMode hGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode hGraphics, GP_POM_HighSpeed
        
        'If modified alpha is requested, pass the new value to this image container
        If (newAlpha <> 1!) Then
            m_AttributesMatrix(3, 3) = newAlpha
            GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1, VarPtr(m_AttributesMatrix(0, 0)), 0, GP_CMF_Default
        End If
        
        'Perform the draw
        GdipDrawImagePointsRect hGraphics, tBitmap, VarPtr(plgPoints(0)), 3, x2, y2, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&
        
        'Release our image attributes object
        If (imgAttributesHandle <> 0) Then GdipDisposeImageAttributes imgAttributesHandle
        
        'Reset alpha in the (reusable) identity matrix
        If (newAlpha <> 1!) Then m_AttributesMatrix(3, 3) = 1!
        
        'Update premultiplication status in the target
        dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
        
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDisposeImage tBitmap
    GdipDeleteGraphics hGraphics
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Sub

'Given a source DIB and an angle, rotate it into a destination DIB.  The destination DIB can be automatically resized
' to fit the rotated image, or a parameter can be set, instructing the function to use the destination DIB "as-is"
Public Sub GDIPlus_RotateDIBPlgStyle(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotateAngle As Single, Optional ByVal dstDIBAlreadySized As Boolean = False, Optional ByVal rotateQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal transparentBackground As Boolean = True, Optional ByVal newBackColor As Long = vbWhite)
    
    'Shortcut angle = 0!
    If (rotateAngle = 0!) Then
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateFromExistingDIB srcDIB
        Exit Sub
    End If
    
    'Before doing any rotating or blurring, we need to figure out the size of our destination image.  If we don't
    ' do this, the rotation will chop off the image's corners!
    Dim nWidth As Double, nHeight As Double
    PDMath.FindBoundarySizeOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, nWidth, nHeight, False
    
    'Use these dimensions to size the destination image, as requested by the user
    If dstDIBAlreadySized Then
        nWidth = dstDIB.GetDIBWidth
        nHeight = dstDIB.GetDIBHeight
        If (Not transparentBackground) Then GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.GetDIBWidth + 1, dstDIB.GetDIBHeight + 1, newBackColor, 255, GP_CM_SourceCopy
    Else
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        If transparentBackground Then
            dstDIB.CreateBlank nWidth, nHeight, srcDIB.GetDIBColorDepth, 0, 0
        Else
            dstDIB.CreateBlank nWidth, nHeight, srcDIB.GetDIBColorDepth, newBackColor, 255
        End If
    End If
    
    'We also want a copy of the corner points of the rotated rect; we'll use these to perform a fast PlgBlt-like operation,
    ' which is how we draw both the rotation and the corner extensions.
    Dim listOfPoints() As PointFloat
    ReDim listOfPoints(0 To 3) As PointFloat
    PDMath.FindCornersOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, listOfPoints, True
    
    'Calculate the size difference between the source and destination images.  We need to add this offset to all
    ' rotation coordinates, to ensure the rotated image is fully contained within the destination DIB.
    Dim hOffset As Double, vOffset As Double
    hOffset = (nWidth - srcDIB.GetDIBWidth) * 0.5
    vOffset = (nHeight - srcDIB.GetDIBHeight) * 0.5
    
    'Apply those offsets to all rotation points, and because GDI+ requires us to use an offset pixel mode for
    ' non-shit results along edges, pad all coordinates with an extra half-pixel as well.
    ' (NOTE: we now move the +0.5 to the very end of the transform, as that's the only place it matters.)
    Dim i As Long
    For i = 0 To 3
        listOfPoints(i).x = listOfPoints(i).x + hOffset
        listOfPoints(i).y = listOfPoints(i).y + vOffset
    Next i
    
    'If a background color is being applied, "cut out" the target region now
    If (Not transparentBackground) Then
        
        Dim tmpPoints() As PointFloat
        ReDim tmpPoints(0 To 3) As PointFloat
        tmpPoints(0) = listOfPoints(0)
        tmpPoints(1) = listOfPoints(1)
        tmpPoints(2) = listOfPoints(3)
        tmpPoints(3) = listOfPoints(2)
        
        'Find the "center" of the rotated image
        Dim cx As Single, cy As Single
        For i = 0 To 3
            cx = cx + tmpPoints(i).x
            cy = cy + tmpPoints(i).y
        Next i
        
        cx = cx * 0.25
        cy = cy * 0.25
        
        'Re-center the points around (0, 0) and add 0.5 for GDI+ half-pixel offset requirements
        For i = 0 To 3
            tmpPoints(i).x = tmpPoints(i).x - cx + 0.5!
            tmpPoints(i).y = tmpPoints(i).y - cy + 0.5!
        Next i
        
        'For each corner of the rotated square, convert the point to polar coordinates, then shrink the radius by one.
        Dim tmpAngle As Double, tmpRadius As Double, tmpX As Double, tmpY As Double
        For i = 0 To 3
            
            PDMath.ConvertCartesianToPolar tmpPoints(i).x, tmpPoints(i).y, tmpRadius, tmpAngle
            tmpRadius = tmpRadius - 1#
            PDMath.ConvertPolarToCartesian tmpAngle, tmpRadius, tmpX, tmpY
            
            'Re-center around the original center point
            tmpPoints(i).x = tmpX + cx
            tmpPoints(i).y = tmpY + cy
            
        Next i
        
        'Paint the selected area transparent
        Dim tmpGraphics As Long, tmpBrush As Long
        GdipCreateFromHDC dstDIB.GetDIBDC, tmpGraphics
        tmpBrush = GetGDIPlusSolidBrushHandle(0, 0)
        
        GdipSetCompositingMode tmpGraphics, GP_CM_SourceCopy
        GdipSetPixelOffsetMode tmpGraphics, GP_POM_HighQuality
        GdipSetInterpolationMode tmpGraphics, rotateQuality
        GdipFillPolygon tmpGraphics, tmpBrush, VarPtr(tmpPoints(0)), 4, GP_FM_Alternate
        GdipSetPixelOffsetMode tmpGraphics, GP_POM_HighSpeed
        GdipSetCompositingMode tmpGraphics, GP_CM_SourceOver
        
        ReleaseGDIPlusBrush tmpBrush
        ReleaseGDIPlusGraphics tmpGraphics
        
    End If
    
    'Rotate the source DIB into the destination DIB.  At this point, corners are still blank - we'll deal with those momentarily.
    GDI_Plus.GDIPlus_PlgBlt dstDIB, listOfPoints, srcDIB, 0!, 0!, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 1!, rotateQuality, True
    
End Sub

'Given a regular ol' DIB and an angle, return a DIB that is rotated by that angle, with its edge values clamped and extended
' to fill all empty space around the rotated image.  This very cool operation allows us to support angles for any filter
' with a grid implementation (e.g. something that operates on the (x, y) axes of an image, like pixellate or blur).
Public Sub GDIPlus_GetRotatedClampedDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotateAngle As Single, Optional ByVal padToIntegerCalcs As Boolean = True)
    
    'Before doing any rotating or blurring, we need to figure out the size of our destination image.  If we don't
    ' do this, the rotation will chop off the image's corners!
    Dim nWidth As Double, nHeight As Double
    PDMath.FindBoundarySizeOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, nWidth, nHeight, padToIntegerCalcs
    
    'Use these dimensions to size the destination image
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Shortcut angle = 0
    If (rotateAngle = 0!) Then
        dstDIB.CreateFromExistingDIB srcDIB
        Exit Sub
    End If
    
    If (dstDIB.GetDIBWidth <> nWidth) Or (dstDIB.GetDIBHeight <> nHeight) Or (dstDIB.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
        dstDIB.CreateBlank nWidth, nHeight, srcDIB.GetDIBColorDepth, 0, 0
    Else
        dstDIB.ResetDIB 0
    End If
    
    'We also want a copy of the corner points of the rotated rect; we'll use these to perform a fast PlgBlt-like operation,
    ' which is how we draw both the rotation and the corner extensions.
    Dim listOfPoints() As PointFloat
    ReDim listOfPoints(0 To 3) As PointFloat
    PDMath.FindCornersOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, listOfPoints, True
    
    'Calculate the size difference between the source and destination images.  We need to add this offset to all
    ' rotation coordinates, to ensure the rotated image is fully contained within the destination DIB.
    Dim hOffset As Double, vOffset As Double
    hOffset = (nWidth - srcDIB.GetDIBWidth) / 2#
    vOffset = (nHeight - srcDIB.GetDIBHeight) / 2#
    
    'Apply those offsets to all rotation points, and because GDI+ requires us to use an offset pixel mode for
    ' non-shit results along edges, pad all coordinates with an extra half-pixel as well.
    Dim useHalfPixels As Boolean, fixEdges As Boolean
    useHalfPixels = True
    fixEdges = True
    
    Dim i As Long
    For i = 0 To 3
        listOfPoints(i).x = listOfPoints(i).x + hOffset
        listOfPoints(i).y = listOfPoints(i).y + vOffset
        If useHalfPixels Then
            listOfPoints(i).x = listOfPoints(i).x + 0.5
            listOfPoints(i).y = listOfPoints(i).y + 0.5
        End If
    Next i
    
    'Rotate the source DIB into the destination DIB.  At this point, corners are still blank - we'll deal with those momentarily.
    GDI_Plus.GDIPlus_PlgBlt dstDIB, listOfPoints, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 1, GP_IM_HighQualityBilinear, useHalfPixels, fixEdges
    
    Dim intrplMode As GP_InterpolationMode
    intrplMode = GP_IM_HighQualityBilinear
    useHalfPixels = True
    fixEdges = True
    
    'We're now going to calculate a whole bunch of geometry based around the concept of extending a rectangle from
    ' each edge of our rotated image, out to the corner of the rotation DIB.  We will then fill this dead space with a
    ' stretched version of the image edge, resulting in "clamped" edge behavior.
    Dim diagDistance As Double, distDiff As Double
    Dim dx As Double, dy As Double, lineLength As Double, pX As Double, pY As Double
    Dim padPoints() As PointFloat
    ReDim padPoints(0 To 2) As PointFloat
    
    'Calculate the distance from the center of the rotated image to the corner of the rotated image
    diagDistance = Sqr(nWidth * nWidth + nHeight * nHeight) * 0.5
    
    'Get the difference between the diagonal distance, and the original height of the image.  This is the distance
    ' where we need to provide clamped pixels on this edge.
    distDiff = diagDistance - (srcDIB.GetDIBHeight / 2#)
    
    'Calculate delta x/y values for the top line, then convert those into unit vectors
    dx = listOfPoints(1).x - listOfPoints(0).x
    dy = listOfPoints(1).y - listOfPoints(0).y
    lineLength = Sqr(dx * dx + dy * dy)
    dx = dx / lineLength
    dy = dy / lineLength
    
    'dX/Y now represent a vector in the direction of the line.  We want a perpendicular vector instead (because we're
    ' extending a rectangle out from that image edge), and we want the vector to be of length distDiff, so it reaches
    ' all the way to the corner.
    pX = distDiff * -dy
    pY = distDiff * dx
    
    'Use this perpendicular vector to calculate new parallelogram coordinates, which "extrude" the top of the image
    ' from where it appears on the rotated image, to the very edge of the image.
    padPoints(0).x = listOfPoints(0).x - pX
    padPoints(0).y = listOfPoints(0).y - pY
    padPoints(1).x = listOfPoints(1).x - pX
    padPoints(1).y = listOfPoints(1).y - pY
    padPoints(2).x = listOfPoints(0).x
    padPoints(2).y = listOfPoints(0).y
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, 0, srcDIB.GetDIBWidth, 1, 1, intrplMode, useHalfPixels, fixEdges
    
    'Now repeat the above steps for the bottom of the image.  Note that we can reuse almost all of the calculations,
    ' as this line is parallel to the one we just calculated.
    padPoints(0).x = listOfPoints(2).x - (pX / distDiff)
    padPoints(0).y = listOfPoints(2).y - (pY / distDiff)
    padPoints(1).x = listOfPoints(3).x - (pX / distDiff)
    padPoints(1).y = listOfPoints(3).y - (pY / distDiff)
    padPoints(2).x = listOfPoints(2).x + pX
    padPoints(2).y = listOfPoints(2).y + pY
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, srcDIB.GetDIBHeight - 2, srcDIB.GetDIBWidth, 1, 1, intrplMode, useHalfPixels, fixEdges
    
    'We are now going to repeat the above steps, but for the left and right edges of the image.  The end result of this
    ' will be a rotated destination image, with clamped values extending from all image edges.
    
    'Get the difference between the diagonal distance, and the original width of the image.  This is the distance
    ' where we need to provide clamped pixels on this edge.
    distDiff = diagDistance - (srcDIB.GetDIBWidth / 2#)
    
    'Calculate delta x/y values for the left line, then convert those into unit vectors
    dx = listOfPoints(2).x - listOfPoints(0).x
    dy = listOfPoints(2).y - listOfPoints(0).y
    lineLength = Sqr(dx * dx + dy * dy)
    dx = dx / lineLength
    dy = dy / lineLength
    
    'dX/Y now represent a vector in the direction of the line.  We want a perpendicular vector instead,
    ' of length distDiff.
    pX = distDiff * -dy
    pY = distDiff * dx
    
    'Use the perpendicular vector to calculate new parallelogram coordinates, which "extrude" the left of the image
    ' from where it appears on the rotated image, to the very edge of the image.
    padPoints(0).x = listOfPoints(0).x + pX
    padPoints(0).y = listOfPoints(0).y + pY
    padPoints(1).x = listOfPoints(0).x
    padPoints(1).y = listOfPoints(0).y
    padPoints(2).x = listOfPoints(2).x + pX
    padPoints(2).y = listOfPoints(2).y + pY
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, 0, 1, srcDIB.GetDIBHeight, 1, intrplMode, useHalfPixels, fixEdges
    
    '...and finally, repeat everything for the right side of the image
    padPoints(0).x = listOfPoints(1).x + (pX / distDiff)
    padPoints(0).y = listOfPoints(1).y + (pY / distDiff)
    padPoints(1).x = listOfPoints(1).x - pX
    padPoints(1).y = listOfPoints(1).y - pY
    padPoints(2).x = listOfPoints(3).x + (pX / distDiff)
    padPoints(2).y = listOfPoints(3).y + (pY / distDiff)
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, srcDIB.GetDIBWidth - 2, 0, 1, srcDIB.GetDIBHeight, 1, intrplMode, useHalfPixels, fixEdges
    
    'Our work here is complete!

End Sub

'Given a GUID string, return a Long-type image format identifier
Private Function GetFIFFromGUID(ByRef srcGUID As String) As PD_IMAGE_FORMAT
    
    Select Case srcGUID
    
        Case "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_BMP
            
        Case "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_EMF
            
        Case "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_WMF
        
        Case "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_JPEG
            
        Case "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_PNG
            
        Case "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_GIF
            
        Case "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = PDIF_TIFF
            
        Case "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"
            GetFIFFromGUID = FIF_ICO
        
        Case Else
            GetFIFFromGUID = PDIF_UNKNOWN
            
    End Select

End Function


'************************************************************************************************************
'
'This module is currently undergoing heavy clean-up.  Functions that have been modernized and revised
' are included below this line.  (The goal is to move *all* functions down here eventually.)
'
'************************************************************************************************************



'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function GDIP_StartEngine(Optional ByVal hookDebugProc As Boolean = False) As Boolean
    
    'Prep a generic GDI+ startup interface
    Dim gdiCheck As GDIPlusStartupInput
    With gdiCheck
        .GDIPlusVersion = 1&
        
        'Hypothetically you could set a callback function here, but I haven't tested this thoroughly, so use with caution!
        'If hookDebugProc Then
        '    .DebugEventCallback = FakeProcPtr(AddressOf GDIP_Debug_Proc)
        'Else
            .DebugEventCallback = 0&
        'End If
        
        .SuppressBackgroundThread = 0&
        .SuppressExternalCodecs = 0&
    End With
    
    'Retrieve a GDI+ token for this session
    GDIP_StartEngine = (GdiplusStartup(m_GDIPlusToken, gdiCheck, 0&) = GP_OK)
    If GDIP_StartEngine Then
        
        'As a convenience, create a dummy graphics container.  This is useful for various GDI+ functions that require world
        ' transformation data.
        Set m_TransformDIB = New pdDIB
        m_TransformDIB.CreateBlank 8, 8, 32, 0, 0
        GdipCreateFromHDC m_TransformDIB.GetDIBDC, m_TransformGraphics
        
        'Note that these dummy objects are released when GDI+ terminates.
        
        'Next, create a default identity matrix for image attributes.
        ReDim m_AttributesMatrix(0 To 4, 0 To 4) As Single
        m_AttributesMatrix(0, 0) = 1!
        m_AttributesMatrix(1, 1) = 1!
        m_AttributesMatrix(2, 2) = 1!
        m_AttributesMatrix(3, 3) = 1!
        m_AttributesMatrix(4, 4) = 1!
        
        'Want to list all available decoders?  Do so here.
        'DEBUG_ListGdipDecoders
        
        'Next, check to see if v1.1 is available.  This allows for advanced fx work.
        Dim hMod As Long, strGDIPName As String
        strGDIPName = "gdiplus.dll"
        hMod = LoadLibrary(StrPtr(strGDIPName))
        If (hMod <> 0) Then
            Dim testAddress As Long
            testAddress = GetProcAddress(hMod, "GdipDrawImageFX")
            m_GDIPlus11Available = (testAddress <> 0)
            FreeLibrary hMod
        Else
            m_GDIPlus11Available = False
        End If
        
    Else
        m_GDIPlus11Available = False
    End If

End Function

'At shutdown, this function must be called to release our GDI+ instance
Public Function GDIP_StopEngine() As Boolean
    
    'Release any custom collection we have created
    GDIPlus_ReleaseRuntimeFonts
    
    'Release any dummy containers we have created
    GdipDeleteGraphics m_TransformGraphics
    Set m_TransformDIB = Nothing
    
    'Release GDI+ using the same token we received at startup time
    GDIP_StopEngine = (GdiplusShutdown(m_GDIPlusToken) = GP_OK)
    
End Function

'Want to know if GDI+ v1.1 is available?  Use this wrapper.
Public Function IsGDIPlusV11Available() As Boolean
    IsGDIPlusV11Available = m_GDIPlus11Available
End Function

Private Function FakeProcPtr(ByVal AddressOfResult As Long) As Long
    FakeProcPtr = AddressOfResult
End Function

'At GDI+ startup, the caller can request that we provide a debug proc for GDI+ to call on warnings and errors.
' This is that proc.
'
'NOTE: this feature is currently disabled due to lack of testing.
'Private Function GDIP_Debug_Proc(ByVal deLevel As GP_DebugEventLevel, ByVal ptrChar As Long) As Long
'
'    'Pull the GDI+ message into a local string
'    Dim debugString As String
'    'debugString = Strings.StringFromCharPtr(ptrChar, False)
'    debugString = "Unknown GDI+ error was passed to the GDIPlus debug procedure."
'
'    If (deLevel = GP_DebugEventLevelWarning) Then
'        Debug.Print "GDI+ WARNING: " & debugString
'    ElseIf (deLevel = GP_DebugEventLevelFatal) Then
'        Debug.Print "GDI+ ERROR: " & debugString
'    Else
'        Debug.Print "GDI+ UNKNOWN: " & debugString
'    End If
'
'End Function

Private Sub InternalGDIPlusError(Optional ByVal errName As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal errNumber As GP_Result = GP_OK)
        
    'If the caller passes an error number but no error name, attempt to automatically populate
    ' it based on the error number.
    If ((LenB(errName) = 0) And (errNumber <> GP_OK)) Then
        
        Select Case errNumber
            Case GP_GenericError
                errName = "Generic Error"
            Case GP_InvalidParameter
                errName = "Invalid parameter"
            Case GP_OutOfMemory
                errName = "Out of memory"
            Case GP_ObjectBusy
                errName = "Object busy"
            Case GP_InsufficientBuffer
                errName = "Insufficient buffer size"
            Case GP_NotImplemented
                errName = "Feature is not implemented"
            Case GP_Win32Error
                errName = "Win32 error"
            Case GP_WrongState
                errName = "Wrong state"
            Case GP_Aborted
                errName = "Operation aborted"
            Case GP_FileNotFound
                errName = "File not found"
            Case GP_ValueOverflow
                errName = "Value too large (overflow)"
            Case GP_AccessDenied
                errName = "Access denied"
            Case GP_UnknownImageFormat
                errName = "Image format was not recognized"
            Case GP_FontFamilyNotFound
                errName = "Font family not found"
            Case GP_FontStyleNotFound
                errName = "Font style not found"
            Case GP_NotTrueTypeFont
                errName = "Font is not TrueType (only TT fonts are supported)"
            Case GP_UnsupportedGDIPlusVersion
                errName = "GDI+ version is not supported"
            Case GP_GDIPlusNotInitialized
                errName = "GDI+ was not initialized correctly"
            Case GP_PropertyNotFound
                errName = "Property missing"
            Case GP_PropertyNotSupported
                errName = "Property not supported"
            Case Else
                errName = "Undefined error (number doesn't match known returns)"
        End Select
        
    End If
    
    Dim tmpString As String
    If (errNumber <> 0) Then
        tmpString = "WARNING!  Internal GDI+ error #" & errNumber & ", """ & errName & """"
    Else
        tmpString = "WARNING!  GDI+ module error, """ & errName & """"
    End If
    
    If (LenB(errDescription) <> 0) Then tmpString = tmpString & ": " & errDescription
    PDDebug.LogAction tmpString, PDM_External_Lib
    
End Sub

'GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from a standard RGB()
' long and supplied alpha.  It's not a very efficient conversion, but I need it so infrequently that I don't really care.
Public Function FillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
    
    'The vbRGB constant may be an OLE color constant; if that happens, we want to convert it to a normal RGB quad.
    vbRGB = TranslateColor(vbRGB)
    
    Dim dstQuad As RGBQuad
    dstQuad.Red = Drawing2D.ExtractRed(vbRGB)
    dstQuad.Green = Drawing2D.ExtractGreen(vbRGB)
    dstQuad.Blue = Drawing2D.ExtractBlue(vbRGB)
    dstQuad.Alpha = alphaValue
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    FillQuadWithVBRGB = placeHolder.lngResult
    
End Function

Public Function FillLongWithRGBA(ByVal srcR As Long, ByVal srcG As Long, ByVal srcB As Long, ByVal srcA As Long) As Long
    
    Dim dstQuad As RGBQuad
    dstQuad.Red = srcR
    dstQuad.Green = srcG
    dstQuad.Blue = srcB
    dstQuad.Alpha = srcA
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    FillLongWithRGBA = placeHolder.lngResult
    
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the opacity value on the scale [0, 100]
Public Function GetOpacityFromPARGB(ByVal pARGB As Long) As Single
    Dim srcQuad As RGBQuad
    CopyMemoryStrict VarPtr(srcQuad), VarPtr(pARGB), 4&
    GetOpacityFromPARGB = CSng(srcQuad.Alpha) * (100! / 255!)
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the RGB component in combined vbRGB format.
' (Note that various GDI+ settings, like brush colors, return their RGBA results as *not* premultiplied.)
Public Function GetColorFromPARGB(ByVal pARGB As Long, Optional ByVal removePremultiplication As Boolean = False) As Long
    
    Dim srcQuad As RGBQuad
    CopyMemoryStrict VarPtr(srcQuad), VarPtr(pARGB), 4&
    
    If removePremultiplication Then
    
        If (srcQuad.Alpha = 255) Then
            GetColorFromPARGB = RGB(srcQuad.Red, srcQuad.Green, srcQuad.Blue)
        Else
        
            Dim tmpSingle As Single
            Const ONE_DIV_255 As Single = 1! / 255!
            tmpSingle = CSng(srcQuad.Alpha) * ONE_DIV_255
            
            If (tmpSingle > 0!) Then
                Dim tmpRed As Long, tmpGreen As Long, tmpBlue As Long
                tmpRed = CSng(srcQuad.Red) / tmpSingle
                tmpGreen = CSng(srcQuad.Green) / tmpSingle
                tmpBlue = CSng(srcQuad.Blue) / tmpSingle
                GetColorFromPARGB = RGB(tmpRed, tmpGreen, tmpBlue)
            Else
                GetColorFromPARGB = 0
            End If
            
        End If
        
    Else
        GetColorFromPARGB = RGB(srcQuad.Red, srcQuad.Green, srcQuad.Blue)
    End If
    
End Function

'Translate an OLE color to an RGB Long.  Note that the API function returns -1 on failure; if this happens, we return white.
Private Function TranslateColor(ByVal colorRef As Long) As Long
    If OleTranslateColor(colorRef, 0, TranslateColor) Then TranslateColor = vbWhite
End Function

Public Function GetGDIPlusSolidBrushHandle(ByVal brushColor As Long, Optional ByVal brushOpacity As Byte = 255) As Long
    GdipCreateSolidFill FillQuadWithVBRGB(brushColor, brushOpacity), GetGDIPlusSolidBrushHandle
End Function

Public Function GetGDIPlusLinearBrushHandle(ByRef srcRect As RectF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As PD_2D_WrapMode) As Long
    Dim gdiReturn As Long
    gdiReturn = GdipCreateLineBrushFromRectWithAngle(srcRect, firstRGBA, secondRGBA, gradAngle, isAngleScalable, gradientWrapMode, GetGDIPlusLinearBrushHandle)
    If (gdiReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, gdiReturn
End Function

Public Function OverrideGDIPlusLinearGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusLinearGradient = (GdipSetLinePresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

Public Function GetGDIPlusPathBrushHandle(ByVal hGraphicsPath As Long) As Long
    GdipCreatePathGradientFromPath hGraphicsPath, GetGDIPlusPathBrushHandle
End Function

Public Function SetGDIPlusPathBrushCenter(ByVal hBrush As Long, ByVal centerX As Single, ByVal centerY As Single) As Long
    Dim centerPoint As PointFloat
    centerPoint.x = centerX
    centerPoint.y = centerY
    GdipSetPathGradientCenterPoint hBrush, centerPoint
End Function

Public Function SetGDIPlusPathBrushWrap(ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As Boolean
    SetGDIPlusPathBrushWrap = (GdipSetPathGradientWrapMode(hBrush, newWrapMode) = GP_OK)
End Function

Public Function OverrideGDIPlusPathGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusPathGradient = (GdipSetPathGradientPresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

'Simpler shorthand function for obtaining a GDI+ bitmap handle from a pdDIB object.  Note that 24/32bpp cases have to be
' handled separately because GDI+ is unpredictable at automatically detecting color depth with 32-bpp DIBs.  (This behavior
' is forgivable, given GDI's unreliable handling of alpha bytes.)
Public Function GetGdipBitmapHandleFromDIB(ByRef dstBitmapHandle As Long, ByRef srcDIB As pdDIB) As Boolean
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetGdipBitmapHandleFromDIB = (GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBStride, GP_PF_32bppPARGB, srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.GetDIBColorDepth
            .Width = srcDIB.GetDIBWidth
            .Height = -srcDIB.GetDIBHeight
        End With
        GetGdipBitmapHandleFromDIB = (GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
        
    End If

End Function

'Retrieve a persistent handle to a GDI+-format graphics container.  Optionally, a smoothing mode can be specified so that it does
' not have to be repeatedly specified by a caller function.  (GDI+ sets smoothing mode by graphics container, not by function call.)
Public Function GetGDIPlusGraphicsFromDC(ByVal srcDC As Long, Optional ByVal graphicsAntialiasing As GP_SmoothingMode = GP_SM_None, Optional ByVal graphicsPixelOffsetMode As GP_PixelOffsetMode = GP_POM_None) As Long
    If (GdipCreateFromHDC(srcDC, GetGDIPlusGraphicsFromDC) = GP_OK) Then
        GdipSetSmoothingMode GetGDIPlusGraphicsFromDC, graphicsAntialiasing
        GDI_Plus.SetGDIPlusGraphicsPixelOffset GetGDIPlusGraphicsFromDC, graphicsPixelOffsetMode
    Else
        GetGDIPlusGraphicsFromDC = 0
    End If
End Function

'Retrieve a persistent handle to a GDI+-format graphics container.
Public Function GetGDIPlusGraphicsFromDC_Fast(ByVal srcDC As Long) As Long
    Dim gpResult As GP_Result
    gpResult = GdipCreateFromHDC(srcDC, GetGDIPlusGraphicsFromDC_Fast)
    If (gpResult <> GP_OK) Then
        GetGDIPlusGraphicsFromDC_Fast = 0
        InternalGDIPlusError "GetGDIPlusGraphicsFromDC_Fast failed", "CreateFromHDC failed", gpResult
    End If
End Function

'I'm honestly not sure how creating a graphics object from an hWnd works (it's possible the hWnd is just
' used to drive color-management, since the corresponding System.Drawing method takes icm as an input),
' but this is useful in PD for objects like regions that are unlikely to be naturally associated with a DC.
Public Function GetGDIPlusGraphicsFromHWnd(ByVal srcHWnd As Long) As Long
    Dim gpResult As GP_Result
    gpResult = GdipCreateFromHWND(srcHWnd, GetGDIPlusGraphicsFromHWnd)
    If (gpResult <> GP_OK) Then
        GetGDIPlusGraphicsFromHWnd = 0
        InternalGDIPlusError "GetGDIPlusGraphicsFromHWnd failed", errNumber:=gpResult
    End If
End Function

Public Function ReleaseGDIPlusBrush(ByRef srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusBrush = (GdipDeleteBrush(srcHandle) = GP_OK)
        If ReleaseGDIPlusBrush Then srcHandle = 0
    Else
        ReleaseGDIPlusBrush = True
    End If
End Function

Public Function ReleaseGDIPlusGraphics(ByRef srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusGraphics = (GdipDeleteGraphics(srcHandle) = GP_OK)
        If ReleaseGDIPlusGraphics Then srcHandle = 0
    Else
        ReleaseGDIPlusGraphics = True
    End If
End Function

Public Function ReleaseGDIPlusImage(ByRef srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusImage = (GdipDisposeImage(srcHandle) = GP_OK)
        If ReleaseGDIPlusImage Then srcHandle = 0 Else InternalGDIPlusError "ReleaseGDIPlusImage failed", , ReleaseGDIPlusImage
    Else
        ReleaseGDIPlusImage = True
    End If
End Function

Public Function SetGDIPlusGraphicsPixelOffset(ByVal hGraphics As Long, ByVal newSetting As GP_PixelOffsetMode) As Boolean
    If (hGraphics <> 0) Then SetGDIPlusGraphicsPixelOffset = (GdipSetPixelOffsetMode(hGraphics, newSetting) = GP_OK)
End Function

Public Function SetGDIPlusGraphicsBlendUsingSRGBGamma(ByVal hGraphics As Long, ByVal newSetting As GP_CompositingQuality) As Boolean
    If (hGraphics <> 0) Then SetGDIPlusGraphicsBlendUsingSRGBGamma = (GdipSetCompositingQuality(hGraphics, newSetting) = GP_OK)
End Function

Public Function GDIPlus_SetTextureBrushTransform(ByVal hBrush As Long, ByVal hTransform As Long) As Boolean
    GDIPlus_SetTextureBrushTransform = (GdipSetTextureTransform(hBrush, hTransform) = GP_OK)
End Function

'All generic draw and fill functions follow

'GDI+ arcs use bounding boxes to describe their placement.  As such, we manually convert the incoming centerX/Y and radius values
' to bounding box coordinates.
Public Function GDIPlus_DrawArcF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    GDIPlus_DrawArcF = (GdipDrawArc(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawArcI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    GDIPlus_DrawArcI = (GdipDrawArcI(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawClosedCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawClosedCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImage(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageI(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1!) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1!) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the (reusable) identity matrix
    If (opacityModifier <> 1!) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1!
    End If
        
End Function

Public Function GDIPlus_DrawImageRectRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1!) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1!) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the (reusable) identity matrix
    If (opacityModifier <> 1!) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1!
    End If
        
End Function

Public Function GDIPlus_DrawImagePointsRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As PointFloat, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1!) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1!) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRect(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the (reusable) identity matrix
    If (opacityModifier <> 1!) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1!
    End If
        
End Function

Public Function GDIPlus_DrawImagePointsRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As PointLong, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1!) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1!) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRectI(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the (reusable) identity matrix
    If (opacityModifier <> 1!) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1!
    End If
        
End Function

Public Function GDIPlus_DrawLineF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLine(dstGraphics, srcPen, x1, y1, x2, y2)
    GDIPlus_DrawLineF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawLineI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    GDIPlus_DrawLineI = (GdipDrawLineI(dstGraphics, srcPen, x1, y1, x2, y2) = GP_OK)
End Function

Public Function GDIPlus_DrawLinesF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLines(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawLinesF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawLinesI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLinesI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawLinesI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPath(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_DrawPath = (GdipDrawPath(dstGraphics, srcPen, srcPath) = GP_OK)
End Function

Public Function GDIPlus_DrawPolygonF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygon(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawPolygonF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPolygonI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygonI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawPolygonI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawRectF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_DrawRectF = (GdipDrawRectangle(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawRectI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_DrawRectI = (GdipDrawRectangleI(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_DrawEllipseF = (GdipDrawEllipse(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_DrawEllipseI = (GdipDrawEllipseI(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillClosedCurveF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillClosedCurveI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5!, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2I(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPath(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_FillPath = (GdipFillPath(dstGraphics, srcBrush, srcPath) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_FillEllipseF = (GdipFillEllipse(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_FillEllipseI = (GdipFillEllipseI(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillPolygonF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygon(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonF = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPolygonI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygonI(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonI = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillRectF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_FillRectF = (GdipFillRectangle(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_FillRectI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_FillRectI = (GdipFillRectangleI(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_FillRegion(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcRegion As Long) As Boolean
    GDIPlus_FillRegion = (GdipFillRegion(dstGraphics, srcBrush, srcRegion) = GP_OK)
End Function

Public Function GDIPlus_GraphicsSetCompositingMode(ByVal dstGraphics As Long, Optional ByVal newCompositeMode As GP_CompositingMode = GP_CM_SourceOver) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetCompositingMode(dstGraphics, newCompositeMode)
    GDIPlus_GraphicsSetCompositingMode = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Note that this function creates an image from an array containing a valid image file (e.g. not an array with
' bare RGB values).  This is helpful for interop with other software, or if you prefer to roll your own filesystem code.
Public Function GDIPlus_ImageCreateFromArray(ByRef srcArray() As Byte, Optional ByRef isImageMetafile As Boolean = False) As Long
    
    'GDI+ requires a stream object for import, so we're going to wrap a temporary stream around the source array.
    Dim tmpStream As Long
    
    Dim tmpHMem As Long
    Const GMEM_MOVEABLE As Long = &H2&
    tmpHMem = GlobalAlloc(GMEM_MOVEABLE, UBound(srcArray) - LBound(srcArray) + 1)
    If (tmpHMem <> 0) Then
        
        Dim tmpLockMem As Long
        tmpLockMem = GlobalLock(tmpHMem)
        If (tmpLockMem <> 0) Then
            CopyMemoryStrict tmpLockMem, VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1
            GlobalUnlock tmpHMem
            CreateStreamOnHGlobal tmpHMem, 1&, VarPtr(tmpStream)
        End If
        
    End If
    
    If (tmpStream <> 0) Then
    
        Dim tmpReturn As GP_Result
        tmpReturn = GdipLoadImageFromStream(tmpStream, GDIPlus_ImageCreateFromArray)
        If (tmpReturn = GP_OK) Then
            Dim imgType As GP_ImageType
            GdipGetImageType GDIPlus_ImageCreateFromArray, imgType
            isImageMetafile = (imgType = GP_IT_Metafile)
        Else
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
        
    Else
        InternalGDIPlusError "IStream failure", "GDIPlus_ImageCreateFromArray() failed to wrap an IStream around the source array; load aborted."
    End If
    
End Function

Public Function GDIPlus_ImageCreateFromFile(ByVal srcFilename As String, Optional ByRef isImageMetafile As Boolean = False) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipLoadImageFromFile(StrPtr(srcFilename), GDIPlus_ImageCreateFromFile)
    If (tmpReturn = GP_OK) Then
        Dim imgType As GP_ImageType
        GdipGetImageType GDIPlus_ImageCreateFromFile, imgType
        isImageMetafile = (imgType = GP_IT_Metafile)
    Else
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
End Function

'Note that this function creates an image from a bare pointer that points to a valid image file
' (e.g. not a stream of bare RGB values - a valid BMP/GIF/JPEG/TIFF wrapper must be used).
Public Function GDIPlus_ImageCreateFromPtr(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByRef isImageMetafile As Boolean = False) As Long
    
    'GDI+ requires a stream object for import, so we're going to wrap a temporary stream around the pointer.
    Dim tmpStream As Long
    
    Dim tmpHMem As Long
    Const GMEM_MOVEABLE As Long = &H2&
    tmpHMem = GlobalAlloc(GMEM_MOVEABLE, srcLen)
    If (tmpHMem <> 0) Then
        
        Dim tmpLockMem As Long
        tmpLockMem = GlobalLock(tmpHMem)
        If (tmpLockMem <> 0) Then
            CopyMemoryStrict tmpLockMem, srcPtr, srcLen
            GlobalUnlock tmpHMem
            CreateStreamOnHGlobal tmpHMem, 1&, VarPtr(tmpStream)
        End If
        
    End If
    
    If (tmpStream <> 0) Then
    
        Dim tmpReturn As GP_Result
        tmpReturn = GdipLoadImageFromStream(tmpStream, GDIPlus_ImageCreateFromPtr)
        If (tmpReturn = GP_OK) Then
            Dim imgType As GP_ImageType
            GdipGetImageType GDIPlus_ImageCreateFromPtr, imgType
            isImageMetafile = (imgType = GP_IT_Metafile)
        Else
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
        
    Else
        InternalGDIPlusError "IStream failure", "GDIPlus_ImageCreateFromPtr() failed to wrap an IStream around the source pointer; load aborted."
    End If
    
End Function

'This function only works on bitmaps (never metafiles!), and the source image *must* already be in 32-bpp format.
Public Function GDIPlus_ImageForcePremultipliedAlpha(ByVal hImage As Long, ByVal imgWidth As Long, ByVal imgHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCloneBitmapAreaI(0, 0, imgWidth, imgHeight, GP_PF_32bppPARGB, hImage, hImage)
    GDIPlus_ImageForcePremultipliedAlpha = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'RANDOM FACT! GdipGetImageDimension works fine on bitmaps.  On metafiles, it returns bizarre values that may be
' astronomically large.  I assume that metafile dimensions are not necessarily returned in pixels (though pixels
' are the default for bitmaps...?).  Anyway, to avoid this problem, we only use GdipGetImageWidth/Height, which
' always return "correct" pixel values.
Public Function GDIPlus_ImageGetDimensions(ByVal hImage As Long, ByRef dstWidth As Long, ByRef dstHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImageWidth(hImage, dstWidth)
    If (tmpReturn = GP_OK) Then
        tmpReturn = GdipGetImageHeight(hImage, dstHeight)
        GDIPlus_ImageGetDimensions = (tmpReturn = GP_OK)
    Else
        GDIPlus_ImageGetDimensions = False
    End If
End Function

Public Function GDIPlus_ImageGetFileFormat(ByVal hImage As Long) As PD_2D_FileFormatImport
    GDIPlus_ImageGetFileFormat = GetPd2dFileFormatFromGUID(GDIPlus_ImageGetFileFormatGUID(hImage))
End Function

Public Function GDIPlus_ImageGetFileFormatGUID(ByVal hImage As Long) As String
    
    Dim tmpReturn As GP_Result
    
    'Start by retrieving the raw bytes of the GUID
    Dim guidBytes() As Byte
    ReDim guidBytes(0 To 15) As Byte
    tmpReturn = GdipGetImageRawFormat(hImage, VarPtr(guidBytes(0)))
    
    If (tmpReturn = GP_OK) Then
    
        'Byte array comparisons against predefined constants are messy in VB, so retrieve a string instead
        Dim imgStringPointer As Long
        If (StringFromCLSID(VarPtr(guidBytes(0)), imgStringPointer) = 0) Then
            Dim strLength As Long
            strLength = lstrlenW(imgStringPointer)
            If (strLength <> 0) Then
                GDIPlus_ImageGetFileFormatGUID = String$(strLength, 48)
                CopyMemoryStrict StrPtr(GDIPlus_ImageGetFileFormatGUID), imgStringPointer, strLength * 2
            End If
        Else
            InternalGDIPlusError "Failed to convert clsID to string", "GDIPlus_ImageGetFileFormatGUID failed"
        End If
        
    Else
        GDIPlus_ImageGetFileFormatGUID = GP_FF_GUID_Undefined
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
    
End Function

'Given a GDI+ GUID format identifier, return a Long-type pd2D file format identifier
Private Function GetPd2dFileFormatFromGUID(ByRef srcGUID As String) As PD_2D_FileFormatImport
    Select Case srcGUID
        Case GP_FF_GUID_BMP, GP_FF_GUID_MemoryBMP
            GetPd2dFileFormatFromGUID = P2_FFI_BMP
        Case GP_FF_GUID_EMF
            GetPd2dFileFormatFromGUID = P2_FFI_EMF
        Case GP_FF_GUID_WMF
            GetPd2dFileFormatFromGUID = P2_FFI_WMF
        Case GP_FF_GUID_JPEG
            GetPd2dFileFormatFromGUID = P2_FFI_JPEG
        Case GP_FF_GUID_PNG
            GetPd2dFileFormatFromGUID = P2_FFI_PNG
        Case GP_FF_GUID_GIF
            GetPd2dFileFormatFromGUID = P2_FFI_GIF
        Case GP_FF_GUID_TIFF
            GetPd2dFileFormatFromGUID = P2_FFI_TIFF
        Case GP_FF_GUID_Icon
            GetPd2dFileFormatFromGUID = P2_FFI_ICO
        Case Else
            GetPd2dFileFormatFromGUID = P2_FFI_Undefined
    End Select
End Function

Public Function GDIPlus_ImageGetPixelFormat(ByVal hImage As Long) As GP_PixelFormat
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImagePixelFormat(hImage, GDIPlus_ImageGetPixelFormat)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Retrieve an image "property" (usually something defined inside a metadata region).  For example, you can
' use this to pull an ICC profile out of an image, or EXIF data like "orientation" for smartphone photos.
'
'Note that it is *critical* to check the return value of this function.  It will return FALSE if the image
' does not contain/provide/support the requested property.
'
'Also note that all properties are returned as byte arrays.  It is up to the caller to make sense of the
' bytes in whatever way is appropriate for the requested property.  This MSDN guide can help with that:
' https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx
'
'The returned header can also help you interpret the data correctly, but note that we have to format the
' header type a little weirdly due to padding issues.  As such, the "property type" member is not
' specifically declared as an enum, although you can treat it as an integer of GP_PropertyTagType.
'
'Note also that the header's property value pointer is manually overwritten by this function, so that it
' points at the returned array instead of the temporary buffer passed to GDI+.
Public Function GDIPlus_ImageGetProperty(ByVal hImage As Long, ByVal gpPropertyID As GP_PropertyTag, ByRef dstHeader As GP_PropertyItem, ByRef dstValueBuffer() As Byte) As Boolean
    
    GDIPlus_ImageGetProperty = False
    
    Dim tmpReturn As GP_Result, propSize As Long
    tmpReturn = GdipGetPropertyItemSize(hImage, gpPropertyID, propSize)
    
    If (tmpReturn = GP_OK) Then
        
        'Even if the call was successful, the property may not exist; check for a non-zero size.
        ' (Note that the returned buffer contains two pieces of data: a 16-byte header, and (n) bytes
        '  of actual property value information.  If (n) = 0, the value buffer is empty, and we don't
        '  want to return anything whatsoever, despite the presence of a property header.)
        GDIPlus_ImageGetProperty = (propSize > 16)
        If GDIPlus_ImageGetProperty Then
        
            Dim tmpBuffer() As Byte
            ReDim tmpBuffer(0 To propSize - 1) As Byte
            tmpReturn = GdipGetPropertyItem(hImage, gpPropertyID, propSize, VarPtr(tmpBuffer(0)))
            GDIPlus_ImageGetProperty = (tmpReturn = GP_OK)
            
            If GDIPlus_ImageGetProperty Then
            
                'The returned buffer is formatted in a unique way.  The first 16-bytes are a GP_PropertyItem header;
                ' followed by the actual property value as a stream of raw bytes (whose interpretation varies based
                ' on the property type defined by the header).  As a convenience, let's parse the raw buffer into
                ' separate, usable chunks, while also performing some failsafe checks on the returned data.
                
                'First, pull out the header
                CopyMemoryStrict VarPtr(dstHeader), VarPtr(tmpBuffer(0)), 16
                
                'Make sure the header type matches the property we requested, and make sure the data pointer also
                ' points at the temporary buffer *immediately* following the header.  (If it doesn't, something weird
                ' is afoot, and I'd like to figure out wtf is happening.)
                If (dstHeader.propID <> gpPropertyID) Then InternalGDIPlusError "GdipGetPropertyItem returned wrong property?", dstHeader.propID & " vs " & gpPropertyID
                If (dstHeader.propValue <> VarPtr(tmpBuffer(16))) Then InternalGDIPlusError "GdipGetPropertyItem returned a crazy pointer?", dstHeader.propValue & " vs " & VarPtr(tmpBuffer(16))
                If (dstHeader.propLength <> (propSize - 16)) Then InternalGDIPlusError "GdipGetPropertyItem returned a strangely sized buffer?", dstHeader.propLength & " vs " & propSize
                
                'Now we can size the value buffer and copy the relevant property value bytes into it.
                GDIPlus_ImageGetProperty = (dstHeader.propLength > 0)
                If GDIPlus_ImageGetProperty Then
                    
                    ReDim dstValueBuffer(0 To dstHeader.propLength - 1) As Byte
                    CopyMemoryStrict VarPtr(dstValueBuffer(0)), dstHeader.propValue, dstHeader.propLength
                    
                    'Cheat and overwrite the header's pointer with a new pointer to the value buffer we're returning
                    dstHeader.propValue = VarPtr(dstValueBuffer(0))
                    
                Else
                    InternalGDIPlusError "GdipGetPropertyItem returned an invalid property length", CStr(dstHeader.propLength)
                End If
            
            Else
                InternalGDIPlusError vbNullString, vbNullString, tmpReturn
            End If
        
        Else
            InternalGDIPlusError "GdipGetPropertyItem returned a propsize less than 16", propSize
        End If
    
    'It's totally okay for an image to not provide a given property.  This is not an error (or even a warning),
    ' so don't report it.
    End If
    
End Function

Public Function GDIPlus_ImageGetResolution(ByVal hImage As Long, ByRef dstHResolution As Single, ByRef dstVResolution As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImageHorizontalResolution(hImage, dstHResolution)
    If (tmpReturn = GP_OK) Then
        tmpReturn = GdipGetImageVerticalResolution(hImage, dstVResolution)
        GDIPlus_ImageGetResolution = (tmpReturn = GP_OK)
    Else
        GDIPlus_ImageGetResolution = False
    End If
End Function

Public Function GDIPlus_ImageLockBits(ByVal hImage As Long, ByRef srcRect As RectL, ByRef srcCopyData As GP_BitmapData, ByVal lockFlags As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapLockBits(hImage, srcRect, lockFlags, dstPixelFormat, srcCopyData)
    GDIPlus_ImageLockBits = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_ImageRotateFlip(ByVal hImage As Long, ByVal typeOfRotateFlip As GP_RotateFlip) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipImageRotateFlip(hImage, typeOfRotateFlip)
    GDIPlus_ImageRotateFlip = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Save a surface to a VB byte array.  The destination array *must* be dynamic, and does not need to be dimensionsed.
' (It will be auto-dimensioned correctly by this function.)  As with saving to file, note that the only export property
' currently supported is JPEG quality; other properties (like PNG compression level) are not exposed by GDI+.
Public Function GDIPlus_ImageSaveToArray(ByVal hImage As Long, ByRef dstArray() As Byte, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 85) As Boolean
        
    On Error GoTo GDIPlusSaveError
    
    'GDI+ uses GUIDs to define image export encoders; retrieve the relevant encoder GUID now
    Dim exporterGUID(0 To 15) As Byte
    If GetEncoderGUIDForPd2dFormat(dstFileFormat, VarPtr(exporterGUID(0))) Then
    
        'Like export format, GDI+ also uses GUIDs to define export properties.  If multiple encoder parameters
        ' are in use, these need to be merged into sequential order (because GDI+ only takes a pointer).
        ' pd2D does not currently cover this use-case; it always assumes there are only 0 or 1 parameters in use.
        ' To use multiple parameters, you would need copy the first GP_EncoderParameters entry into the
        ' fullEncoderParams() array, like normal, but with the Count value set to the number of parameters.
        ' Then, you would need to copy subsequent parameters into place *after* it.  (But *only* the parameters,
        ' not additional "Count" values.)
        '
        'Look at PhotoDemon's source code for an example of how to do this.
        Dim paramsInUse As Boolean: paramsInUse = False
        Dim tmpEncoderParams As GP_EncoderParameters, tmpConstString As String
        
        If (dstFileFormat = P2_FFE_JPEG) Then
            
            paramsInUse = True
            
            tmpEncoderParams.EP_Count = 1
            With tmpEncoderParams.EP_Parameter
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                tmpConstString = GP_EP_Quality
                CLSIDFromString StrPtr(tmpConstString), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(jpegQuality)
            End With
            
        End If
        
        'Prep an IStream to receive the export.  Note that we deliberately mark the stream as "free on release",
        ' which spares us from manually releasing the stream's contents.  (They will be auto-freed when the tmpStream
        ' object goes out of scope.)
        Dim tmpStream As Long
        CreateStreamOnHGlobal 0&, 1&, VarPtr(tmpStream)
        
        'Perform the export
        Dim tmpReturn As GP_Result
        If paramsInUse Then
            tmpReturn = GdipSaveImageToStream(hImage, tmpStream, VarPtr(exporterGUID(0)), VarPtr(tmpEncoderParams))
        Else
            tmpReturn = GdipSaveImageToStream(hImage, tmpStream, VarPtr(exporterGUID(0)), 0&)
        End If
        
        If (tmpReturn = GP_OK) Then
        
            'We now need to copy the contents of the stream into a VB array
            Dim tmpHMem As Long, hMemSize As Long
            If (GetHGlobalFromStream(tmpStream, tmpHMem) = 0) Then
                hMemSize = GlobalSize(tmpHMem)
                If (hMemSize <> 0) Then
                
                    Dim lockedMem As Long
                    lockedMem = GlobalLock(tmpHMem)
                    If (lockedMem <> 0) Then
                        ReDim dstArray(0 To hMemSize - 1) As Byte
                        CopyMemoryStrict VarPtr(dstArray(0)), lockedMem, hMemSize
                        GlobalUnlock lockedMem
                        GDIPlus_ImageSaveToArray = True
                    End If
                    
                End If
            End If
            
        Else
            GDIPlus_ImageSaveToArray = False
            InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToArray() failed; additional details follow"
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
    
    Else
        InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToArray() failed; no encoder found for that image format"
    End If
    
    Exit Function
    
GDIPlusSaveError:
    InternalGDIPlusError "Image was not saved", "A VB error occurred inside GDIPlus_ImageSaveToFile: " & Err.Description
    GDIPlus_ImageSaveToArray = False
End Function

'Save a surface to file.  The only property currently supported is JPEG quality; other properties are set automatically by GDI+.
Public Function GDIPlus_ImageSaveToFile(ByVal hImage As Long, ByVal dstFilename As String, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 85) As Boolean
        
    On Error GoTo GDIPlusSaveError
    
    'GDI+ uses GUIDs to define image export encoders; retrieve the relevant encoder GUID now
    Dim exporterGUID(0 To 15) As Byte
    If GetEncoderGUIDForPd2dFormat(dstFileFormat, VarPtr(exporterGUID(0))) Then
    
        'Like export format, GDI+ also uses GUIDs to define export properties.  If multiple encoder parameters
        ' are in use, these need to be merged into sequential order (because GDI+ only takes a pointer).
        ' pd2D does not currently cover this use-case; it always assumes there are only 0 or 1 parameters in use.
        ' To use multiple parameters, you would need copy the first GP_EncoderParameters entry into the
        ' fullEncoderParams() array, like normal, but with the Count value set to the number of parameters.
        ' Then, you would need to copy subsequent parameters into place *after* it.  (But *only* the parameters,
        ' not additional "Count" values.)
        '
        'Look at PhotoDemon's source code for an example of how to do this.
        Dim paramsInUse As Boolean: paramsInUse = False
        Dim tmpEncoderParams As GP_EncoderParameters, tmpConstString As String
        
        If (dstFileFormat = P2_FFE_JPEG) Then
            
            paramsInUse = True
            
            tmpEncoderParams.EP_Count = 1
            With tmpEncoderParams.EP_Parameter
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                tmpConstString = GP_EP_Quality
                CLSIDFromString StrPtr(tmpConstString), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(jpegQuality)
            End With
            
        End If
        
        'Perform the export and return
        Dim tmpReturn As GP_Result
        If paramsInUse Then
            tmpReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(exporterGUID(0)), VarPtr(tmpEncoderParams))
        Else
            tmpReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(exporterGUID(0)), 0&)
        End If
        
        If (tmpReturn = GP_OK) Then
            GDIPlus_ImageSaveToFile = True
        Else
            GDIPlus_ImageSaveToFile = False
            InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToFile() failed to save " & dstFilename & "; additional details follow"
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
    
    Else
        InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToFile() failed to save " & dstFilename & "; no encoder found for that image format"
    End If
    
    Exit Function
    
GDIPlusSaveError:
    InternalGDIPlusError "Image was not saved", "A VB error occurred inside GDIPlus_ImageSaveToFile: " & Err.Description
    GDIPlus_ImageSaveToFile = False
End Function

'When exporting images, we need to find the unique GUID for a given exporter.  Matching via mimetype is a
' straightforward way to do this, and is the recommended solution from MSDN (see https://msdn.microsoft.com/en-us/library/ms533843(v=vs.85).aspx)
Private Function GetEncoderGUIDForPd2dFormat(ByVal srcFormat As PD_2D_FileFormatExport, ByVal ptrToDstGuid As Long) As Boolean
    
    GetEncoderGUIDForPd2dFormat = False
    
    'Generate a matching mimetype for the given format
    Dim srcMimetype As String
    Select Case srcFormat
        Case P2_FFE_BMP
            srcMimetype = "image/bmp"
        Case P2_FFE_GIF
            srcMimetype = "image/gif"
        Case P2_FFE_JPEG
            srcMimetype = "image/jpeg"
        Case P2_FFE_PNG
            srcMimetype = "image/png"
        Case P2_FFE_TIFF
            srcMimetype = "image/tiff"
        Case Else
            srcMimetype = vbNullString
    End Select
    
    If (LenB(srcMimetype) <> 0) Then
        
        'Start by retrieving the number of encoders, and the size of the full encoder list
        Dim numOfEncoders As Long, sizeOfEncoders As Long
        If (GdipGetImageEncodersSize(numOfEncoders, sizeOfEncoders) = GP_OK) Then
            
            If (numOfEncoders > 0) And (sizeOfEncoders > 0) Then
            
                Dim encoderBuffer() As Byte
                Dim tmpCodec As GP_ImageCodecInfo
                
                'Hypothetically, we could probably pull the encoder list directly into a GP_ImageCodecInfo() array,
                ' but I haven't tested to see if the byte values of the encoder sizes are exact.  To avoid any problems,
                ' let's just dump the return into a byte array, then parse out what we need as we go.
                ReDim encoderBuffer(0 To sizeOfEncoders - 1) As Byte
                If (GdipGetImageEncoders(numOfEncoders, sizeOfEncoders, VarPtr(encoderBuffer(0))) = GP_OK) Then
                
                    'Iterate through the encoder list, searching for a match
                    Dim i As Long
                    For i = 0 To numOfEncoders - 1
                    
                        'Extract this codec
                        CopyMemoryStrict VarPtr(tmpCodec), VarPtr(encoderBuffer(0)) + LenB(tmpCodec) * i, LenB(tmpCodec)
                        
                        'Compare mimetypes
                        If Strings.StringsEqual(Strings.StringFromCharPtr(tmpCodec.IC_MimeType, True), srcMimetype, True) Then
                            GetEncoderGUIDForPd2dFormat = True
                            CopyMemoryStrict ptrToDstGuid, VarPtr(tmpCodec.IC_ClassID(0)), 16&
                            Exit For
                        End If
                        
                    Next i
                
                End If
                
            End If
        End If
        
    End If

End Function

'Debug only: list the decoders available on this system.  Users may have additional decoders installed,
' besides those offered by default (JPEG, PNG, etc).
'
'NOTE: as of Win 10, this function is disabled.  Testing shows that it's basically useless; GDI+ extenders
' don't seem to exist in the wild.  Everyone writes extensions for WIC now (as they should).
'Public Sub DEBUG_ListGdipDecoders()
'
'    Dim numDecoders As Long, sizeEncodersBytes As Long
'    If (GdipGetImageDecodersSize(numDecoders, sizeEncodersBytes) = GP_OK) Then
'
'        'For reasons I don't fully understand, the sizeEncodersBytes value is often significantly larger
'        ' than the size you'd expect given the number of codecs.  As such, declare our array safely.
'        Dim tmpExampleDec As GP_ImageCodecInfo
'
'        Dim safeNumEncoders As Long
'        safeNumEncoders = (sizeEncodersBytes \ LenB(tmpExampleDec)) + 1
'
'        Dim encList() As GP_ImageCodecInfo
'        ReDim encList(0 To safeNumEncoders - 1) As GP_ImageCodecInfo
'
'        If (GdipGetImageDecoders(numDecoders, sizeEncodersBytes, VarPtr(encList(0))) = GP_OK) Then
'
'            Debug.Print "Found " & CStr(numDecoders) & " GDI+ decoders on this PC.  The list includes:"
'            Dim i As Long
'            For i = 0 To numDecoders - 1
'                Debug.Print vbTab & CStr(i + 1) & ": " & Strings.StringFromCharPtr(encList(i).IC_CodecName, True)
'            Next i
'
'        Else
'            Debug.Print "WARNING: GDI+ failed to retrieve decoder list; has the library been initialized correctly?"
'        End If
'
'    Else
'        Debug.Print "WARNING: GDI+ returned no valid decoders; has the library been initialized correctly?"
'    End If
'
'End Sub

Public Function GDIPlus_ImageUnlockBits(ByVal hImage As Long, ByRef srcCopyData As GP_BitmapData) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapUnlockBits(hImage, srcCopyData)
    GDIPlus_ImageUnlockBits = (tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Convert an EMF or WMF to the new EMF+ format.  Note that this is done in-memory, and the source file is not touched.
' Conversion allows us to render the metafile with antialiasing, alpha bytes, and more.
' REQUIRES GDI+ v1.1 (Win 7 or later only; conditionally available on Vista if explicitly requested via manifest)
'
'If successful, this function will generate a new handle.  It *must* be freed separately from the old handle!
Public Function GDIPlus_ImageUpgradeMetafile(ByVal hImage As Long, ByVal srcGraphicsForConvertSettings As Long, ByRef dstNewMetafile As Long) As Boolean
    
    dstNewMetafile = 0
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipConvertToEmfPlus(srcGraphicsForConvertSettings, hImage, ByVal 0&, GP_MT_EmfDual, 0&, dstNewMetafile)
    
    GDIPlus_ImageUpgradeMetafile = (tmpReturn = GP_OK)
    If Not GDIPlus_ImageUpgradeMetafile Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
End Function

Public Sub GDIPlus_LineGradientSetGamma(ByVal hGradientBrush As Long, ByVal newGamma As Boolean)
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetLineGammaCorrection(hGradientBrush, IIf(newGamma, 1, 0))
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Sub

Public Sub GDIPlus_PathGradientSetGamma(ByVal hGradientBrush As Long, ByVal newGamma As Boolean)
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetPathGradientGammaCorrection(hGradientBrush, IIf(newGamma, 1, 0))
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Sub

'In 2025.4, I added the ability for users to add their own fonts at run-time.  To be available to GDI+,
' these must be maintained in a private font collection.
Public Function GDIPlus_AddRuntimeFont(ByRef srcFontFile As String) As Boolean
    
    Dim gpResult As GP_Result
    
    'Ensure the central GDI+ font collection exists
    If (m_UserFontCollection = 0) Then
        gpResult = GdipNewPrivateFontCollection(m_UserFontCollection)
        If (gpResult <> GP_OK) Then InternalGDIPlusError "Couldn't create GDI+ font collection", vbNullString, gpResult
    End If
    
    'We can only add fonts if the central collection was successfully created
    If (m_UserFontCollection <> 0) Then
        gpResult = GdipPrivateAddFontFile(m_UserFontCollection, StrPtr(srcFontFile))
        If (gpResult <> GP_OK) Then InternalGDIPlusError "Couldn't add font", srcFontFile, gpResult
    End If
    
End Function

Public Function GDIPlus_GetUserFontCollection() As Long
    GDIPlus_GetUserFontCollection = m_UserFontCollection
End Function

'At termination, free all GDI+ fonts
Public Sub GDIPlus_ReleaseRuntimeFonts()
        
    On Error GoTo BadGdipBehavior
    
    If (m_UserFontCollection <> 0) Then
        
        'NOTE: if any private font families were freed *before* freeing the private font collection,
        ' this function will crash.  Do *not* manually free private font families - they are freed
        ' automatically by GDI+ when the parent collection is freed.
        Dim gpResult As GP_Result
        gpResult = GdipDeletePrivateFontCollection(m_UserFontCollection)
        If (gpResult = GP_OK) Then
            m_UserFontCollection = 0
        Else
            InternalGDIPlusError "Couldn't free GDI+ font collection", vbNullString, gpResult
        End If
        
    End If
    
BadGdipBehavior:
    
End Sub
