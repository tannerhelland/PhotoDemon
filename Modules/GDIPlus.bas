Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright 2012-2016 by Tanner Helland
'Created: 1/September/12
'Last updated: 26/June/16
'Last update: add more integer-specific rendering functions
'
'This interface provides a means for interacting with various GDI+ features.  GDI+ was originally used as a fallback for image loading
' and saving if the FreeImage DLL was not found, but over time it has become more and more integrated into PD.  As of version 6.0, GDI+
' is used for a number of specialized tasks, including viewport rendering of 32bpp images, regional blur of selection masks, antialiased
' lines and circles on various dialogs, and more.
'
'Note that - by design - some enums in this class differ subtly from the actual GDI+ enums.  This is a deliberate decision
' to make certain enums play more nicely with other imaging libraries and/or features.  PD handles translation between the
' correct enums as necessary.
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

'As of 2016, this module is undergoing massive reorganization.  Enums, constants, and functions that have been migrated
' to the new (clean) format are placed in this top section.

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

Private Enum GP_DebugEventLevel
    GP_DebugEventLevelFatal = 0
    GP_DebugEventLevelWarning = 1
End Enum

#If False Then
    Private Const GP_DebugEventLevelFatal = 0, GP_DebugEventLevelWarning = 1
#End If

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

Public Enum GP_BrushType        'IMPORTANT NOTE!  This enum is *not* the same as PD's internal 2D brush modes!
    GP_BT_SolidColor = 0
    GP_BT_HatchFill = 1
    GP_BT_TextureFill = 2
    GP_BT_PathGradient = 3
    GP_BT_LinearGradient = 4
End Enum

#If False Then
    Private Const GP_BT_SolidColor = 0, GP_BT_HatchFill = 1, GP_BT_TextureFill = 2, GP_BT_PathGradient = 3, GP_BT_LinearGradient = 4
#End If

'Coloar adjustments are handled internally, at present, so we don't need to expose them to other objects
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

Public Enum GP_PenAlignment
    GP_PA_Center = 0&
    GP_PA_Inset = 1&
End Enum

#If False Then
    Private Const GP_PA_Center = 0&, GP_PA_Inset = 1&
#End If

'GDI+ pixel format IDs use a bitfield system:
' [0, 7] = format index
' [8, 15] = pixel size (in bits)
' [16, 23] = flags
' [24, 31] = reserved (current unused)

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
    GP_PT_Artist = &H13B
    GP_PT_BitsPerSample = &H102
    GP_PT_CellHeight = &H109
    GP_PT_CellWidth = &H108
    GP_PT_ChrominanceTable = &H5091
    GP_PT_ColorMap = &H140
    GP_PT_ColorTransferFunction = &H501A
    GP_PT_Compression = &H103
    GP_PT_Copyright = &H8298
    GP_PT_DateTime = &H132
    GP_PT_DocumentName = &H10D
    GP_PT_DotRange = &H150
    GP_PT_EquipMake = &H10F
    GP_PT_EquipModel = &H110
    GP_PT_ExifAperture = &H9202
    GP_PT_ExifBrightness = &H9203
    GP_PT_ExifCfaPattern = &HA302
    GP_PT_ExifColorSpace = &HA001
    GP_PT_ExifCompBPP = &H9102
    GP_PT_ExifCompConfig = &H9101
    GP_PT_ExifDTDigitized = &H9004
    GP_PT_ExifDTDigSS = &H9292
    GP_PT_ExifDTOrig = &H9003
    GP_PT_ExifDTOrigSS = &H9291
    GP_PT_ExifDTSubsec = &H9290
    GP_PT_ExifExposureBias = &H9204
    GP_PT_ExifExposureIndex = &HA215
    GP_PT_ExifExposureProg = &H8822
    GP_PT_ExifExposureTime = &H829A
    GP_PT_ExifFileSource = &HA300
    GP_PT_ExifFlash = &H9209
    GP_PT_ExifFlashEnergy = &HA20B
    GP_PT_ExifFNumber = &H829D
    GP_PT_ExifFocalLength = &H920A
    GP_PT_ExifFocalResUnit = &HA210
    GP_PT_ExifFocalXRes = &HA20E
    GP_PT_ExifFocalYRes = &HA20F
    GP_PT_ExifFPXVer = &HA000
    GP_PT_ExifIFD = &H8769
    GP_PT_ExifInterop = &HA005
    GP_PT_ExifISOSpeed = &H8827
    GP_PT_ExifLightSource = &H9208
    GP_PT_ExifMakerNote = &H927C
    GP_PT_ExifMaxAperture = &H9205
    GP_PT_ExifMeteringMode = &H9207
    GP_PT_ExifOECF = &H8828
    GP_PT_ExifPixXDim = &HA002
    GP_PT_ExifPixYDim = &HA003
    GP_PT_ExifRelatedWav = &HA004
    GP_PT_ExifSceneType = &HA301
    GP_PT_ExifSensingMethod = &HA217
    GP_PT_ExifShutterSpeed = &H9201
    GP_PT_ExifSpatialFR = &HA20C
    GP_PT_ExifSpectralSense = &H8824
    GP_PT_ExifSubjectDist = &H9206
    GP_PT_ExifSubjectLoc = &HA214
    GP_PT_ExifUserComment = &H9286
    GP_PT_ExifVer = &H9000
    GP_PT_ExtraSamples = &H152
    GP_PT_FillOrder = &H10A
    GP_PT_FrameDelay = &H5100
    GP_PT_FreeByteCounts = &H121
    GP_PT_FreeOffset = &H120
    GP_PT_Gamma = &H301
    GP_PT_GlobalPalette = &H5102
    GP_PT_GpsAltitude = &H6
    GP_PT_GpsAltitudeRef = &H5
    GP_PT_GpsDestBear = &H18
    GP_PT_GpsDestBearRef = &H17
    GP_PT_GpsDestDist = &H1A
    GP_PT_GpsDestDistRef = &H19
    GP_PT_GpsDestLat = &H14
    GP_PT_GpsDestLatRef = &H13
    GP_PT_GpsDestLong = &H16
    GP_PT_GpsDestLongRef = &H15
    GP_PT_GpsGpsDop = &HB
    GP_PT_GpsGpsMeasureMode = &HA
    GP_PT_GpsGpsSatellites = &H8
    GP_PT_GpsGpsStatus = &H9
    GP_PT_GpsGpsTime = &H7
    GP_PT_GpsIFD = &H8825
    GP_PT_GpsImgDir = &H11
    GP_PT_GpsImgDirRef = &H10
    GP_PT_GpsLatitude = &H2
    GP_PT_GpsLatitudeRef = &H1
    GP_PT_GpsLongitude = &H4
    GP_PT_GpsLongitudeRef = &H3
    GP_PT_GpsMapDatum = &H12
    GP_PT_GpsSpeed = &HD
    GP_PT_GpsSpeedRef = &HC
    GP_PT_GpsTrack = &HF
    GP_PT_GpsTrackRef = &HE
    GP_PT_GpsVer = &H0
    GP_PT_GrayResponseCurve = &H123
    GP_PT_GrayResponseUnit = &H122
    GP_PT_GridSize = &H5011
    GP_PT_HalftoneDegree = &H500C
    GP_PT_HalftoneHints = &H141
    GP_PT_HalftoneLPI = &H500A
    GP_PT_HalftoneLPIUnit = &H500B
    GP_PT_HalftoneMisc = &H500E
    GP_PT_HalftoneScreen = &H500F
    GP_PT_HalftoneShape = &H500D
    GP_PT_HostComputer = &H13C
    GP_PT_ICCProfile = &H8773
    GP_PT_ICCProfileDescriptor = &H302
    GP_PT_ImageDescription = &H10E
    GP_PT_ImageHeight = &H101
    GP_PT_ImageTitle = &H320
    GP_PT_ImageWidth = &H100
    GP_PT_IndexBackground = &H5103
    GP_PT_IndexTransparent = &H5104
    GP_PT_InkNames = &H14D
    GP_PT_InkSet = &H14C
    GP_PT_JPEGACTables = &H209
    GP_PT_JPEGDCTables = &H208
    GP_PT_JPEGInterFormat = &H201
    GP_PT_JPEGInterLength = &H202
    GP_PT_JPEGLosslessPredictors = &H205
    GP_PT_JPEGPointTransforms = &H206
    GP_PT_JPEGProc = &H200
    GP_PT_JPEGQTables = &H207
    GP_PT_JPEGQuality = &H5010
    GP_PT_JPEGRestartInterval = &H203
    GP_PT_LoopCount = &H5101
    GP_PT_LuminanceTable = &H5090
    GP_PT_MaxSampleValue = &H119
    GP_PT_MinSampleValue = &H118
    GP_PT_NewSubfileType = &HFE
    GP_PT_NumberOfInks = &H14E
    GP_PT_Orientation = &H112
    GP_PT_PageName = &H11D
    GP_PT_PageNumber = &H129
    GP_PT_PaletteHistogram = &H5113
    GP_PT_PhotometricInterp = &H106
    GP_PT_PixelPerUnitX = &H5111
    GP_PT_PixelPerUnitY = &H5112
    GP_PT_PixelUnit = &H5110
    GP_PT_PlanarConfig = &H11C
    GP_PT_Predictor = &H13D
    GP_PT_PrimaryChromaticities = &H13F
    GP_PT_PrintFlags = &H5005
    GP_PT_PrintFlagsBleedWidth = &H5008
    GP_PT_PrintFlagsBleedWidthScale = &H5009
    GP_PT_PrintFlagsCrop = &H5007
    GP_PT_PrintFlagsVersion = &H5006
    GP_PT_REFBlackWhite = &H214
    GP_PT_ResolutionUnit = &H128
    GP_PT_ResolutionXLengthUnit = &H5003
    GP_PT_ResolutionXUnit = &H5001
    GP_PT_ResolutionYLengthUnit = &H5004
    GP_PT_ResolutionYUnit = &H5002
    GP_PT_RowsPerStrip = &H116
    GP_PT_SampleFormat = &H153
    GP_PT_SamplesPerPixel = &H115
    GP_PT_SMaxSampleValue = &H155
    GP_PT_SMinSampleValue = &H154
    GP_PT_SoftwareUsed = &H131
    GP_PT_SRGBRenderingIntent = &H303
    GP_PT_StripBytesCount = &H117
    GP_PT_StripOffsets = &H111
    GP_PT_SubfileType = &HFF
    GP_PT_T4Option = &H124
    GP_PT_T6Option = &H125
    GP_PT_TargetPrinter = &H151
    GP_PT_ThreshHolding = &H107
    GP_PT_ThumbnailArtist = &H5034
    GP_PT_ThumbnailBitsPerSample = &H5022
    GP_PT_ThumbnailColorDepth = &H5015
    GP_PT_ThumbnailCompressedSize = &H5019
    GP_PT_ThumbnailCompression = &H5023
    GP_PT_ThumbnailCopyRight = &H503B
    GP_PT_ThumbnailData = &H501B
    GP_PT_ThumbnailDateTime = &H5033
    GP_PT_ThumbnailEquipMake = &H5026
    GP_PT_ThumbnailEquipModel = &H5027
    GP_PT_ThumbnailFormat = &H5012
    GP_PT_ThumbnailHeight = &H5014
    GP_PT_ThumbnailImageDescription = &H5025
    GP_PT_ThumbnailImageHeight = &H5021
    GP_PT_ThumbnailImageWidth = &H5020
    GP_PT_ThumbnailOrientation = &H5029
    GP_PT_ThumbnailPhotometricInterp = &H5024
    GP_PT_ThumbnailPlanarConfig = &H502F
    GP_PT_ThumbnailPlanes = &H5016
    GP_PT_ThumbnailPrimaryChromaticities = &H5036
    GP_PT_ThumbnailRawBytes = &H5017
    GP_PT_ThumbnailRefBlackWhite = &H503A
    GP_PT_ThumbnailResolutionUnit = &H5030
    GP_PT_ThumbnailResolutionX = &H502D
    GP_PT_ThumbnailResolutionY = &H502E
    GP_PT_ThumbnailRowsPerStrip = &H502B
    GP_PT_ThumbnailSamplesPerPixel = &H502A
    GP_PT_ThumbnailSize = &H5018
    GP_PT_ThumbnailSoftwareUsed = &H5032
    GP_PT_ThumbnailStripBytesCount = &H502C
    GP_PT_ThumbnailStripOffsets = &H5028
    GP_PT_ThumbnailTransferFunction = &H5031
    GP_PT_ThumbnailWhitePoint = &H5035
    GP_PT_ThumbnailWidth = &H5013
    GP_PT_ThumbnailYCbCrCoefficients = &H5037
    GP_PT_ThumbnailYCbCrPositioning = &H5039
    GP_PT_ThumbnailYCbCrSubsampling = &H5038
    GP_PT_TileByteCounts = &H145
    GP_PT_TileLength = &H143
    GP_PT_TileOffset = &H144
    GP_PT_TileWidth = &H142
    GP_PT_TransferFunction = &H12D
    GP_PT_TransferRange = &H156
    GP_PT_WhitePoint = &H13E
    GP_PT_XPosition = &H11E
    GP_PT_XResolution = &H11A
    GP_PT_YCbCrCoefficients = &H211
    GP_PT_YCbCrPositioning = &H213
    GP_PT_YCbCrSubsampling = &H212
    GP_PT_YPosition = &H11F
    GP_PT_YResolution = &H11B
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
    Colorused As Long
    ColorImportant As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQUAD
End Type

'This (stupid) type is used so we can take advantage of LSet when performing some conversions
Private Type tmpLong
    lngResult As Long
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
Private Const GP_FF_GUID_EXIF = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_Icon = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"

'Like image formats, export encoder properties are also defined by GUID.  These values come from the Win 8.1
' version of gdiplusimaging.h.  Note that some are restricted to GDI+ v1.1.
Private Const GP_EP_Compression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const GP_EP_ColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const GP_EP_ScanMethod As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const GP_EP_Version As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const GP_EP_RenderMethod As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const GP_EP_Quality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const GP_EP_Transformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const GP_EP_LuminanceTable As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const GP_EP_ChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const GP_EP_SaveFlag As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"

'REQUIRES GDI+ v1.1 OR LATER!
Private Const GP_EP_ColorSpace As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Private Const GP_EP_SaveAsCMYK As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"

'Core GDI+ functions:
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef gdipToken As Long, ByRef startupStruct As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GP_Result
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal gdipToken As Long) As GP_Result

'Object creation/destruction/property functions
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GP_Result
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As GP_Result

Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcRect As RECTL, ByVal lockMode As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat, ByRef srcBitmapData As GP_BitmapData) As GP_Result
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcBitmapData As GP_BitmapData) As GP_Result

Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal newPixelFormat As GP_PixelFormat, ByVal hSrcBitmap As Long, ByRef hDstBitmap As Long) As GP_Result
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal srcRegion As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectF As RECTF, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectL As RECTL, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcRegion As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcPath As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result

'This EMF convert function only works on Vista+!
Private Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result

Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef origGDIBitmapInfo As BITMAPINFO, ByRef srcBitmapData As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal bmpWidth As Long, ByVal bmpHeight As Long, ByVal bmpStride As Long, ByVal bmpPixelFormat As GP_PixelFormat, ByRef Scan0 As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef dstGraphics As Long) As GP_Result
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal bHatchStyle As GP_PatternStyle, ByVal bForeColor As Long, ByVal bBackColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef dstImageAttributes As Long) As GP_Result
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef firstPoint As POINTFLOAT, ByRef secondPoint As POINTFLOAT, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal brushWrapMode As GP_WrapMode, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As GP_WrapMode, ByRef dstLineGradientBrush As Long) As GP_Result
Private Declare Function GdipCreateMatrix Lib "gdiplus" (ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal pathFillMode As GP_FillMode, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As GP_Result
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal srcColor As Long, ByVal srcWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreatePenFromBrush Lib "gdiplus" Alias "GdipCreatePen2" (ByVal srcBrush As Long, ByVal penWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreateRegion Lib "gdiplus" (ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal hPath As Long, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRect Lib "gdiplus" (ByRef srcRect As RECTF, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRgnData Lib "gdiplus" (ByVal ptrToRegionData As Long, ByVal sizeOfRegionData As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal srcColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal hImage As Long, ByVal textureWrapMode As GP_WrapMode, ByRef dstTexture As Long) As GP_Result

Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GP_Result
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GP_Result
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GP_Result
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImageAttributes As Long) As GP_Result

Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As GP_Result
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

Private Declare Function GdipGetClip Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef dstHeight As Long) As GP_Result
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstHResolution As Single) As GP_Result
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, ByRef dstPixelFormat As GP_PixelFormat) As GP_Result
Private Declare Function GdipGetImageType Lib "gdiplus" (ByVal srcImage As Long, ByRef dstImageType As GP_ImageType) As GP_Result
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstVResolution As Single) As GP_Result
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Long) As GP_Result
Private Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByRef dstFillRule As GP_FillMode) As GP_Result
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTF, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTL, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPenColor Lib "gdiplus" (ByVal hPen As Long, ByRef dstPARGBColor As Long) As GP_Result
Private Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal hPen As Long, ByRef dstCap As GP_DashCap) As GP_Result
Private Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByRef dstDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByRef dstMiterLimit As Single) As GP_Result
Private Declare Function GdipGetPenMode Lib "gdiplus" (ByVal hPen As Long, ByRef dstPenMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByRef dstWidth As Single) As GP_Result
Private Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByVal srcPropertySize As Long, ByVal ptrToDstBuffer As Long) As GP_Result
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByRef dstPropertySize As Long) As GP_Result
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectF As RECTF) As GP_Result
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectL As RECTL) As GP_Result
Private Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstX As Long, ByRef dstY As Long) As GP_Result
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef dstColor As Long) As GP_Result
Private Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByRef dstWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDstGuid As Long) As GP_Result
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rotateFlipType As GP_RotateFlip) As GP_Result
Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal srcRegion1 As Long, ByVal srcRegion2 As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal ptrSrcFilename As Long, ByRef dstGdipImage As Long) As GP_Result
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal srcIStream As Long, ByRef dstGdipImage As Long) As GP_Result

Private Declare Function GdipResetClip Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToFilename As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal hImage As Long, ByVal dstIStream As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSetClipRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal hGraphics As Long, ByVal hRegion As Long, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingMode As GP_CompositingMode) As GP_Result
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal newWrapMode As GP_WrapMode, ByVal argbOfClampMode As Long, ByVal bClampMustBeZero As Long) As GP_Result
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal typeOfAdjustment As GP_ColorAdjustType, ByVal enableSeparateAdjustmentFlag As Long, ByVal ptrToColorMatrix As Long, ByVal ptrToGrayscaleMatrix As Long, ByVal extraColorMatrixFlags As GP_ColorMatrixFlags) As GP_Result
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoints As POINTFLOAT) As GP_Result
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal pathFillMode As GP_FillMode) As GP_Result
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal hPen As Long, ByVal pARGBColor As Long) As GP_Result
Private Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal hPen As Long, ByVal newCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal newDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByVal endCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal startCap As GP_LineCap, ByVal endCap As GP_LineCap, ByVal dashCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal newLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByVal newMiterLimit As Single) As GP_Result
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal hPen As Long, ByVal penMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByVal startCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByVal penWidth As Single) As GP_Result
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Long, ByVal y As Long) As GP_Result
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByVal newColor As Long) As GP_Result
Private Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal hPath As Long, ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As GP_Result
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformationMatrix As Long, ByVal allowableError As Single) As GP_Result

'Non-GDI+ helper functions:
Private Declare Function CLSIDFromString Lib "ole32" (ByVal ptrToGuidString As Long, ByVal ptrToByteArray As Long) As Long
Private Declare Function CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByVal ptrToDstStream As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal srcIStream As Long, ByRef dstHGlobal As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal ptrToString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal ptrToString As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long) As Long
Private Declare Function StringFromCLSID Lib "ole32" (ByVal ptrToGuid As Long, ByRef ptrToDstString As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal srcWCharPtr As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal srcAnsiPtr As Long, ByVal srcLength As Long) As String

'Internally cached values:

'Startup values
Private m_GDIPlusToken As Long, m_GDIPlus11Available As Boolean

'Some GDI+ functions require world transformation data.  This dummy graphics container is used to host any such transformations.
' It is created when GDI+ is initialized, and destroyed when GDI+ is released.  To be a good citizen, please undo any world transforms
' before a function releases.  This ensures that subsequent functions are not messed up.
Private m_TransformDIB As pdDIB, m_TransformGraphics As Long

'To modify opacity in GDI+, an image attributes matrix is used.  Rather than recreating one every time an alpha operation is required,
' we simply create a default identity matrix at initialization, then re-use it as necessary.
Private m_AttributesMatrix() As Single

'***************************************************************************

'Old declarations and descriptions follow.  These need to be reworked into something coherent, but it's a
' slog of a process...

Public Enum GDIPlusImageFormat
    [ImageBMP] = 0
    [ImageGIF] = 1
    [ImageJPEG] = 2
    [ImagePNG] = 3
    [ImageTIFF] = 4
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

'EMFs can be converted between various formats.  GDI+ prefers "EMF+", which supports GDI+ primitives as well
Private Enum MetafileType
   MetafileTypeInvalid            'Invalid metafile
   MetafileTypeWmf                'Standard WMF
   MetafileTypeWmfPlaceable       'Placeable WMF
   MetafileTypeEmf                'EMF (not EMF+)
   MetafileTypeEmfPlusOnly        'EMF+ without dual down-level records
   MetafileTypeEmfPlusDual        'EMF+ with dual down-level records
End Enum

Private Enum EMFType
    EmfTypeEmfOnly = MetafileTypeEmf               'no EMF+  only EMF
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly   'no EMF  only EMF+
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual   'both EMF+ and EMF
End Enum

Private Type clsid
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
    ClassID           As clsid
    FormatID          As clsid
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
    Guid           As clsid
    NumberOfValues As Long
    encType           As EncoderParameterValueType
    Value          As Long
End Type

Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type

Public Enum rotateFlipType
   RotateNoneFlipNone = 0
   Rotate90FlipNone = 1
   Rotate180FlipNone = 2
   Rotate270FlipNone = 3

   RotateNoneFlipX = 4
   Rotate90FlipX = 5
   Rotate180FlipX = 6
   Rotate270FlipX = 7

   RotateNoneFlipY = Rotate180FlipX
   Rotate90FlipY = Rotate270FlipX
   Rotate180FlipY = RotateNoneFlipX
   Rotate270FlipY = Rotate90FlipX

   RotateNoneFlipXY = Rotate180FlipNone
   Rotate90FlipXY = Rotate270FlipNone
   Rotate180FlipXY = RotateNoneFlipNone
   Rotate270FlipXY = Rotate90FlipNone
End Enum

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
Private Const PixelFormatAlpha = &H40000
Private Const PixelFormatPremultipliedAlpha = &H80000
Private Const PixelFormat32bppCMYK = &H200F

Private Const PixelFormat1bppIndexed = &H30101
Private Const PixelFormat4bppIndexed = &H30402
Private Const PixelFormat8bppIndexed = &H30803
Private Const PixelFormat16bppGreyscale = &H101004
Private Const PixelFormat16bppRGB555 = &H21005
Private Const PixelFormat16bppRGB565 = &H21006
Private Const PixelFormat16bppARGB1555 = &H61007
Private Const PixelFormat24bppRGB = &H21808
Private Const PixelFormat32bppRGB = &H22009
Private Const PixelFormat32bppARGB = &H26200A
Public Const PixelFormat32bppPARGB = &HE200B
Private Const PixelFormat48bppRGB = &H10300C
Private Const PixelFormat64bppARGB = &H34400D
Private Const PixelFormat64bppPARGB = &H1C400E

'Now that PD supports the loading of ICC profiles, we use this constant to retrieve it
Private Const PropertyTagICCProfile As Long = &H8773&

'Orientation tag is used to auto-rotate incoming JPEGs
Private Const PropertyTagOrientation As Long = &H112&

'LockBits constants
Private Const ImageLockModeRead = &H1
Private Const ImageLockModeWrite = &H2
Private Const ImageLockModeUserInputBuf = &H4

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap = 1
    ColorAdjustTypeBrush = 2
    ColorAdjustTypePen = 3
    ColorAdjustTypeText = 4
    ColorAdjustTypeCount = 5
    ColorAdjustTypeAny = 6
End Enum

#If False Then
    Const ColorAdjustTypeDefault = 0, ColorAdjustTypeBitmap = 1, ColorAdjustTypeBrush = 2, ColorAdjustTypePen = 3, ColorAdjustTypeText = 4
    Const ColorAdjustTypeCount = 5, ColorAdjustTypeAny = 6
#End If

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

#If False Then
    Const ColorMatrixFlagsDefault = 0, ColorMatrixFlagsSkipGrays = 1, ColorMatrixFlagsAltGray = 2
#End If

'GDI+ image properties
Private Type PropertyItem
    propID As Long              ' ID of this property
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

'Load image from file, process said file, etc.
Private Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal srcFilename As String, ByRef gpImage As Long) As Long
Private Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal gpBitmap As Long, ByRef gpFlags As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gpBitmap As Long, hBmpReturn As Long, ByVal RGBABackground As Long) As GP_Result
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, ByRef imgWidth As Single, ByRef imgHeight As Single) As Long
Private Declare Function GdipGetDC Lib "gdiplus" (ByVal mGraphics As Long, ByRef hDC As Long) As Long
Private Declare Function GdipReleaseDC Lib "gdiplus" (ByVal mGraphics As Long, ByVal hDC As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, ByRef hGraphics As Long) As GP_Result
Private Declare Function GdipCreateMetafileFromFile Lib "gdiplus" (ByVal srcFilePtr As Long, ByRef hMetafile As Long) As GP_Result
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal hGraphics As Long, ByVal lColor As Long) As GP_Result
Private Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal hMetafile As Long, ByVal metafileRasterizationLimitDpi As Long) As GP_Result

'Note: only supported in GDI+ v1.1!
Private Declare Function GdipConvertToEmfPlusToFile Lib "gdiplus" (ByVal refGraphics As Long, ByVal metafilePtr As Long, ByRef conversionSuccess As Long, ByVal filenamePointer As Long, ByVal typeOfEMF As EMFType, ByVal descriptionPointer As Long, ByRef out_metafile_ptr As Long) As Long

'OleCreatePictureIndirect is used to convert GDI+ images to VB's preferred StdPicture
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PictDesc, rIID As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

'CopyMemory
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal cb As Long) As Long

'GDI+ calls related to drawing lines and various shapes
'Private Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal srcGraphics As Long, ByRef dstBitmap As Long) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef mEffect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal mEffect As Long, ByRef eParams As Any, ByVal Size As Long) As Long
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal mEffect As Long) As Long
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal mGraphics As Long, ByVal mImage As Long, ByRef iSource As RECTF, ByVal xForm As Long, ByVal mEffect As Long, ByVal mImageAttributes As Long, ByVal srcUnit As Long) As Long
Private Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef mMatrix As Long) As Long
Private Declare Function GdipDrawCurve Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointFloatArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointLongArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipDrawCurve3 Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointFloatArrayPtr As Long, ByVal nPoints As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipDrawCurve3I Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointLongArrayPtr As Long, ByVal nPoints As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipDrawClosedCurve Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointFloatArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipDrawClosedCurveI Lib "gdiplus" (ByVal mGraphics As Long, ByVal hPen As Long, ByVal pointLongArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipFillClosedCurve Lib "gdiplus" (ByVal mGraphics As Long, ByVal hBrush As Long, ByVal pointFloatArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal mGraphics As Long, ByVal hBrush As Long, ByVal pointLongArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipFillPolygon2 Lib "gdiplus" (ByVal mGraphics As Long, ByVal hBrush As Long, ByVal pointFloatArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal mGraphics As Long, ByVal hBrush As Long, ByVal pointLongArrayPtr As Long, ByVal nPoints As Long) As Long
Private Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphics As Long, ByRef boolResult As Long) As Long
Private Declare Function GdipIsVisibleRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal hGraphics As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipSetImageAttributesToIdentity Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal clrAdjType As ColorAdjustType) As Long

'Transforms
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal mGraphics As Long, ByVal Angle As Single, ByVal order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal mGraphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long

Private Type BlurParams
    bRadius As Single
    ExpandEdge As Long
End Type

'Information about image pixel data
Private Type BitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type


'Use GDI+ to resize a DIB.  (Technically, to copy a resized portion of a source image into a destination image.)
' The call is formatted similar to StretchBlt, as it used to replace StretchBlt when working with 32bpp data.
' FOR FUTURE REFERENCE: after a bunch of profiling on my Win 7 PC, I can state with 100% confidence that
' the HighQualityBicubic interpolation mode is actually the fastest mode for downsizing 32bpp images.  I have no idea
' why this is, but many, many iterative tests confirmed it.  Stranger still, in descending order after that, the fastest
' algorithms are: HighQualityBilinear, Bilinear, Bicubic.  Regular bicubic interpolation is some 4x slower than the
' high quality mode!!
Public Function GDIPlusResizeDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal interpolationType As GP_InterpolationMode) As Boolean

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
        GdipSetPixelOffsetMode hGdipGraphics, GP_POM_HighSpeed
        
        'Perform the resize
        If GdipDrawImageRectRectI(hGdipGraphics, hGdipBitmap, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle) <> 0 Then
            GDIPlusResizeDIB = False
        End If
        
        'Release our image attributes object
        GdipDisposeImageAttributes imgAttributesHandle
        
    Else
        GDIPlusResizeDIB = False
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics hGdipGraphics
    GdipDisposeImage hGdipBitmap
    
    'GDI+ draw functions always result in a premultiplied image
    dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Function

'Simpler rotate/flip function, and limited to the constants specified by the enum.
Public Function GDIPlusRotateFlipDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotationType As rotateFlipType) As Boolean

    GDIPlusRotateFlipDIB = True
    
    'We need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    Dim tBitmap As Long
    GetGdipBitmapHandleFromDIB tBitmap, srcDIB
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Apply the rotation
    GdipImageRotateFlip tBitmap, rotationType
    
    'Resize the target DIB
    Dim newWidth As Long, newHeight As Long
    GdipGetImageWidth tBitmap, newWidth
    GdipGetImageHeight tBitmap, newHeight
    
    dstDIB.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth, 0
    
    'Obtain a GDI+ handle to the target DIB
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    'Render the rotated image
    GdipDrawImage iGraphics, tBitmap, 0, 0
    
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics iGraphics
    GdipDisposeImage tBitmap
    
End Function

'Use GDI+ to rotate a DIB.  (Technically, to copy a rotated portion of a source image into a destination image.)
' The function currently expects the rotation to occur around the center point of the source image.  Unlike the various
' size interaction calls in this module, all (x,y) coordinate pairs refer to the CENTER of the image, not the top-left
' corner.  This was a deliberate decision to make copying rotated data easier.
Public Function GDIPlusRotateDIB(ByRef dstDIB As pdDIB, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcDIB As pdDIB, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rotationAngle As Single, ByVal interpolationType As GP_InterpolationMode, Optional ByVal wrapModeForEdges As GP_WrapMode = GP_WM_TileFlipXY) As Boolean

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer

    GDIPlusRotateDIB = True

    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, srcDIB
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(iGraphics, interpolationType) = GP_OK Then
    
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        GdipSetImageAttributesWrapMode imgAttributesHandle, wrapModeForEdges, 0&, 0&
        
        'To improve performance, explicitly request high-speed alpha compositing operation
        GdipSetCompositingQuality iGraphics, GP_CQ_AssumeLinear
        
        'PixelOffsetMode doesn't seem to affect rendering speed more than < 5%, but I did notice a slight
        ' improvement from explicitly requesting HighQuality mode - so why not leave it?
        GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality
    
        'Lock the incoming angle to something in the range [-360, 360]
        'rotationAngle = rotationAngle + 180
        If (rotationAngle <= -360) Or (rotationAngle >= 360) Then rotationAngle = (Int(rotationAngle) Mod 360) + (rotationAngle - Int(rotationAngle))
        
        'Perform the rotation
        
        'Transform the destination world matrix twice: once for the rotation angle, and once again to offset all coordinates.
        ' This allows us to rotate the image around its *center* rather than around its top-left corner.
        If GdipRotateWorldTransform(iGraphics, rotationAngle, 0&) = GP_OK Then
            If GdipTranslateWorldTransform(iGraphics, dstX + dstWidth / 2, dstY + dstHeight / 2, 1&) = GP_OK Then
        
                'Render the image onto the destination
                If GdipDrawImageRectRectI(iGraphics, tBitmap, -dstWidth / 2, -dstHeight / 2, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle) <> 0 Then
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
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    'Next, we need a temporary copy of the image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, dstDIB
        
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
            GDIPlusDebug = GdipDrawImageFX(iGraphics, tBitmap, tmpRect, tmpMatrix, hEffect, 0&, GP_U_Pixel)
            
            If GDIPlusDebug = 0 Then
                GDIPlusBlurDIB = True
            Else
                GDIPlusBlurDIB = False
                Message "GDI+ failed to render blur effect (Error Code %1).", GDIPlusDebug
            End If
            
            'Delete our temporary transformation matrix
            GdipDeleteMatrix tmpMatrix
            
        Else
            GDIPlusBlurDIB = False
            Message "GDI+ failed to set effect parameters."
        End If
    
        'Delete our GDI+ blur object
        GdipDeleteEffect hEffect
    
    Else
        GDIPlusBlurDIB = False
        Message "GDI+ failed to create blur effect object"
    End If
        
    'Release both the destination graphics object and the source bitmap object
    GdipDeleteGraphics iGraphics
    GdipDisposeImage tBitmap
    
End Function

'Use GDI+ to render a series of white-black-white circles, which are preferable for on-canvas controls with good readability
Public Function GDIPlusDrawCanvasCircle(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, Optional ByVal cTransparency As Long = 190, Optional ByVal useHighlightColor As Boolean = False) As Boolean

    GDIPlusDrawCircleToDC dstDC, cx, cy, cRadius, RGB(0, 0, 0), cTransparency, 3, True
    
    Dim topColor As Long
    If useHighlightColor Then topColor = g_Themer.GetGenericUIColor(UI_AccentLight) Else topColor = RGB(255, 255, 255)
    GDIPlusDrawCircleToDC dstDC, cx, cy, cRadius, topColor, 220, 1, True
    
End Function

'Identical function to GdiPlusDrawCanvasCircle, above, but a rect is used instead.  Note that it's inconvenient to the user to display
' a square but use circles for hit-detection, so plan accordingly!
Public Function GDIPlusDrawCanvasSquare(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, Optional ByVal cTransparency As Long = 190, Optional ByVal useHighlightColor As Boolean = False) As Boolean

    GDI_Plus.GDIPlusDrawRectOutlineToDC dstDC, cx - cRadius, cy - cRadius, cx + cRadius, cy + cRadius, RGB(0, 0, 0), cTransparency, 3, True, GP_LC_Round, True
    
    Dim topColor As Long
    If useHighlightColor Then topColor = g_Themer.GetGenericUIColor(UI_AccentLight) Else topColor = RGB(255, 255, 255)
    GDI_Plus.GDIPlusDrawRectOutlineToDC dstDC, cx - cRadius, cy - cRadius, cx + cRadius, cy + cRadius, topColor, 220, 1.6, True, GP_LC_Round, True
    
End Function

'Similar function to GdiPlusDrawCanvasCircle, above, but only draws a single line
Public Function GDIPlusDrawCanvasLine(ByVal dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal cTransparency As Long = 190, Optional ByVal useHighlightColor As Boolean = False) As Boolean

    GDI_Plus.GDIPlusDrawLineToDC dstDC, x1, y1, x2, y2, RGB(0, 0, 0), cTransparency, 3, True, GP_LC_Square, True
    
    Dim topColor As Long
    If useHighlightColor Then topColor = g_Themer.GetGenericUIColor(UI_AccentLight) Else topColor = RGB(255, 255, 255)
    GDI_Plus.GDIPlusDrawLineToDC dstDC, x1, y1, x2, y2, topColor, 220, 1.6, True, GP_LC_Round, True
    
End Function

'Similar function to GdiPlusDrawCanvasCircle, above, but draws a RectF outline, specifically
Public Function GDIPlusDrawCanvasRectF(ByVal dstDC As Long, ByRef srcRect As RECTF, Optional ByVal cTransparency As Long = 190, Optional ByVal useHighlightColor As Boolean = False) As Boolean
    GDI_Plus.GDIPlusDrawRectFOutlineToDC dstDC, srcRect, g_Themer.GetGenericUIColor(UI_LineEdge, , , useHighlightColor), cTransparency, 3, True, GP_LJ_Miter
    GDI_Plus.GDIPlusDrawRectFOutlineToDC dstDC, srcRect, g_Themer.GetGenericUIColor(UI_LineCenter, , , useHighlightColor), 220, 1.6, True, GP_LJ_Miter
End Function

'Assuming the client has already obtained a GDI+ graphics handle and a GDI+ pen handle, they can use this function to quickly draw a line using
' the associated objects.
Public Sub GDIPlusDrawLine_Fast(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)

    'This function is just a thin wrapper to the GdipDrawLine function!
    GdipDrawLine dstGraphics, srcPen, x1, y1, x2, y2

End Sub

'Use GDI+ to render a line, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawLineToDC(ByVal dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLineCap As GP_LineCap = GP_LC_Flat, Optional ByVal hqOffsets As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    If hqOffsets Then GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
    
    'Create a pen, which will be used to stroke the line
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(eColor, cTransparency), lineWidth, GP_U_Pixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLineCap > 0 Then GdipSetPenLineCap iPen, customLineCap, customLineCap, 0&
    
    'Render the line
    GdipDrawLine iGraphics, iPen, x1, y1, x2, y2
        
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a filled, closed shape, with optional color, opacity, antialiasing, curvature, and more
Public Function GDIPlusDrawFilledShapeToDC(ByVal dstDC As Long, ByVal numOfPoints As Long, ByVal ptrToFloatArray As Long, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal useAA As Boolean = True, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5, Optional ByVal useFillMode As GP_FillMode = GP_FM_Alternate) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    
    'Create a solid fill brush
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, cTransparency)
    
    If hBrush <> 0 Then
    
        'We have a few different options for drawing the shape, based on the passed parameters.
        If useCurveAlgorithm Then
            GdipFillClosedCurve2 iGraphics, hBrush, ptrToFloatArray, numOfPoints, curvatureTension, useFillMode
        Else
            GdipFillPolygon iGraphics, hBrush, ptrToFloatArray, numOfPoints, useFillMode
        End If
        
        ReleaseGDIPlusBrush hBrush
        
    End If
    
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render the outline of a closed shape, with optional color, opacity, antialiasing, curvature, and more
Public Function GDIPlusStrokePathToDC(ByVal dstDC As Long, ByVal numOfPoints As Long, ByVal ptrToFloatArray As Long, ByVal autoCloseShape As Boolean, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal useAA As Boolean = True, Optional ByVal strokeWidth As Single = 1, Optional ByVal customLineCap As GP_LineCap = GP_LC_Flat, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    
    'Create a pen, which will be used to stroke the line
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(eColor, cTransparency), strokeWidth, GP_U_Pixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLineCap > 0 Then GdipSetPenLineCap iPen, customLineCap, customLineCap, 0&
        
    'We have a few different options for drawing the shape, based on the passed parameters.
    If autoCloseShape Then
    
        If useCurveAlgorithm Then
            GdipDrawClosedCurve2 iGraphics, iPen, ptrToFloatArray, numOfPoints, curvatureTension
        Else
            GdipDrawPolygon iGraphics, iPen, ptrToFloatArray, numOfPoints
        End If
        
    Else
    
        If useCurveAlgorithm Then
            GdipDrawCurve2 iGraphics, iPen, ptrToFloatArray, numOfPoints, curvatureTension
        Else
            GdipDrawLines iGraphics, iPen, ptrToFloatArray, numOfPoints
        End If
    
    End If
    
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a hollow rectangle, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawRectOutlineToDC(ByVal dstDC As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLinejoin As GP_LineJoin = GP_LJ_Bevel, Optional ByVal hqOffsets As Boolean = False, Optional ByVal useInsetMode As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    If hqOffsets Then GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
    
    'Create a pen, which will be used to stroke the line
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(eColor, cTransparency), lineWidth, GP_U_Pixel, iPen
    
    'Apply any other custom settings now
    If customLinejoin > 0 Then GdipSetPenLineJoin iPen, customLinejoin
    If useInsetMode Then GdipSetPenMode iPen, GP_PA_Inset Else GdipSetPenMode iPen, GP_PA_Center
    
    'Render the rectangle
    GdipDrawRectangle iGraphics, iPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop
            
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

Public Function GDIPlusDrawRectLOutlineToDC(ByVal dstDC As Long, ByRef srcRectL As RECTL, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLinejoin As GP_LineJoin = GP_LJ_Bevel, Optional ByVal hqOffsets As Boolean = False, Optional ByVal useInsetMode As Boolean = False) As Boolean
    GDIPlusDrawRectLOutlineToDC = GDIPlusDrawRectOutlineToDC(dstDC, srcRectL.Left, srcRectL.Top, srcRectL.Right, srcRectL.Bottom, eColor, cTransparency, lineWidth, useAA, customLinejoin, hqOffsets, useInsetMode)
End Function

Public Function GDIPlusDrawRectFOutlineToDC(ByVal dstDC As Long, ByRef srcRectF As RECTF, ByVal eColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal lineWidth As Single = 1, Optional ByVal useAA As Boolean = True, Optional ByVal customLinejoin As GP_LineJoin = GP_LJ_Bevel, Optional ByVal hqOffsets As Boolean = False, Optional ByVal useInsetMode As Boolean = False) As Boolean
    GDIPlusDrawRectFOutlineToDC = GDIPlusDrawRectOutlineToDC(dstDC, srcRectF.Left, srcRectF.Top, srcRectF.Left + srcRectF.Width, srcRectF.Top + srcRectF.Height, eColor, cTransparency, lineWidth, useAA, customLinejoin, hqOffsets, useInsetMode)
End Function

'Use GDI+ to render a hollow circle, with optional color, opacity, and antialiasing
Public Function GDIPlusDrawCircleToDC(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, ByVal edgeColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal drawRadius As Single = 1, Optional ByVal useAA As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    
    'Create a pen, which will be used to stroke the circle
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(edgeColor, cTransparency), drawRadius, GP_U_Pixel, iPen
    
    'Render the circle
    GdipDrawEllipse iGraphics, iPen, cx - cRadius, cy - cRadius, cRadius * 2, cRadius * 2
        
    'Release all created objects
    GdipDeletePen iPen
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to render a filled circle, with optional color, opacity, and antialiasing
Public Function GDIPlusFillCircleToDC(ByVal dstDC As Long, ByVal cx As Single, ByVal cy As Single, ByVal cRadius As Single, ByVal fillColor As Long, Optional ByVal cTransparency As Long = 255, Optional ByVal useAA As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    
    'Create a brush, which will be used to stroke the circle
    Dim hBrush As Long
    hBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(fillColor, cTransparency)
    
    If hBrush <> 0 Then
        GDIPlusFillCircleToDC = CBool(GdipFillEllipse(iGraphics, hBrush, cx - cRadius, cy - cRadius, cRadius * 2, cRadius * 2) = 0)
        GDI_Plus.ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to fill a DC with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full
' DIB alpha in a single pass, which is important for 32-bpp DIBs.
Public Function GDIPlusFillRectToDC(ByVal dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    If useAA Then GdipSetSmoothingMode hGraphics, GP_SM_Antialias Else GdipSetSmoothingMode hGraphics, GP_SM_None
    GdipSetCompositingMode hGraphics, dstFillMode
    
    'Create a solid fill brush using the specified color
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, eTransparency)
    
    If hBrush <> 0 Then
        GdipFillRectangle hGraphics, hBrush, x1, y1, xWidth, yHeight
        ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillRectToDC = True

End Function


'Use GDI+ to fill a DC with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full
' DIB alpha in a single pass, which is important for 32-bpp DIBs.
Public Function GDIPlusFillRectLToDC(ByVal dstDC As Long, ByRef srcRect As RECTL, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    If useAA Then GdipSetSmoothingMode hGraphics, GP_SM_Antialias Else GdipSetSmoothingMode hGraphics, GP_SM_None
    GdipSetCompositingMode hGraphics, dstFillMode
    
    'Create a solid fill brush using the specified color
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, eTransparency)
    
    If hBrush <> 0 Then
        GdipFillRectangle hGraphics, hBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top
        ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillRectLToDC = True

End Function

'Use GDI+ to fill a DC with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full
' DIB alpha in a single pass, which is important for 32-bpp DIBs.
Public Function GDIPlusFillRectFToDC(ByVal dstDC As Long, ByRef srcRect As RECTF, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    
    If useAA Then GdipSetSmoothingMode hGraphics, GP_SM_Antialias Else GdipSetSmoothingMode hGraphics, GP_SM_None
    GdipSetCompositingMode hGraphics, dstFillMode
    
    'Create a solid fill brush using the specified color
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, eTransparency)
    
    If hBrush <> 0 Then
        GdipFillRectangle hGraphics, hBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height
        ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillRectFToDC = True

End Function

'Given a source DIB, fill it with the alpha checkerboard pattern.  32bpp images can then be alpha blended onto it.
Public Function GDIPlusFillPatternToDC(ByVal dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByRef srcDIB As pdDIB, Optional ByVal fixBoundaryPainting As Boolean = False) As Boolean
    
    'Create a GDI+ copy of the image and request AA
    Dim hGraphics As Long
    GdipCreateFromHDC dstDC, hGraphics
    GdipSetSmoothingMode hGraphics, GP_SM_Antialias
    GdipSetCompositingQuality hGraphics, GP_CQ_AssumeLinear
    GdipSetPixelOffsetMode hGraphics, GP_POM_HighSpeed
        
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
        xWidth = Int(xWidth - xDif - 0.5)
        yHeight = Int(yHeight - yDif - 0.5)
        
    End If
    
    'Apply the brush
    GdipFillRectangle hGraphics, hBrush, x1, y1, xWidth, yHeight
    
    'Release all created objects
    ReleaseGDIPlusBrush hBrush
    GdipDisposeImage srcBitmap
    GdipDeleteGraphics hGraphics
    
    GDIPlusFillPatternToDC = True
    
End Function

'Use GDI+ to render a filled ellipse, with optional antialiasing
Public Function GDIPlusFillEllipseToDC(ByRef dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal useAA As Boolean = True, Optional ByVal eTransparency As Byte = 255, Optional ByVal hqOffsets As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request matching AA and offset behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    If hqOffsets Then GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
    
    'Create a solid fill brush
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, eTransparency)
    
    If hBrush <> 0 Then
        GdipFillEllipseI iGraphics, hBrush, x1, y1, xWidth, yHeight
        ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics iGraphics
    
    GDIPlusFillEllipseToDC = True

End Function

'Use GDI+ to render an ellipse outline, with optional antialiasing
Public Function GDIPlusStrokeEllipseToDC(ByRef dstDC As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal useAA As Boolean = True, Optional ByVal eTransparency As Byte = 255, Optional ByVal strokeWidth As Single = 1#) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
        
    'Create a pen with matching attributes
    Dim hPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(eColor, eTransparency), strokeWidth, GP_U_Pixel, hPen
    
    'Render the ellipse
    GdipDrawEllipse iGraphics, hPen, x1, y1, xWidth, yHeight
    
    'Release all created objects
    GdipDeletePen hPen
    GdipDeleteGraphics iGraphics

    GDIPlusStrokeEllipseToDC = True

End Function

'Use GDI+ to render a rectangle with rounded corners, with optional antialiasing
Public Function GDIPlusDrawRoundRect(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal rRadius As Single, ByVal eColor As Long, Optional ByVal useAA As Boolean = True, Optional ByVal FillRect As Boolean = True) As Boolean

    'Create a GDI+ copy of the image and request matching AA behavior
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    If useAA Then GdipSetSmoothingMode iGraphics, GP_SM_Antialias Else GdipSetSmoothingMode iGraphics, GP_SM_None
    
    'GDI+ doesn't have a direct rounded rectangles call, so we have to do it ourselves with a custom path
    Dim rrPath As Long
    GdipCreatePath GP_FM_Winding, rrPath
        
    'The path will be rendered in two sections: first, filling it.  Second, stroking the path itself to complete the
    ' 1px outside border.
    xWidth = xWidth - 1
    yHeight = yHeight - 1
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = rRadius
    yCurvature = rRadius
    
    If xCurvature > xWidth Then xCurvature = xWidth
    If yCurvature > yHeight Then yCurvature = yHeight
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    GdipAddPathArc rrPath, x1 + xWidth - xCurvature, y1, xCurvature, yCurvature, 270, 90
    GdipAddPathArc rrPath, x1 + xWidth - xCurvature, y1 + yHeight - yCurvature, xCurvature, yCurvature, 0, 90
    GdipAddPathArc rrPath, x1, y1 + yHeight - yCurvature, xCurvature, yCurvature, 90, 90
    GdipAddPathArc rrPath, x1, y1, xCurvature, yCurvature, 180, 90
    GdipClosePathFigure rrPath
    
    'Create a solid fill brush
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, 255)
    
    If hBrush <> 0 Then
        If FillRect Then GdipFillPath iGraphics, hBrush, rrPath
        ReleaseGDIPlusBrush hBrush
    End If
    
    'Stroke the path as well (to fill the 1px exterior border)
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(eColor, 255), 1, GP_U_Pixel, iPen
    GdipDrawPath iGraphics, iPen, rrPath
    
    'Release all created objects
    GdipDeletePen iPen
    GdipDeletePath rrPath
    GdipDeleteGraphics iGraphics

End Function

'Use GDI+ to fill a DIB with a color and optional alpha value; while not as efficient as using GDI, this allows us to set the full DIB alpha
' in a single pass.
Public Function GDIPlusFillDIBRect(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, ByVal eColor As Long, Optional ByVal cOpacity As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    If useAA Then
        GdipSetSmoothingMode iGraphics, GP_SM_Antialias
    Else
        GdipSetSmoothingMode iGraphics, GP_SM_None
    End If
    
    GdipSetCompositingMode iGraphics, dstFillMode
    
    'Create a solid fill brush from the source image
    Dim hBrush As Long
    hBrush = GetGDIPlusSolidBrushHandle(eColor, cOpacity)
    
    If hBrush <> 0 Then
        GdipFillRectangle iGraphics, hBrush, x1, y1, xWidth, yHeight
        ReleaseGDIPlusBrush hBrush
    End If
    
    GdipDeleteGraphics iGraphics
    
    GDIPlusFillDIBRect = True

End Function

Public Function GDIPlusFillDIBRectL(ByRef dstDIB As pdDIB, ByRef srcRectL As RECTL, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean
    GDIPlusFillDIBRectL = GDIPlusFillDIBRect(dstDIB, srcRectL.Left, srcRectL.Top, srcRectL.Right - srcRectL.Left, srcRectL.Bottom - srcRectL.Top, eColor, eTransparency, dstFillMode, useAA)
End Function

Public Function GDIPlusFillDIBRectF(ByRef dstDIB As pdDIB, ByRef srcRectF As RECTF, ByVal eColor As Long, Optional ByVal eTransparency As Long = 255, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean
    GDIPlusFillDIBRectF = GDIPlusFillDIBRect(dstDIB, srcRectF.Left, srcRectF.Top, srcRectF.Width, srcRectF.Height, eColor, eTransparency, dstFillMode, useAA)
End Function

'Given a source DIB, fill it with the alpha checkerboard pattern.  32bpp images can then be alpha blended onto it.
Public Function GDIPlusFillDIBRect_Pattern(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal bltWidth As Single, ByVal bltHeight As Single, ByRef srcDIB As pdDIB, Optional ByVal useThisDCInstead As Long = 0, Optional ByVal fixBoundaryPainting As Boolean = False) As Boolean
    
    'Create a GDI+ copy of the image and request AA
    Dim iGraphics As Long
    
    If useThisDCInstead <> 0 Then
        GdipCreateFromHDC useThisDCInstead, iGraphics
    Else
        GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    End If
    
    GdipSetSmoothingMode iGraphics, GP_SM_Antialias
    GdipSetCompositingQuality iGraphics, GP_CQ_AssumeLinear
    GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
        
    'Create a texture fill brush from the source image
    Dim srcBitmap As Long, iBrush As Long
    GetGdipBitmapHandleFromDIB srcBitmap, srcDIB
    GdipCreateTexture srcBitmap, GP_WM_Tile, iBrush
    
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
    GdipFillRectangle iGraphics, iBrush, x1, y1, bltWidth, bltHeight
    
    'Release all created objects
    ReleaseGDIPlusBrush iBrush
    GdipDisposeImage srcBitmap
    GdipDeleteGraphics iGraphics
    
    GDIPlusFillDIBRect_Pattern = True
    
End Function

'Use GDI+ to fill an arbitrary DC with an arbitrary GDI+ brush
Public Function GDIPlusFillDC_Brush(ByRef dstDC As Long, ByVal srcBrushHandle As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal xWidth As Single, ByVal yHeight As Single, Optional ByVal dstFillMode As GP_CompositingMode = GP_CM_SourceOver, Optional ByVal useAA As Boolean = False) As Boolean

    'Create a GDI+ copy of the image and request AA
    Dim iGraphics As Long
    GdipCreateFromHDC dstDC, iGraphics
    
    If useAA Then
        GdipSetSmoothingMode iGraphics, GP_SM_Antialias
    Else
        GdipSetSmoothingMode iGraphics, GP_SM_None
    End If
    
    GdipSetCompositingMode iGraphics, dstFillMode
    
    'Apply the brush
    GdipFillRectangle iGraphics, srcBrushHandle, x1, y1, xWidth, yHeight
    
    'Release all created objects
    GdipDeleteGraphics iGraphics
    
    GDIPlusFillDC_Brush = True

End Function

'Use GDI+ to quickly convert a 24bpp DIB to 32bpp with solid alpha channel
Public Sub GDIPlusConvertDIB24to32(ByRef dstDIB As pdDIB)
    
    If dstDIB.GetDIBColorDepth = 32 Then Exit Sub
    
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
    
    GdipCreateBitmapFromGdiDib imgHeader, ByVal srcDIB.GetDIBPointer, srcBitmap
    
    'Next, recreate the destination DIB as 32bpp
    dstDIB.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, , 255
    
    'Clone the bitmap area from source to destination, while converting format as necessary
    Dim gdipReturn As Long
    gdipReturn = GdipCloneBitmapAreaI(0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, PixelFormat32bppARGB, srcBitmap, dstBitmap)
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    'Paint the converted image to the destination
    GdipDrawImage iGraphics, dstBitmap, 0, 0
    
    'The target image will always have premultiplied alpha (not really relevant, as the source is 24-bpp, but this
    ' lets us use various accelerated codepaths throughout the project).
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'Release our bitmap copies and GDI+ instances
    GdipDisposeImage srcBitmap
    GdipDisposeImage dstBitmap
    GdipDeleteGraphics iGraphics
 
End Sub

'Use GDI+ to load an image file.  Pretty bare-bones, but should be sufficient for any supported image type.
Public Function GDIPlusLoadPicture(ByVal srcFilename As String, ByRef dstDIB As pdDIB) As Boolean

    'Used to hold the return values of various GDI+ calls
    Dim GDIPlusReturn As GP_Result
      
    'Use GDI+ to load the image
    Dim hImage As Long
    GDIPlusReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    If (GDIPlusReturn <> GP_OK) Then
        If (hImage <> 0) Then GdipDisposeImage hImage
        GDIPlusLoadPicture = False
        Exit Function
    End If
    
    'If we're still here, the image (probably) loaded successfully.  Create a destination DIB as necessary.
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Retrieve the image's format as a GUID
    Dim imgCLSID As clsid
    GdipGetImageRawFormat hImage, VarPtr(imgCLSID)
    
    'Convert the GUID into a string
    Dim imgStringPointer As Long, imgFormatGuidString As String
    StringFromCLSID VarPtr(imgCLSID), imgStringPointer
    imgFormatGuidString = pvPtrToStrW(imgStringPointer)
    
    'And finally, convert the string into an FIF long
    Dim imgFormatFIF As Long
    imgFormatFIF = GetFIFFromGUID(imgFormatGuidString)
    
    'Metafiles require special consideration; set that flag in advance
    Dim isMetafile As Boolean
    If (imgFormatFIF = PDIF_EMF) Or (imgFormatFIF = PDIF_WMF) Then
        isMetafile = True
    Else
        isMetafile = False
    End If
    
    'Look for an ICC profile by asking GDI+ to return the ICC profile property's size
    Dim profileSize As Long, HasProfile As Boolean
    
    'NOTE! the passed profileSize value must always be zeroed before using GdipGetPropertyItemSize, because the function will not update
    ' the variable's value if no tag is found.  Seems like an asinine oversight, but oh well.
    profileSize = 0
    GdipGetPropertyItemSize hImage, PropertyTagICCProfile, profileSize
    
    'If the returned size is > 0, this image contains an ICC profile!  Retrieve it now.
    If (profileSize > 0) Then
    
        HasProfile = True
    
        Dim iccProfileBuffer() As Byte
        ReDim iccProfileBuffer(0 To profileSize - 1) As Byte
        GdipGetPropertyItem hImage, PropertyTagICCProfile, profileSize, ByVal VarPtr(iccProfileBuffer(0))
        
        dstDIB.ICCProfile.LoadICCFromPtr profileSize - 16, VarPtr(iccProfileBuffer(0)) + 16
        
        Erase iccProfileBuffer
        
    End If
    
    'Look for orientation flags.  This is most relevant for JPEGs coming from a digital camera.
    profileSize = 0
    GdipGetPropertyItemSize hImage, PropertyTagOrientation, profileSize
    
    If (profileSize > 0) And g_UserPreferences.GetPref_Boolean("Loading", "ExifAutoRotate", True) Then
        
        'Orientation tag will only ever be 2 bytes
        Dim tmpPropertyBuffer() As Byte
        ReDim tmpPropertyBuffer(0 To profileSize - 1) As Byte
        GdipGetPropertyItem hImage, PropertyTagOrientation, profileSize, ByVal VarPtr(tmpPropertyBuffer(0))
        
        'The first 16 bytes of a GDI+ property are a standard header.  We need the MSB of the 2-byte trailer of the returned array.
        Select Case tmpPropertyBuffer(profileSize - 2)
        
            'Standard orientation - ignore!
            Case 1
        
            'The 0th row is at the visual top of the image, and the 0th column is the visual right-hand side
            Case 2
                GdipImageRotateFlip hImage, RotateNoneFlipX
            
            'The 0th row is at the visual bottom of the image, and the 0th column is the visual right-hand side
            Case 3
                GdipImageRotateFlip hImage, Rotate180FlipNone
            
            'The 0th row is at the visual bottom of the image, and the 0th column is the visual left-hand side
            Case 4
                GdipImageRotateFlip hImage, RotateNoneFlipY
            
            'The 0th row is the visual left-hand side of of the image, and the 0th column is the visual top
            Case 5
                GdipImageRotateFlip hImage, Rotate270FlipY
            
            'The 0th row is the visual right -hand side of of the image, and the 0th column is the visual top
            Case 6
                GdipImageRotateFlip hImage, Rotate90FlipNone
                
            'The 0th row is the visual right -hand side of of the image, and the 0th column is the visual bottom
            Case 7
                GdipImageRotateFlip hImage, Rotate90FlipY
                
            'The 0th row is the visual left-hand side of of the image, and the 0th column is the visual bottom
            Case 8
                GdipImageRotateFlip hImage, Rotate270FlipNone
        
        End Select
        
    End If
    
    'Metafiles can contain brushes and other objects stored at extremely high DPIs.  Limit these to 300 dpi to prevent OOM errors later on.
    If isMetafile Then GdipSetMetafileDownLevelRasterizationLimit hImage, 300
    
    'Retrieve the image's size
    ' RANDOM FACT! GdipGetImageDimension works fine on bitmaps.  On metafiles, it returns bizarre values that may be astronomically large.
    '  My assumption is that image dimensions are not necessarily returned in pixels (though pixels are the default for bitmaps).  Anyway,
    '  it's trivial to switch to GdipGetImageWidth/Height.  Original code was: GdipGetImageDimension hImage, imgWidth, imgHeight -- and
    '  note that the original code required Single-type values instead of Longs.
    Dim imgWidth As Long, imgHeight As Long
    GdipGetImageWidth hImage, imgWidth
    GdipGetImageHeight hImage, imgHeight
    
    'Retrieve the image's horizontal and vertical resolution (if any)
    Dim imgHResolution As Single, imgVResolution As Single
    GdipGetImageHorizontalResolution hImage, imgHResolution
    GdipGetImageVerticalResolution hImage, imgVResolution
    dstDIB.SetDPI imgHResolution, imgVResolution
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "GDI+ image resolution reported as: " & imgHResolution & "x" & imgVResolution
    #End If
    
    'Metafile containers (EMF, WMF) require special handling.
    Dim emfPlusConversionSuccessful As Boolean
    emfPlusConversionSuccessful = False
    
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
            If (imgHResolution <> 0) Then imgWidth = imgWidth * CDbl(96 / imgHResolution) Else imgHResolution = 96
            If (imgVResolution <> 0) Then imgHeight = imgHeight * CDbl(96 / imgVResolution) Else imgVResolution = 96
            
        End If
        
        'If GDI+ v1.1 is available, we can translate EMFs and WMFs into the new GDI+ EMF+ format, which supports antialiasing
        ' and alpha channels (among other things).
        If GDI_Plus.IsGDIPlusV11Available Then
            
            'Create a temporary GDI+ graphics object, whose properties will be used to control the render state of the EMF
            Dim tmpSettingsDIB As pdDIB
            Set tmpSettingsDIB = New pdDIB
            tmpSettingsDIB.CreateBlank 8, 8, 32, 0, 0
            
            Dim tmpGraphics As Long
            If GdipCreateFromHDC(tmpSettingsDIB.GetDIBDC, tmpGraphics) = GP_OK Then
                
                'Set high-quality antialiasing and interpolation
                GdipSetSmoothingMode tmpGraphics, GP_SM_Antialias
                GdipSetInterpolationMode tmpGraphics, GP_IM_HighQualityBicubic
                
                'Attempt to convert the EMF to EMF+ format
                Dim mfHandleDst As Long, convSuccess As Long
                
                'For reference: if we ever want to write our improved EMF+ data to file, we can use code like the following:
                'Dim newEmfPlusFilename As String
                'newEmfPlusFilename = srcFilename
                'StripOffExtension newEmfPlusFilename
                'newEmfPlusFilename = newEmfPlusFilename & " (EMFPlus).emf"
                'If GdipConvertToEmfPlusToFile(tmpGraphics, hImage, convSuccess, StrPtr(newEmfPlusFilename), EmfTypeEmfPlusOnly, 0, mfHandleDst) = 0 Then
                
                If GdipConvertToEmfPlus(tmpGraphics, hImage, convSuccess, EmfTypeEmfPlusOnly, 0, mfHandleDst) = 0 Then

                    'Conversion successful!  Replace our current image handle with the EMF+ copy
                    emfPlusConversionSuccessful = True
                    GdipDisposeImage hImage
                    hImage = mfHandleDst

                End If
                
                'Release our temporary graphics container
                GdipDeleteGraphics tmpGraphics
                
            End If
            
            'Release our temporary settings DIB
            Set tmpSettingsDIB = Nothing
            
        End If
        
    End If
        
    'Retrieve the image's alpha channel data (if any)
    Dim hasAlpha As Boolean
    hasAlpha = False
    
    Dim iPixelFormat As Long
    GdipGetImagePixelFormat hImage, iPixelFormat
    If (iPixelFormat And PixelFormatAlpha) <> 0 Then hasAlpha = True
    If (iPixelFormat And PixelFormatPremultipliedAlpha) <> 0 Then hasAlpha = True
    
    'Make a note of the image's specific color depth, as relevant to PD
    Dim imgColorDepth As Long
    imgColorDepth = GetColorDepthFromPixelFormat(iPixelFormat)
    
    'Check for CMYK images
    Dim isCMYK As Boolean
    If (iPixelFormat = PixelFormat32bppCMYK) Then isCMYK = True
    
    'Create a blank PD-compatible DIB
    If isCMYK Then
        dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 24
    Else
        
        'Metafiles require special handling on Vista and earlier
        If isMetafile Then
            
            If emfPlusConversionSuccessful Or hasAlpha Then
                dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 32
            Else
                dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 24
            End If
        
        'Non-metafiles can always be placed into a 32bpp container.
        Else
            dstDIB.CreateBlank CLng(imgWidth), CLng(imgHeight), 32
        End If
        
    End If
    
    Dim copyBitmapData As GP_BitmapData
    Dim tmpRect As RECTL
    Dim iGraphics As Long
    
    'We now copy over image data in one of two ways.  If the image is 24bpp, our job is simple - use BitBlt and an hBitmap.
    ' 32bpp (including CMYK) images require a bit of extra work.
    If hasAlpha Then
        
        'Make sure the image is in 32bpp premultiplied ARGB format
        If (iPixelFormat <> PixelFormat32bppPARGB) Then GdipCloneBitmapAreaI 0, 0, imgWidth, imgHeight, PixelFormat32bppPARGB, hImage, hImage
        
        'We are now going to copy the image's data directly into our destination DIB by using LockBits.  Very fast, and not much code!
        
        'Start by preparing a BitmapData variable with instructions on where GDI+ should paste the bitmap data
        With copyBitmapData
            .BD_Width = imgWidth
            .BD_Height = imgHeight
            .BD_PixelFormat = PixelFormat32bppPARGB
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
        GdipBitmapLockBits hImage, tmpRect, ImageLockModeUserInputBuf Or ImageLockModeWrite Or ImageLockModeRead, PixelFormat32bppPARGB, copyBitmapData
        GdipBitmapUnlockBits hImage, copyBitmapData
    
    Else
    
        'CMYK is handled separately from regular RGB data, as we want to perform an ICC profile conversion as well.
        ' Note that if a CMYK profile is not present, we allow GDI+ to convert the image to RGB for us.
        If (isCMYK And HasProfile) Then
        
            'Create a blank 32bpp DIB, which will hold the CMYK data
            Dim tmpCMYKDIB As pdDIB
            Set tmpCMYKDIB = New pdDIB
            tmpCMYKDIB.CreateBlank imgWidth, imgHeight, 32
        
            'Next, prepare a BitmapData variable with instructions on where GDI+ should paste the bitmap data
            With copyBitmapData
                .BD_Width = imgWidth
                .BD_Height = imgHeight
                .BD_PixelFormat = PixelFormat32bppCMYK
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
            GdipBitmapLockBits hImage, tmpRect, ImageLockModeUserInputBuf Or ImageLockModeWrite Or ImageLockModeRead, PixelFormat32bppCMYK, copyBitmapData
            GdipBitmapUnlockBits hImage, copyBitmapData
            
            'Apply the transformation using the dedicated CMYK transform handler
            If ColorManagement.ApplyCMYKTransform_WindowsCMS(dstDIB.ICCProfile.GetICCDataPointer, dstDIB.ICCProfile.GetICCDataSize, tmpCMYKDIB, dstDIB, dstDIB.ICCProfile.GetSourceRenderIntent) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Copying newly transformed sRGB data..."
                #End If
            
                'The transform was successful, and the destination DIB is ready to go!
                dstDIB.ICCProfile.MarkSuccessfulProfileApplication
                                
            'Something went horribly wrong.  Use GDI+ to apply a generic CMYK -> RGB transform.
            Else
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "ICC-based CMYK transformation failed.  Falling back to default CMYK conversion..."
                #End If
            
                GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
                GdipDrawImageRect iGraphics, hImage, 0, 0, imgWidth, imgHeight
                GdipDeleteGraphics iGraphics
            
            End If
            
            Set tmpCMYKDIB = Nothing
        
        Else
            
            'Render the GDI+ image directly onto the newly created DIB
            GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
            GdipDrawImageRect iGraphics, hImage, 0, 0, imgWidth, imgHeight
            GdipDeleteGraphics iGraphics
            
        End If
    
    End If
    
    'Note some original file settings inside the DIB
    dstDIB.SetOriginalFormat imgFormatFIF
    dstDIB.SetOriginalColorDepth imgColorDepth
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'Release any remaining GDI+ handles and exit
    GdipDisposeImage hImage
    GDIPlusLoadPicture = True
    
End Function

'Given a GDI+ pixel format value, return a numeric color depth (e.g. 24, 32, etc)
Private Function GetColorDepthFromPixelFormat(ByVal gdipPixelFormat As Long) As Long

    If (gdipPixelFormat = PixelFormat1bppIndexed) Then
        GetColorDepthFromPixelFormat = 1
    ElseIf (gdipPixelFormat = PixelFormat4bppIndexed) Then
        GetColorDepthFromPixelFormat = 4
    ElseIf (gdipPixelFormat = PixelFormat8bppIndexed) Then
        GetColorDepthFromPixelFormat = 8
    ElseIf (gdipPixelFormat = PixelFormat16bppGreyscale) Or (gdipPixelFormat = PixelFormat16bppRGB555) Or (gdipPixelFormat = PixelFormat16bppRGB565) Or (gdipPixelFormat = PixelFormat16bppARGB1555) Then
        GetColorDepthFromPixelFormat = 16
    ElseIf (gdipPixelFormat = PixelFormat24bppRGB) Or (gdipPixelFormat = PixelFormat32bppRGB) Then
        GetColorDepthFromPixelFormat = 24
    ElseIf (gdipPixelFormat = PixelFormat32bppARGB) Or (gdipPixelFormat = PixelFormat32bppPARGB) Then
        GetColorDepthFromPixelFormat = 32
    ElseIf (gdipPixelFormat = PixelFormat48bppRGB) Then
        GetColorDepthFromPixelFormat = 48
    ElseIf (gdipPixelFormat = PixelFormat64bppARGB) Or (gdipPixelFormat = PixelFormat64bppPARGB) Then
        GetColorDepthFromPixelFormat = 64
    Else
        GetColorDepthFromPixelFormat = 24
    End If

End Function

'Save an image using GDI+.  Per the current save spec, ImageID must be specified.
' Additional save options are currently available for JPEGs (save quality, range [1,100]) and TIFFs (compression type).
Public Function GDIPlusSavePicture(ByRef srcPDImage As pdImage, ByVal dstFilename As String, ByVal imgFormat As GDIPlusImageFormat, ByVal outputColorDepth As Long, Optional ByVal jpegQuality As Long = 92) As Boolean

    On Error GoTo GDIPlusSaveError

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Prepping image for GDI+ export..."
    #End If
    
    'If the output format is 24bpp (e.g. JPEG) but the input image is 32bpp, composite it against white
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcPDImage.GetCompositedImage tmpDIB, False
    If (tmpDIB.GetDIBColorDepth <> 24) And imgFormat = [ImageJPEG] Then tmpDIB.CompositeBackgroundColor 255, 255, 255

    'Begin by creating a generic bitmap header for the current DIB
    Dim imgHeader As BITMAPINFO
    
    With imgHeader.Header
        .Size = Len(imgHeader.Header)
        .Planes = 1
        .BitCount = tmpDIB.GetDIBColorDepth
        .Width = tmpDIB.GetDIBWidth
        .Height = -tmpDIB.GetDIBHeight
    End With

    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hImage As Long
    
    Message "Creating GDI+ compatible image copy..."
        
    'Different GDI+ calls are required for different color depths. GdipCreateBitmapFromGdiDib leads to a blank
    ' alpha channel for 32bpp images, so use GdipCreateBitmapFromScan0 in that case.
    If tmpDIB.GetDIBColorDepth = 32 Then
        
        'Use GdipCreateBitmapFromScan0 to create a 32bpp DIB with alpha preserved
        GDIPlusReturn = GdipCreateBitmapFromScan0(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpDIB.GetDIBWidth * 4, PixelFormat32bppARGB, ByVal tmpDIB.GetDIBPointer, hImage)
    
    Else
        GDIPlusReturn = GdipCreateBitmapFromGdiDib(imgHeader, ByVal tmpDIB.GetDIBPointer, hImage)
    End If
    
    If (GDIPlusReturn <> 0) Then
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
    Dim uEncCLSID As clsid
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
                .Value = VarPtr(jpegQuality)
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
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(dstFilename) Then cFile.KillFile dstFilename
    
    Message "Saving the file..."
    
    'Perform the encode and save
    GDIPlusReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(uEncCLSID), aEncParams(1))
    
    If (GDIPlusReturn <> 0) Then
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
    
    'Use GDI+ to create a GDI+-compatible bitmap
    Dim GDIPlusReturn As Long
    Dim hGdipBitmap As Long
    GetGdipBitmapHandleFromDIB hGdipBitmap, srcDIB
        
    'Request a PNG encoder from GDI+
    Dim uEncCLSID As clsid
    Dim uEncParams As EncoderParameters
    Dim aEncParams() As Byte
        
    pvGetEncoderClsID "image/png", uEncCLSID
    uEncParams.Count = 1
    ReDim aEncParams(1 To Len(uEncParams))
    
    Dim gdipColorDepth As Long
    gdipColorDepth = srcDIB.GetDIBColorDepth
    
    With uEncParams.Parameter
        .NumberOfValues = 1
        .encType = EncoderParameterValueTypeLong
        .Guid = pvDEFINE_GUID(EncoderColorDepth)
        .Value = VarPtr(gdipColorDepth)
    End With
    
    CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
    
    'Check to see if a file already exists at this location
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    If cFile.FileExist(dstFilename) Then cFile.KillFile dstFilename
    
    'Perform the encode and save
    GDIPlusReturn = GdipSaveImageToFile(hGdipBitmap, StrPtr(dstFilename), VarPtr(uEncCLSID), aEncParams(1))
    
    If (GDIPlusReturn <> 0) Then
        GdipDisposeImage hGdipBitmap
        GDIPlusQuickSavePNG = False
        Exit Function
    End If
    
    'Release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hGdipBitmap)
    
    GDIPlusQuickSavePNG = True
    Exit Function
    
GDIPlusQuickSaveError:

    GDIPlusQuickSavePNG = False
    
End Function

'Given an arbitrary array of points, return a handle to a GDI+ region created from the closed shape formed by the points.
' Note that this function does not perform automatic management of the returned region.  The caller must release the region manually,
' using ReleaseGDIPlusRegion() below.
Public Function GetGDIPlusRegionFromPoints(ByVal numOfPoints As Long, ByVal ptrFloatArray As Long, Optional ByVal useFillMode As GP_FillMode = GP_FM_Winding, Optional ByVal useCurveMode As Boolean = False, Optional ByVal curveTension As Single) As Long

    'Start by creating a blank GDI+ path object.
    Dim gdipRegionHandle As Long, gdipPathHandle As Long
    GdipCreatePath useFillMode, gdipPathHandle
    
    'Populate the region with the polygon point array we were passed.
    If useCurveMode Then
        GdipAddPathClosedCurve2 gdipPathHandle, ptrFloatArray, numOfPoints, curveTension
    Else
        GdipAddPathPolygon gdipPathHandle, ptrFloatArray, numOfPoints
    End If
    
    'Use the path to create a region
    GdipCreateRegionPath gdipPathHandle, gdipRegionHandle
    
    'Release the path
    GdipDeletePath gdipPathHandle
    
    'Return the newly formed region handle
    GetGDIPlusRegionFromPoints = gdipRegionHandle

End Function

'I'm not sure whether a pure GDI+ solution or a manual solution is faster, but because the manual solution guarantees the
' smallest possible rect (unlike GDI+), I'm going with it for now.
Public Function IntersectRectF(ByRef dstRect As RECTF, ByRef srcRect1 As RECTF, ByRef srcRect2 As RECTF) As Boolean

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

'Given an arbitrary array of points, use GDI+ to find a bounding rect for the region created from the closed shape formed by the points.
' This function is self-managing, meaning it will delete any GDI+ objects it generates.
Public Function GetGDIPlusBoundingRectFromPoints(ByVal numOfPoints As Long, ByVal ptrFloatArray As Long, Optional ByVal useFillMode As GP_FillMode = GP_FM_Alternate, Optional ByVal useCurveMode As Boolean = False, Optional ByVal curveTension As Single, Optional ByVal penWidth As Single = 1#, Optional ByVal customLineCap As GP_LineCap = GP_LC_Flat) As RECTF

    'Start by creating a blank GDI+ path object.
    Dim gdipRegionHandle As Long, gdipPathHandle As Long
    GdipCreatePath useFillMode, gdipPathHandle
    
    'Populate the region with the polygon point array we were passed.
    If useCurveMode Then
        GdipAddPathClosedCurve2 gdipPathHandle, ptrFloatArray, numOfPoints, curveTension
    Else
        GdipAddPathPolygon gdipPathHandle, ptrFloatArray, numOfPoints
    End If
    
    'Create a pen object with width and linecaps matching the passed params; these are important in the bounds calculation, as a wider pen
    ' means a wider region.
    Dim iPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(0, 255), penWidth, GP_U_Pixel, iPen
    
    'If a custom line cap was specified, apply it now
    If customLineCap > 0 Then GdipSetPenLineCap iPen, customLineCap, customLineCap, 0&
    
    'Using the generated pen, calculate a bounding rect for the path as drawn with that pen
    GdipGetPathWorldBounds gdipPathHandle, GetGDIPlusBoundingRectFromPoints, 0, 0& 'iPen
    
    'Release the path and pen before exiting
    GdipDeletePath gdipPathHandle
    GdipDeletePen iPen
    
End Function

'Given an arbitrary array of points, and a pdImage handle, use GDI+ to find the union a rect of the path and the image.  This is relevant for shapes,
' which may be placed off the image, and we are only interested in the part the shape that actually overlaps the image itself.
Public Function GetGDIPlusUnionFromPointsAndImage(ByVal numOfPoints As Long, ByVal ptrFloatArray As Long, ByRef srcImage As pdImage, Optional ByVal useFillMode As GP_FillMode = GP_FM_Alternate, Optional ByVal useCurveMode As Boolean = False, Optional ByVal curveTension As Single) As RECTF

    'Start by creating a blank GDI+ path object.
    Dim gdipRegionHandle As Long, gdipPathHandle As Long
    GdipCreatePath useFillMode, gdipPathHandle
    
    'Populate the region with the polygon point array we were passed.
    If useCurveMode Then
        GdipAddPathClosedCurve2 gdipPathHandle, ptrFloatArray, numOfPoints, curveTension
    Else
        GdipAddPathPolygon gdipPathHandle, ptrFloatArray, numOfPoints
    End If
    
    'Convert the created path to a region.
    GdipCreateRegionPath gdipPathHandle, gdipRegionHandle
    
    'Next, create a rect that represents the bounds of the image
    Dim imgRect As RECTF
    With imgRect
        .Left = 0
        .Top = 0
        .Width = srcImage.Width
        .Height = srcImage.Height
    End With
    
    'Combine the image rect with the path region, using INTERSECT mode.
    GdipCombineRegionRect gdipRegionHandle, imgRect, GP_CM_Intersect
    
    'The region now contains only the union of the path and the region itself.  Retrive the region's bounds.
    
    'Start by creating a blank graphics object to supply to the region boundary check.  (This object normally contains any world transforms,
    ' but we don't care about transforms in this function.)
    Dim tmpSettingsDIB As pdDIB
    Set tmpSettingsDIB = New pdDIB
    tmpSettingsDIB.CreateBlank 8, 8, 32, 0, 0
    
    Dim tmpGraphics As Long
    If GdipCreateFromHDC(tmpSettingsDIB.GetDIBDC, tmpGraphics) = 0 Then
    
        'Retrieve the new bounding rect of the region, and place it directly into the function return
        GdipGetRegionBounds gdipRegionHandle, tmpGraphics, GetGDIPlusUnionFromPointsAndImage
        
        'Release our temporary graphics object
        GdipDeleteGraphics tmpGraphics
        
    End If
    
    'Release the region and path before exiting
    GdipDeleteRegion gdipRegionHandle
    GdipDeletePath gdipPathHandle
    
End Function

'Given a point and a region, return whether the point is inside or not inside the region.  Because GDI+ does not maintain the concept of
' "partially within a region", antialiasing has no effect here - only the "perfect" theoretical boundary of the region is used for hit-testing.
Public Function IsPointInGDIPlusRegion(ByVal srcX As Single, ByVal srcY As Single, ByRef regionHandle As Long) As Boolean
    
    'Use GDI+ to test the point
    Dim retLong As Long
    GdipIsVisibleRegionPoint regionHandle, srcX, srcY, 0&, retLong
    
    IsPointInGDIPlusRegion = (retLong = 1)
    
End Function

'Nearly identical to StretchBlt, but using GDI+ so we can:
' 1) support fractional source/dest/width/height
' 2) apply variable opacity
' 3) control stretch mode directly inside the call
Public Sub GDIPlus_StretchBlt(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal newAlpha As Single = 1#, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal useThisDestinationDCInstead As Long = 0, Optional ByVal disableEdgeFix As Boolean = False, Optional ByVal isZoomedIn As Boolean = False)

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    If (useThisDestinationDCInstead <> 0) Then
        GdipCreateFromHDC useThisDestinationDCInstead, iGraphics
    Else
        GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    End If
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, srcDIB
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(iGraphics, interpolationType) = GP_OK Then
        
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        GdipCreateImageAttributes imgAttributesHandle
        
        'To improve performance, explicitly request high-speed (aka linear) alpha compositing operation, and standard
        ' pixel offsets (on pixel borders, instead of center points)
        If (Not disableEdgeFix) Then GdipSetImageAttributesWrapMode imgAttributesHandle, GP_WM_TileFlipXY, 0, 0
        GdipSetCompositingQuality iGraphics, GP_CQ_AssumeLinear
        If isZoomedIn Then GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
        
        'If modified alpha is requested, pass the new value to this image container
        If (newAlpha <> 1) Then
            m_AttributesMatrix(3, 3) = newAlpha
            GdipSetImageAttributesColorMatrix imgAttributesHandle, ColorAdjustTypeBitmap, 1, VarPtr(m_AttributesMatrix(0, 0)), 0, ColorMatrixFlagsDefault
        End If
    
        'Perform the resize
        GdipDrawImageRectRect iGraphics, tBitmap, x1, y1, dstWidth, dstHeight, x2, y2, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle
        
        'Release our image attributes object
        GdipDisposeImageAttributes imgAttributesHandle
        
        'Reset alpha in the master identity matrix
        If (newAlpha <> 1) Then m_AttributesMatrix(3, 3) = 1
        
        'Update premultiplication status in the target
        If (Not (dstDIB Is Nothing)) Then dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
        
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDisposeImage tBitmap
    GdipDeleteGraphics iGraphics
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
    
End Sub

'Similar function to GDIPlus_StretchBlt, above, but using a destination parallelogram instead of a rect.
'
'Note that the supplied plgPoints array *MUST HAVE THREE POINTS* in it, in the specific order: top-left, top-right, bottom-left.
' The fourth point is inferred from the other three.
Public Sub GDIPlus_PlgBlt(ByRef dstDIB As pdDIB, ByRef plgPoints() As POINTFLOAT, ByRef srcDIB As pdDIB, ByVal x2 As Single, ByVal y2 As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal newAlpha As Single = 1#, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal useHQOffsets As Boolean = True)

    'Because this function is such a crucial part of PD's render chain, I occasionally like to profile it against
    ' viewport engine changes.  Uncomment the two lines below, and the reporting line at the end of the sub to
    ' have timing reports sent to the debug window.
    'Dim profileTime As Double
    'profileTime = Timer
    
    'Create a GDI+ graphics object that points to the destination DIB's DC
    Dim iGraphics As Long, tBitmap As Long
    GdipCreateFromHDC dstDIB.GetDIBDC, iGraphics
    
    'Next, we need a copy of the source image (in GDI+ Bitmap format) to use as our source image reference.
    ' 32bpp and 24bpp are handled separately, to ensure alpha preservation for 32bpp images.
    GetGdipBitmapHandleFromDIB tBitmap, srcDIB
    
    'iGraphics now contains a pointer to the destination image, while tBitmap contains a pointer to the source image.
    
    'Request the smoothing mode we were passed
    If GdipSetInterpolationMode(iGraphics, interpolationType) = GP_OK Then
        
        'To fix antialiased fringing around image edges, specify a wrap mode.  This will prevent the faulty GDI+ resize
        ' algorithm from drawing semi-transparent lines randomly around image borders.
        ' Thank you to http://stackoverflow.com/questions/1890605/ghost-borders-ringing-when-resizing-in-gdi for the fix.
        Dim imgAttributesHandle As Long
        If (newAlpha <> 1) Then GdipCreateImageAttributes imgAttributesHandle Else imgAttributesHandle = 0
        
        'To improve performance and quality, explicitly request high-speed (aka linear) alpha compositing operation, and high-quality
        ' pixel offsets (treat pixels as if they fall on pixel borders, instead of center points - this provides rudimentary edge
        ' antialiasing, which is the best we can do without murdering performance)
        GdipSetCompositingQuality iGraphics, GP_CQ_AssumeLinear
        If useHQOffsets Then GdipSetPixelOffsetMode iGraphics, GP_POM_HighQuality Else GdipSetPixelOffsetMode iGraphics, GP_POM_HighSpeed
        
        'If modified alpha is requested, pass the new value to this image container
        If (newAlpha <> 1) Then
            m_AttributesMatrix(3, 3) = newAlpha
            GdipSetImageAttributesColorMatrix imgAttributesHandle, ColorAdjustTypeBitmap, 1, VarPtr(m_AttributesMatrix(0, 0)), 0, ColorMatrixFlagsDefault
        End If
        
        'Perform the draw
        GdipDrawImagePointsRect iGraphics, tBitmap, VarPtr(plgPoints(0)), 3, x2, y2, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&
        
        'Release our image attributes object
        If (imgAttributesHandle <> 0) Then GdipDisposeImageAttributes imgAttributesHandle
        
        'Reset alpha in the master identity matrix
        If (newAlpha <> 1) Then m_AttributesMatrix(3, 3) = 1
        
        'Update premultiplication status in the target
        dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
        
    End If
    
    'Release both the destination graphics object and the source bitmap object
    GdipDisposeImage tBitmap
    GdipDeleteGraphics iGraphics
    
    'Uncomment the line below to receive timing reports
    'Debug.Print Format(CStr((Timer - profileTime) * 1000), "0000.00")
    
End Sub

'Given a source DIB and an angle, rotate it into a destination DIB.  The destination DIB can be automatically resized
' to fit the rotated image, or a parameter can be set, instructing the function to use the destination DIB "as-is"
Public Sub GDIPlus_RotateDIBPlgStyle(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotateAngle As Single, Optional ByVal dstDIBAlreadySized As Boolean = False, Optional ByVal rotateQuality As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal transparentBackground As Boolean = True, Optional ByVal newBackColor As Long = vbWhite)
    
    'Before doing any rotating or blurring, we need to figure out the size of our destination image.  If we don't
    ' do this, the rotation will chop off the image's corners!
    Dim nWidth As Double, nHeight As Double
    Math_Functions.FindBoundarySizeOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, nWidth, nHeight, False
    
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
    Dim listOfPoints() As POINTFLOAT
    ReDim listOfPoints(0 To 3) As POINTFLOAT
    Math_Functions.FindCornersOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, listOfPoints, True
    
    'Calculate the size difference between the source and destination images.  We need to add this offset to all
    ' rotation coordinates, to ensure the rotated image is fully contained within the destination DIB.
    Dim hOffset As Double, vOffset As Double
    hOffset = (nWidth - srcDIB.GetDIBWidth) / 2
    vOffset = (nHeight - srcDIB.GetDIBHeight) / 2
    
    'Apply those offsets to all rotation points, and because GDI+ requires us to use an offset pixel mode for
    ' non-shit results along edges, pad all coordinates with an extra half-pixel as well.
    Dim i As Long
    For i = 0 To 3
        listOfPoints(i).x = listOfPoints(i).x + 0.5 + hOffset
        listOfPoints(i).y = listOfPoints(i).y + 0.5 + vOffset
    Next i
    
    'If a background color is being applied, "cut out" the target region now
    If (Not transparentBackground) Then
        
        Dim tmpPoints() As POINTFLOAT
        ReDim tmpPoints(0 To 3) As POINTFLOAT
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
        cx = cx / 4
        cy = cy / 4
        
        'For each corner of the rotated square, convert the point to polar coordinates, then shrink the radius by one.
        Dim tmpAngle As Double, tmpRadius As Double, tmpX As Double, tmpY As Double
        For i = 0 To 3
            Math_Functions.ConvertCartesianToPolar tmpPoints(i).x, tmpPoints(i).y, tmpRadius, tmpAngle, cx, cy
            tmpRadius = tmpRadius - 1#
            Math_Functions.ConvertPolarToCartesian tmpAngle, tmpRadius, tmpX, tmpY, cx, cy
            tmpPoints(i).x = tmpX
            tmpPoints(i).y = tmpY
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
    GDI_Plus.GDIPlus_PlgBlt dstDIB, listOfPoints, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 1, rotateQuality, True
    
End Sub

'Given a regular ol' DIB and an angle, return a DIB that is rotated by that angle, with its edge values clamped and extended
' to fill all empty space around the rotated image.  This very cool operation allows us to support angles for any filter
' with a grid implementation (e.g. something that operates on the (x, y) axes of an image, like pixellate or blur).
Public Sub GDIPlus_GetRotatedClampedDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal rotateAngle As Single)

    'Before doing any rotating or blurring, we need to figure out the size of our destination image.  If we don't
    ' do this, the rotation will chop off the image's corners!
    Dim nWidth As Double, nHeight As Double
    Math_Functions.FindBoundarySizeOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, nWidth, nHeight
    
    'Use these dimensions to size the destination image
    If dstDIB Is Nothing Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> nWidth) Or (dstDIB.GetDIBHeight <> nHeight) Or (dstDIB.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
        dstDIB.CreateBlank nWidth, nHeight, srcDIB.GetDIBColorDepth, 0, 0
    Else
        dstDIB.ResetDIB 0
    End If
    
    'We also want a copy of the corner points of the rotated rect; we'll use these to perform a fast PlgBlt-like operation,
    ' which is how we draw both the rotation and the corner extensions.
    Dim listOfPoints() As POINTFLOAT
    ReDim listOfPoints(0 To 3) As POINTFLOAT
    Math_Functions.FindCornersOfRotatedRect srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, rotateAngle, listOfPoints, True
    
    'Calculate the size difference between the source and destination images.  We need to add this offset to all
    ' rotation coordinates, to ensure the rotated image is fully contained within the destination DIB.
    Dim hOffset As Double, vOffset As Double
    hOffset = (nWidth - srcDIB.GetDIBWidth) / 2
    vOffset = (nHeight - srcDIB.GetDIBHeight) / 2
    
    'Apply those offsets to all rotation points, and because GDI+ requires us to use an offset pixel mode for
    ' non-shit results along edges, pad all coordinates with an extra half-pixel as well.
    Dim i As Long
    For i = 0 To 3
        listOfPoints(i).x = listOfPoints(i).x + hOffset '+ 0.5
        listOfPoints(i).y = listOfPoints(i).y + vOffset '+ 0.5
    Next i
    
    'Rotate the source DIB into the destination DIB.  At this point, corners are still blank - we'll deal with those momentarily.
    GDI_Plus.GDIPlus_PlgBlt dstDIB, listOfPoints, srcDIB, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 1, GP_IM_HighQualityBicubic, False
    
    'We're now going to calculate a whole bunch of geometry based around the concept of extending a rectangle from
    ' each edge of our rotated image, out to the corner of the rotation DIB.  We will then fill this dead space with a
    ' stretched version of the image edge, resulting in "clamped" edge behavior.
    Dim diagDistance As Double, distDiff As Double
    Dim dx As Double, dy As Double, lineLength As Double, pX As Double, pY As Double
    Dim padPoints() As POINTFLOAT
    ReDim padPoints(0 To 2) As POINTFLOAT
    
    'Calculate the distance from the center of the rotated image to the corner of the rotated image
    diagDistance = Sqr(nWidth * nWidth + nHeight * nHeight) / 2
    
    'Get the difference between the diagonal distance, and the original height of the image.  This is the distance
    ' where we need to provide clamped pixels on this edge.
    distDiff = diagDistance - (srcDIB.GetDIBHeight / 2)
    
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
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, 0, srcDIB.GetDIBWidth, 1, 1, GP_IM_HighQualityBilinear, False
    
    'Now repeat the above steps for the bottom of the image.  Note that we can reuse almost all of the calculations,
    ' as this line is parallel to the one we just calculated.
    padPoints(0).x = listOfPoints(2).x - (pX / distDiff)
    padPoints(0).y = listOfPoints(2).y - (pY / distDiff)
    padPoints(1).x = listOfPoints(3).x - (pX / distDiff)
    padPoints(1).y = listOfPoints(3).y - (pY / distDiff)
    padPoints(2).x = listOfPoints(2).x + pX
    padPoints(2).y = listOfPoints(2).y + pY
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, srcDIB.GetDIBHeight - 2, srcDIB.GetDIBWidth, 1, 1, GP_IM_HighQualityBilinear, False
    
    'We are now going to repeat the above steps, but for the left and right edges of the image.  The end result of this
    ' will be a rotated destination image, with clamped values extending from all image edges.
    
    'Get the difference between the diagonal distance, and the original width of the image.  This is the distance
    ' where we need to provide clamped pixels on this edge.
    distDiff = diagDistance - (srcDIB.GetDIBWidth / 2)
    
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
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, 0, 0, 1, srcDIB.GetDIBHeight, 1, GP_IM_HighQualityBilinear, False
    
    '...and finally, repeat everything for the right side of the image
    padPoints(0).x = listOfPoints(1).x + (pX / distDiff)
    padPoints(0).y = listOfPoints(1).y + (pY / distDiff)
    padPoints(1).x = listOfPoints(1).x - pX
    padPoints(1).y = listOfPoints(1).y - pY
    padPoints(2).x = listOfPoints(3).x + (pX / distDiff)
    padPoints(2).y = listOfPoints(3).y + (pY / distDiff)
    GDI_Plus.GDIPlus_PlgBlt dstDIB, padPoints, srcDIB, srcDIB.GetDIBWidth - 2, 0, 1, srcDIB.GetDIBHeight, 1, GP_IM_HighQualityBilinear, False
    
    'Our work here is complete!

End Sub

'Thanks to Carles P.V. for providing the following four functions, which are used as part of GDI+ image saving.
' You can download Carles's full project from http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
Private Function pvGetEncoderClsID(strMimeType As String, ClassID As clsid) As Long

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

Private Function pvDEFINE_GUID(ByVal sGuid As String) As clsid
'-- Courtesy of: Dana Seaman
'   Helper routine to convert a CLSID(aka GUID) string to a structure
'   Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
    Call CLSIDFromString(StrPtr(sGuid), VarPtr(pvDEFINE_GUID))
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

'Given a GUID string, return a Long-type image format identifier
Private Function GetFIFFromGUID(ByRef srcGUID As String) As PHOTODEMON_IMAGE_FORMAT

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
    GDIP_StartEngine = CBool(GdiplusStartup(m_GDIPlusToken, gdiCheck, 0&) = GP_OK)
    If GDIP_StartEngine Then
        
        'As a convenience, create a dummy graphics container.  This is useful for various GDI+ functions that require world
        ' transformation data.
        Set m_TransformDIB = New pdDIB
        m_TransformDIB.CreateBlank 8, 8, 32, 0, 0
        GdipCreateFromHDC m_TransformDIB.GetDIBDC, m_TransformGraphics
        
        'Note that these dummy objects are released when GDI+ terminates.
        
        'Next, create a default identity matrix for image attributes.
        ReDim m_AttributesMatrix(0 To 4, 0 To 4) As Single
        m_AttributesMatrix(0, 0) = 1#
        m_AttributesMatrix(1, 1) = 1#
        m_AttributesMatrix(2, 2) = 1#
        m_AttributesMatrix(3, 3) = 1#
        m_AttributesMatrix(4, 4) = 1#
        
        'Next, check to see if v1.1 is available.  This allows for advanced fx work.
        Dim hMod As Long, strGDIPName As String
        strGDIPName = "gdiplus.dll"
        hMod = LoadLibrary(StrPtr(strGDIPName))
        If (hMod <> 0) Then
            Dim testAddress As Long
            testAddress = GetProcAddress(hMod, "GdipDrawImageFX")
            m_GDIPlus11Available = CBool(testAddress <> 0)
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

    'Release any dummy containers we have created
    GdipDeleteGraphics m_TransformGraphics
    Set m_TransformDIB = Nothing
    
    'Release GDI+ using the same token we received at startup time
    GDIP_StopEngine = CBool(GdiplusShutdown(m_GDIPlusToken) = GP_OK)
    
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
Private Function GDIP_Debug_Proc(ByVal deLevel As GP_DebugEventLevel, ByVal ptrChar As Long) As Long
    
    'Pull the GDI+ message into a local string
    'Dim cUnicode As pdUnicode
    'Set cUnicode = New pdUnicode
    
    Dim debugString As String
    'debugString = cUnicode.ConvertCharPointerToVBString(ptrChar, False)
    debugString = "Unknown GDI+ error was passed to the GDIPlus debug procedure."
    
    If (deLevel = GP_DebugEventLevelWarning) Then
        Debug.Print "GDI+ WARNING: " & debugString
    ElseIf (deLevel = GP_DebugEventLevelFatal) Then
        Debug.Print "GDI+ ERROR: " & debugString
    Else
        Debug.Print "GDI+ UNKNOWN: " & debugString
    End If
    
End Function

Private Function InternalGDIPlusError(Optional ByVal errName As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal errNumber As GP_Result = GP_OK)
        
    'If the caller passes an error number but no error name, attempt to automatically populate
    ' it based on the error number.
    If ((Len(errName) = 0) And (errNumber <> GP_OK)) Then
        
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
    
    If (Len(errDescription) <> 0) Then tmpString = tmpString & ": " & errDescription
    Debug.Print tmpString
    
End Function

'GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from a standard RGB()
' long and supplied alpha.  It's not a very efficient conversion, but I need it so infrequently that I don't really care.
Public Function FillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
    
    'The vbRGB constant may be an OLE color constant; if that happens, we want to convert it to a normal RGB quad.
    vbRGB = TranslateColor(vbRGB)
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = Drawing2D.ExtractRed(vbRGB)
    dstQuad.Green = Drawing2D.ExtractGreen(vbRGB)
    dstQuad.Blue = Drawing2D.ExtractBlue(vbRGB)
    dstQuad.alpha = alphaValue
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    FillQuadWithVBRGB = placeHolder.lngResult
    
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the opacity value on the scale [0, 100]
Public Function GetOpacityFromPARGB(ByVal pARGB As Long) As Single
    Dim srcQuad As RGBQUAD
    CopyMemory_Strict VarPtr(srcQuad), VarPtr(pARGB), 4&
    GetOpacityFromPARGB = CSng(srcQuad.alpha) * CSng(100# / 255#)
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the RGB component in combined vbRGB format
Public Function GetColorFromPARGB(ByVal pARGB As Long) As Long
    
    Dim srcQuad As RGBQUAD
    CopyMemory_Strict VarPtr(srcQuad), VarPtr(pARGB), 4&
    
    If (srcQuad.alpha = 255) Then
        GetColorFromPARGB = RGB(srcQuad.Red, srcQuad.Green, srcQuad.Blue)
    Else
    
        Dim tmpSingle As Single
        tmpSingle = CSng(srcQuad.alpha) / 255
        
        If (tmpSingle <> 0) Then
            Dim tmpRed As Long, tmpGreen As Long, tmpBlue As Long
            tmpRed = CSng(srcQuad.Red) / tmpSingle
            tmpGreen = CSng(srcQuad.Green) / tmpSingle
            tmpBlue = CSng(srcQuad.Blue) / tmpSingle
            GetColorFromPARGB = RGB(tmpRed, tmpGreen, tmpBlue)
        Else
            GetColorFromPARGB = 0
        End If
        
    End If
    
End Function

'Translate an OLE color to an RGB Long.  Note that the API function returns -1 on failure; if this happens, we return white.
Private Function TranslateColor(ByVal colorRef As Long) As Long
    If OleTranslateColor(colorRef, 0, TranslateColor) Then TranslateColor = vbWhite
End Function

Public Function GetGDIPlusSolidBrushHandle(ByVal brushColor As Long, Optional ByVal brushOpacity As Byte = 255) As Long
    GdipCreateSolidFill FillQuadWithVBRGB(brushColor, brushOpacity), GetGDIPlusSolidBrushHandle
End Function

Public Function GetGDIPlusPatternBrushHandle(ByVal brushPattern As GP_PatternStyle, ByVal bFirstColor As Long, ByVal bFirstColorOpacity As Byte, ByVal bSecondColor As Long, ByVal bSecondColorOpacity As Byte) As Long
    GdipCreateHatchBrush brushPattern, FillQuadWithVBRGB(bFirstColor, bFirstColorOpacity), FillQuadWithVBRGB(bSecondColor, bSecondColorOpacity), GetGDIPlusPatternBrushHandle
End Function

Public Function GetGDIPlusLinearBrushHandle(ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As PD_2D_WrapMode) As Long
    GdipCreateLineBrushFromRectWithAngle srcRect, firstRGBA, secondRGBA, gradAngle, isAngleScalable, gradientWrapMode, GetGDIPlusLinearBrushHandle
End Function

Public Function OverrideGDIPlusLinearGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusLinearGradient = CBool(GdipSetLinePresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

Public Function GetGDIPlusPathBrushHandle(ByVal hGraphicsPath As Long) As Long
    GdipCreatePathGradientFromPath hGraphicsPath, GetGDIPlusPathBrushHandle
End Function

Public Function SetGDIPlusPathBrushCenter(ByVal hBrush As Long, ByVal centerX As Single, ByVal centerY As Single) As Long
    Dim centerPoint As POINTFLOAT
    centerPoint.x = centerX
    centerPoint.y = centerY
    GdipSetPathGradientCenterPoint hBrush, centerPoint
End Function

Public Function SetGDIPlusPathBrushWrap(ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As Boolean
    SetGDIPlusPathBrushWrap = CBool(GdipSetPathGradientWrapMode(hBrush, newWrapMode) = GP_OK)
End Function

Public Function OverrideGDIPlusPathGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusPathGradient = CBool(GdipSetPathGradientPresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

'Simpler shorthand function for obtaining a GDI+ bitmap handle from a pdDIB object.  Note that 24/32bpp cases have to be
' handled separately because GDI+ is unpredictable at automatically detecting color depth with 32-bpp DIBs.  (This behavior
' is forgivable, given GDI's unreliable handling of alpha bytes.)
Public Function GetGdipBitmapHandleFromDIB(ByRef dstBitmapHandle As Long, ByRef srcDIB As pdDIB) As Boolean
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetGdipBitmapHandleFromDIB = CBool(GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBWidth * 4, GP_PF_32bppPARGB, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
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
        GetGdipBitmapHandleFromDIB = CBool(GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
        
    End If

End Function

'Retrieving a bitmap from a DC is a messy and performance-intensive process.  Avoid it if at all possible.
Public Function GetGdipBitmapHandleFromDC(ByVal srcDC As Long) As Long

End Function

'Because of the way GDI+ texture brushes work, it is significantly easier to initialize one from a full DIB object
' (which *always* guarantees bitmap bits will be available) vs a GDI+ Graphics object, which is more like a DC in
' that it could be a non-bitmap, or dimensionless, or other weird criteria.
Public Function GetGDIPlusTextureBrush(ByRef srcDIB As pdDIB, Optional ByVal brushWrapMode As GP_WrapMode = GP_WM_Tile) As Long
    Dim srcBitmap As Long, tmpReturn As GP_Result
    GetGdipBitmapHandleFromDIB srcBitmap, srcDIB
    tmpReturn = GdipCreateTexture(srcBitmap, brushWrapMode, GetGDIPlusTextureBrush)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError , , tmpReturn
    tmpReturn = GdipDisposeImage(srcBitmap)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError , , tmpReturn
End Function

'Retrieve a persistent handle to a GDI+-format graphics container.  Optionally, a smoothing mode can be specified so that it does
' not have to be repeatedly specified by a caller function.  (GDI+ sets smoothing mode by graphics container, not by function call.)
Public Function GetGDIPlusGraphicsFromDC(ByVal srcDC As Long, Optional ByVal graphicsAntialiasing As GP_SmoothingMode = GP_SM_None, Optional ByVal graphicsPixelOffsetMode As GP_PixelOffsetMode = GP_POM_None) As Long
    Dim hGraphics As Long
    If (GdipCreateFromHDC(srcDC, hGraphics) = GP_OK) Then
        SetGDIPlusGraphicsProperty hGraphics, P2_SurfaceAntialiasing, graphicsAntialiasing
        SetGDIPlusGraphicsProperty hGraphics, P2_SurfacePixelOffset, graphicsPixelOffsetMode
        GetGDIPlusGraphicsFromDC = hGraphics
    Else
        GetGDIPlusGraphicsFromDC = 0
    End If
End Function

'Shorthand function for quickly creating a new GDI+ pen.  This can be useful if many drawing operations are going to be applied with the same pen.
' (Note that a single parameter is used to set both pen and dash endcaps; if you want these to differ, you must call the separate
' SetPenDashCap function, below.)
Public Function GetGDIPlusPenHandle(ByVal penColor As Long, Optional ByVal penOpacity As Long = 255&, Optional ByVal penWidth As Single = 1#, Optional ByVal penLineCap As GP_LineCap = GP_LC_Flat, Optional ByVal penLineJoin As GP_LineJoin = GP_LJ_Miter, Optional ByVal penDashMode As GP_DashStyle = GP_DS_Solid, Optional ByVal penMiterLimit As Single = 3#, Optional ByVal penAlignment As GP_PenAlignment = GP_PA_Center) As Long

    'Create the base pen
    Dim hPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(penColor, penOpacity), penWidth, GP_U_Pixel, hPen
    
    If (hPen <> 0) Then
        
        GdipSetPenLineCap hPen, penLineCap, penLineCap, 0&
        GdipSetPenLineJoin hPen, penLineJoin
        
        If (penDashMode <> GP_DS_Solid) Then
            
            GdipSetPenDashStyle hPen, penDashMode
            
            'Mirror the line cap across the dashes as well
            If (penLineCap = GP_LC_ArrowAnchor) Or (penLineCap = GP_LC_DiamondAnchor) Then
                GdipSetPenDashCap hPen, GP_DC_Triangle
            ElseIf (penLineCap = GP_LC_Round) Or (penLineCap = GP_LC_RoundAnchor) Then
                GdipSetPenDashCap hPen, GP_DC_Round
            Else
                GdipSetPenDashCap hPen, GP_DC_Flat
            End If
            
        End If
        
        'To avoid major miter errors, we default to 3.0 for a miter limit.  (GDI+ defaults to 10, which can easily cause artifacts.)
        GdipSetPenMiterLimit hPen, penMiterLimit
        
        'Finally, if a non-standard alignment was specified, apply it last
        If (penAlignment <> GP_PA_Center) Then GdipSetPenMode hPen, penAlignment
        
    End If
    
    GetGDIPlusPenHandle = hPen

End Function

Public Function GetGDIPlusPenFromBrush(ByVal hBrush As Long, ByVal penWidth As Single, Optional ByVal penUnit As GP_Unit = GP_U_Pixel) As Long
    GdipCreatePenFromBrush hBrush, penWidth, penUnit, GetGDIPlusPenFromBrush
End Function

Public Function GetGDIPlusRegionHandle() As Long
    GdipCreateRegion GetGDIPlusRegionHandle
End Function

Public Function ReleaseGDIPlusBrush(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusBrush = CBool(GdipDeleteBrush(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusBrush = True
    End If
End Function

Public Function ReleaseGDIPlusGraphics(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusGraphics = CBool(GdipDeleteGraphics(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusGraphics = True
    End If
End Function

Public Function ReleaseGDIPlusImage(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusImage = CBool(GdipDisposeImage(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusImage = True
    End If
End Function

Public Function ReleaseGDIPlusPen(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusPen = CBool(GdipDeletePen(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusPen = True
    End If
End Function

Public Function ReleaseGDIPlusRegion(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusRegion = CBool(GdipDeleteRegion(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusRegion = True
    End If
End Function

'NOTE!  ALL OPACITY SETTINGS are treated as singles on the range [0, 100], *not* as bytes on the range [0, 255].
'NOTE!  When getting or setting brush settings, you need to make sure the current brush type matches.  For example: if your
'       brush handle points to a solid brush, getting/setting its pattern style is meaningless.  You need to set the
'       relevant brush mode PRIOR to getting/setting other settings.
'NOTE!  Some brush settings cannot be set or retrieved.  For example, GDI+ does not allow you to change hatch style, color,
'       or opacity after brush creation.  You must create a new brush from scratch.  If you use the pd2DBrush class instead
'       of interfacing with these functions directly, nuances like this are handled automatically.
Public Function GetGDIPlusBrushProperty(ByVal hBrush As Long, ByVal propID As PD_2D_BRUSH_SETTINGS) As Variant
    
    If (hBrush <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong As Long, tmpSingle As Single
        
        Select Case propID
            
            'GDI+ does provide a function for this, but their enums differ from ours (by design).
            ' As such, you cannot set brush mode with this function; use the pd2DBrush class, instead.
            Case P2_BrushMode
                GetGDIPlusBrushProperty = 0&
                
           Case P2_BrushColor
                gResult = GdipGetSolidFillColor(hBrush, tmpLong)
                GetGDIPlusBrushProperty = GetColorFromPARGB(tmpLong)
                
            Case P2_BrushOpacity
                gResult = GdipGetSolidFillColor(hBrush, tmpLong)
                GetGDIPlusBrushProperty = GetOpacityFromPARGB(tmpLong)
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPatternStyle
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Color
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Opacity
                GetGDIPlusBrushProperty = 0#
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Color
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Opacity
                GetGDIPlusBrushProperty = 0#
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAllSettings
                GetGDIPlusBrushProperty = vbNullString
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientShape
                GetGDIPlusBrushProperty = P2_GS_Linear
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAngle
                GetGDIPlusBrushProperty = 0#
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientWrapMode
                GetGDIPlusBrushProperty = P2_WM_TileFlipXY
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientNodes
                GetGDIPlusBrushProperty = vbNullString
                
            Case P2_BrushTextureWrapMode
                gResult = GdipGetTextureWrapMode(hBrush, tmpLong)
                GetGDIPlusBrushProperty = tmpLong
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusBrushProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusBrushProperty Error", "Null brush handle"
    End If
    
End Function

Public Function SetGDIPlusBrushProperty(ByVal hBrush As Long, ByVal propID As PD_2D_BRUSH_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hBrush <> 0) Then
        
        Dim tmpColor As Long, tmpOpacity As Single
        
        Select Case propID
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushMode
                SetGDIPlusBrushProperty = False
                
            Case P2_BrushColor
                tmpOpacity = GetGDIPlusBrushProperty(hBrush, P2_BrushOpacity)
                SetGDIPlusBrushProperty = CBool(GdipSetSolidFillColor(hBrush, FillQuadWithVBRGB(CLng(newSetting), tmpOpacity * 2.55)) = GP_OK)
                
            Case P2_BrushOpacity
                tmpColor = GetGDIPlusBrushProperty(hBrush, P2_BrushColor)
                SetGDIPlusBrushProperty = CBool(GdipSetSolidFillColor(hBrush, FillQuadWithVBRGB(tmpColor, CSng(newSetting) * 2.55)) = GP_OK)
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPatternStyle
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Color
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Opacity
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Color
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Opacity
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAllSettings
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientShape
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAngle
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientWrapMode
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientNodes
                SetGDIPlusBrushProperty = False
                
            Case P2_BrushTextureWrapMode
                SetGDIPlusBrushProperty = CBool(GdipSetTextureWrapMode(hBrush, CLng(newSetting)) = GP_OK)
                
        End Select
    
    Else
        InternalGDIPlusError "SetGDIPlusBrushProperty Error", "Null brush handle"
    End If
    
End Function

Public Function GetGDIPlusGraphicsProperty(ByVal hGraphics As Long, ByVal propID As PD_2D_SURFACE_SETTINGS) As Variant
    
    If (hGraphics <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong1 As Long, tmpLong2 As Long
        
        Select Case propID
            
            Case P2_SurfaceAntialiasing
                gResult = GdipGetSmoothingMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfacePixelOffset
                gResult = GdipGetPixelOffsetMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfaceRenderingOriginX
                gResult = GdipGetRenderingOrigin(hGraphics, tmpLong1, tmpLong2)
                GetGDIPlusGraphicsProperty = tmpLong1
            
            Case P2_SurfaceRenderingOriginY
                gResult = GdipGetRenderingOrigin(hGraphics, tmpLong1, tmpLong2)
                GetGDIPlusGraphicsProperty = tmpLong2
                
            Case P2_SurfaceBlendUsingSRGBGamma
                gResult = GdipGetCompositingQuality(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfaceResizeQuality
                gResult = GdipGetInterpolationMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Null graphics handle"
    End If
    
End Function

Public Function SetGDIPlusGraphicsProperty(ByVal hGraphics As Long, ByVal propID As PD_2D_SURFACE_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hGraphics <> 0) Then
        
        Select Case propID
            
            Case P2_SurfaceAntialiasing
                SetGDIPlusGraphicsProperty = CBool(GdipSetSmoothingMode(hGraphics, CLng(newSetting)) = GP_OK)
                
            Case P2_SurfacePixelOffset
                SetGDIPlusGraphicsProperty = CBool(GdipSetPixelOffsetMode(hGraphics, CLng(newSetting)) = GP_OK)
            
            Case P2_SurfaceRenderingOriginX
                SetGDIPlusGraphicsProperty = CBool(GdipSetRenderingOrigin(hGraphics, CLng(newSetting), GetGDIPlusGraphicsProperty(hGraphics, P2_SurfaceRenderingOriginY)) = GP_OK)
            
            Case P2_SurfaceRenderingOriginY
                SetGDIPlusGraphicsProperty = CBool(GdipSetRenderingOrigin(hGraphics, GetGDIPlusGraphicsProperty(hGraphics, P2_SurfaceRenderingOriginX), CLng(newSetting)) = GP_OK)
                
            Case P2_SurfaceBlendUsingSRGBGamma
                SetGDIPlusGraphicsProperty = CBool(GdipSetCompositingQuality(hGraphics, CLng(newSetting)) = GP_OK)
                
            Case P2_SurfaceResizeQuality
                SetGDIPlusGraphicsProperty = CBool(GdipSetInterpolationMode(hGraphics, CLng(newSetting)) = GP_OK)
            
        End Select
    
    Else
        InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Null graphics handle"
    End If
    
End Function

'NOTE!  PEN OPACITY setting is treated as a single on the range [0, 100], *not* as a byte on the range [0, 255]
Public Function GetGDIPlusPenProperty(ByVal hPen As Long, ByVal propID As PD_2D_PEN_SETTINGS) As Variant
    
    If (hPen <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong As Long, tmpSingle As Single
        
        Select Case propID
            
            Case P2_PenStyle
                gResult = GdipGetPenDashStyle(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
            
            Case P2_PenColor
                gResult = GdipGetPenColor(hPen, tmpLong)
                GetGDIPlusPenProperty = GetColorFromPARGB(tmpLong)
                
            Case P2_PenOpacity
                gResult = GdipGetPenColor(hPen, tmpLong)
                GetGDIPlusPenProperty = GetOpacityFromPARGB(tmpLong)
                
            Case P2_PenWidth
                gResult = GdipGetPenWidth(hPen, tmpSingle)
                GetGDIPlusPenProperty = tmpSingle
                
            Case P2_PenLineJoin
                gResult = GdipGetPenLineJoin(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenLineCap
                gResult = GdipGetPenStartCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenDashCap
                gResult = GdipGetPenDashCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenMiterLimit
                gResult = GdipGetPenMiterLimit(hPen, tmpSingle)
                GetGDIPlusPenProperty = tmpSingle
                
            Case P2_PenAlignment
                gResult = GdipGetPenMode(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenStartCap
                gResult = GdipGetPenStartCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
            
            Case P2_PenEndCap
                gResult = GdipGetPenEndCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusPenProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusPenProperty Error", "Null pen handle"
    End If
    
End Function

'NOTE!  PEN OPACITY setting is treated as a single on the range [0, 100], *not* as a byte on the range [0, 255]
Public Function SetGDIPlusPenProperty(ByVal hPen As Long, ByVal propID As PD_2D_PEN_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hPen <> 0) Then
        
        Dim tmpColor As Long, tmpOpacity As Single, tmpLong As Long
        
        Select Case propID
            
            Case P2_PenStyle
                SetGDIPlusPenProperty = CBool(GdipSetPenDashStyle(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenColor
                tmpOpacity = GetGDIPlusPenProperty(hPen, P2_PenOpacity)
                SetGDIPlusPenProperty = CBool(GdipSetPenColor(hPen, FillQuadWithVBRGB(CLng(newSetting), tmpOpacity * 2.55)) = GP_OK)
                
            Case P2_PenOpacity
                tmpColor = GetGDIPlusPenProperty(hPen, P2_PenColor)
                SetGDIPlusPenProperty = CBool(GdipSetPenColor(hPen, FillQuadWithVBRGB(tmpColor, CSng(newSetting) * 2.55)) = GP_OK)
                
            Case P2_PenWidth
                SetGDIPlusPenProperty = CBool(GdipSetPenDashStyle(hPen, CSng(newSetting)) = GP_OK)
                
            Case P2_PenLineJoin
                SetGDIPlusPenProperty = CBool(GdipSetPenLineJoin(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenLineCap
                tmpLong = GetGDIPlusPenProperty(hPen, P2_PenDashCap)
                SetGDIPlusPenProperty = CBool(GdipSetPenLineCap(hPen, CLng(newSetting), CLng(newSetting), tmpLong) = GP_OK)
                
            Case P2_PenDashCap
                SetGDIPlusPenProperty = CBool(GdipSetPenDashCap(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenMiterLimit
                SetGDIPlusPenProperty = CBool(GdipSetPenMiterLimit(hPen, CSng(newSetting)) = GP_OK)
                
            Case P2_PenAlignment
                SetGDIPlusPenProperty = CBool(GdipSetPenMode(hPen, CLng(newSetting)) = GP_OK)
            
            Case P2_PenStartCap
                SetGDIPlusPenProperty = CBool(GdipSetPenStartCap(hPen, CLng(newSetting)) = GP_OK)
            
            Case P2_PenEndCap
                SetGDIPlusPenProperty = CBool(GdipSetPenEndCap(hPen, CLng(newSetting)) = GP_OK)
                
        End Select
    
    Else
        InternalGDIPlusError "SetGDIPlusPenProperty Error", "Null pen handle"
    End If
    
End Function

'All generic draw and fill functions follow

'GDI+ arcs use bounding boxes to describe their placement.  As such, we manually convert the incoming centerX/Y and radius values
' to bounding box coordinates.
Public Function GDIPlus_DrawArcF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    GDIPlus_DrawArcF = CBool(GdipDrawArc(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawArcI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    GDIPlus_DrawArcI = CBool(GdipDrawArcI(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawClosedCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawClosedCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImage(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageI(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawImageRectRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawImagePointsRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As POINTFLOAT, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRect(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawImagePointsRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As POINTLONG, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRectI(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawLineF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    GDIPlus_DrawLineF = CBool(GdipDrawLine(dstGraphics, srcPen, x1, y1, x2, y2) = GP_OK)
End Function

Public Function GDIPlus_DrawLineI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    GDIPlus_DrawLineI = CBool(GdipDrawLineI(dstGraphics, srcPen, x1, y1, x2, y2) = GP_OK)
End Function

Public Function GDIPlus_DrawLinesF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLines(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawLinesF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawLinesI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLinesI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawLinesI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPath(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_DrawPath = CBool(GdipDrawPath(dstGraphics, srcPen, srcPath) = GP_OK)
End Function

Public Function GDIPlus_DrawPolygonF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygon(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawPolygonF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPolygonI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygonI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawPolygonI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawRectF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_DrawRectF = CBool(GdipDrawRectangle(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawRectI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_DrawRectI = CBool(GdipDrawRectangleI(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_DrawEllipseF = CBool(GdipDrawEllipse(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_DrawEllipseI = CBool(GdipDrawEllipseI(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillClosedCurveF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillClosedCurveI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2I(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPath(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_FillPath = CBool(GdipFillPath(dstGraphics, srcBrush, srcPath) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_FillEllipseF = CBool(GdipFillEllipse(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_FillEllipseI = CBool(GdipFillEllipseI(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillPolygonF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygon(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPolygonI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygonI(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillRectF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_FillRectF = CBool(GdipFillRectangle(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_FillRectI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_FillRectI = CBool(GdipFillRectangleI(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

'WARNING!  If a graphics object has never specified a clipping region, the default region is infinite.
' For reasons unknown, GDI+ is finicky about returning such a region; it often reports "Object Busy" for no
' apparent reason.  I'm not sure of a good workaround.
Public Function GDIPlus_GraphicsGetClipRegion(ByVal srcGraphics As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetClip(srcGraphics, GDIPlus_GraphicsGetClipRegion)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsResetClipRegion(ByVal dstGraphics As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipResetClip(dstGraphics)
    GDIPlus_GraphicsResetClipRegion = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetClipRect(ByVal dstGraphics As Long, ByVal clipX As Single, ByVal clipY As Single, ByVal clipWidth As Single, ByVal clipHeight As Single, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetClipRect(dstGraphics, clipX, clipY, clipWidth, clipHeight, useCombineMode)
    GDIPlus_GraphicsSetClipRect = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetClipRegion(ByVal dstGraphics As Long, ByVal srcRegion As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetClipRegion(dstGraphics, srcRegion, useCombineMode)
    GDIPlus_GraphicsSetClipRegion = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetCompositingMode(ByVal dstGraphics As Long, Optional ByVal newCompositeMode As GP_CompositingMode = GP_CM_SourceOver) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetCompositingMode(dstGraphics, newCompositeMode)
    GDIPlus_GraphicsSetCompositingMode = CBool(tmpReturn = GP_OK)
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
            CopyMemory_Strict tmpLockMem, VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1
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
            isImageMetafile = CBool(imgType = GP_IT_Metafile)
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
        isImageMetafile = CBool(imgType = GP_IT_Metafile)
    Else
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
End Function

'This function only works on bitmaps (never metafiles!), and the source image *must* already be in 32-bpp format.
Public Function GDIPlus_ImageForcePremultipliedAlpha(ByVal hImage As Long, ByVal imgWidth As Long, ByVal imgHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCloneBitmapAreaI(0, 0, imgWidth, imgHeight, GP_PF_32bppPARGB, hImage, hImage)
    GDIPlus_ImageForcePremultipliedAlpha = CBool(tmpReturn = GP_OK)
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
        GDIPlus_ImageGetDimensions = CBool(tmpReturn = GP_OK)
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
                CopyMemory_Strict StrPtr(GDIPlus_ImageGetFileFormatGUID), imgStringPointer, strLength * 2
            End If
        Else
            InternalGDIPlusError "Failed to convert CLSID to string", "GDIPlus_ImageGetFileFormatGUID failed"
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

'Given a pd2D file format, return a matching GDI+ GUID format identifier (as a string; you'll need to manually
' convert this to a byte array, FYI!)
Private Function GetGUIDFromPd2dFileFormat(ByVal srcFileFormat As PD_2D_FileFormatImport) As String
    Select Case srcFileFormat
        Case P2_FFI_BMP
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_BMP
        Case P2_FFI_EMF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_EMF
        Case P2_FFI_WMF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_WMF
        Case P2_FFI_JPEG
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_JPEG
        Case P2_FFI_PNG
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_PNG
        Case P2_FFI_GIF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_GIF
        Case P2_FFI_TIFF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_TIFF
        Case P2_FFI_ICO
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_Icon
        Case Else
            GetGUIDFromPd2dFileFormat = vbNullString
    End Select
End Function

Public Function GDIPlus_ImageGetPixelFormat(ByVal hImage As Long) As GP_PixelFormat
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImagePixelFormat(hImage, GDIPlus_ImageGetPixelFormat)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'It's important to check the return value of this function; it will be FALSE if the image does not
' contain/provide the requested property.  Also note that all properties are returned as byte arrays.
' It is up to the caller to make sense of this return, presumably using the MSDN guide at
' https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx
Public Function GDIPlus_ImageGetProperty(ByVal hImage As Long, ByVal gpPropertyID As GP_PropertyTag, ByRef dstBuffer() As Byte) As Boolean
    
    Dim tmpReturn As GP_Result, propSize As Long
    tmpReturn = GdipGetPropertyItemSize(hImage, gpPropertyID, propSize)
    If (tmpReturn = GP_OK) Then
    
        If (propSize > 0) Then
            ReDim dstBuffer(0 To propSize - 1) As Byte
            tmpReturn = GdipGetPropertyItem(hImage, gpPropertyID, propSize, VarPtr(dstBuffer(0)))
            If (tmpReturn = GP_OK) Then
                GDIPlus_ImageGetProperty = True
            Else
                InternalGDIPlusError vbNullString, vbNullString, tmpReturn
                GDIPlus_ImageGetProperty = False
            End If
        Else
            GDIPlus_ImageGetProperty = False
        End If
        
    Else
        'NOTE: it's totally okay for an image to not have a given property.  This is not a meaningful error,
        ' so we do not report it.
        'InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        GDIPlus_ImageGetProperty = False
    End If
    
End Function

Public Function GDIPlus_ImageGetResolution(ByVal hImage As Long, ByRef dstHResolution As Single, ByRef dstVResolution As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImageHorizontalResolution(hImage, dstHResolution)
    If (tmpReturn = GP_OK) Then
        tmpReturn = GdipGetImageVerticalResolution(hImage, dstVResolution)
        GDIPlus_ImageGetResolution = CBool(tmpReturn = GP_OK)
    Else
        GDIPlus_ImageGetResolution = False
    End If
End Function

Public Function GDIPlus_ImageLockBits(ByVal hImage As Long, ByRef srcRect As RECTL, ByRef srcCopyData As GP_BitmapData, ByVal lockFlags As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapLockBits(hImage, srcRect, lockFlags, dstPixelFormat, srcCopyData)
    GDIPlus_ImageLockBits = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_ImageRotateFlip(ByVal hImage As Long, ByVal typeOfRotateFlip As GP_RotateFlip) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipImageRotateFlip(hImage, typeOfRotateFlip)
    GDIPlus_ImageRotateFlip = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Save a surface to a VB byte array.  The destination array *must* be dynamic, and does not need to be dimensionsed.
' (It will be auto-dimensioned correctly by thsi function.)
' As with saving to file, note that the only export property currently supported is JPEG quality; other properties
' are set automatically by GDI+.
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
        Dim fullEncoderParams() As Byte
        
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
                        CopyMemory_Strict VarPtr(dstArray(0)), lockedMem, hMemSize
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
        Dim fullEncoderParams() As Byte
        
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
    
    If (Len(srcMimetype) <> 0) Then
        
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
                    Dim i As Long, strLength As Long, tmpMimeType As String
                    For i = 0 To numOfEncoders - 1
                    
                        'Extract this codec
                        CopyMemory_Strict VarPtr(tmpCodec), VarPtr(encoderBuffer(0)) + LenB(tmpCodec) * i, LenB(tmpCodec)
                        
                        'Extract the codec's mimetype
                        strLength = lstrlenW(tmpCodec.IC_MimeType)
                        If (strLength <> 0) Then
                            tmpMimeType = String$(strLength, 0&)
                            CopyMemory_Strict StrPtr(tmpMimeType), tmpCodec.IC_MimeType, strLength * 2
                            
                            'If we find a match, copy the encoder GUID and exit
                            If (StrComp(srcMimetype, tmpMimeType, vbBinaryCompare) = 0) Then
                                GetEncoderGUIDForPd2dFormat = True
                                CopyMemory_Strict ptrToDstGuid, VarPtr(tmpCodec.IC_ClassID(0)), 16&
                                Exit For
                            End If
                        End If
                        
                    Next i
                
                End If
                
            End If
        End If
        
    End If

End Function

Public Function GDIPlus_ImageUnlockBits(ByVal hImage As Long, ByRef srcCopyData As GP_BitmapData) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapUnlockBits(hImage, srcCopyData)
    GDIPlus_ImageUnlockBits = CBool(tmpReturn = GP_OK)
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
    
    If (tmpReturn = GP_OK) Then
        GDIPlus_ImageUpgradeMetafile = True
    Else
        GDIPlus_ImageUpgradeMetafile = False
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
    
End Function

Public Function GDIPlus_MatrixCreate() As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCreateMatrix(GDIPlus_MatrixCreate)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixClone(ByVal srcMatrix As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCloneMatrix(srcMatrix, GDIPlus_MatrixClone)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixDelete(ByVal hMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDeleteMatrix(hMatrix)
    GDIPlus_MatrixDelete = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixInvert(ByVal hMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipInvertMatrix(hMatrix)
    GDIPlus_MatrixInvert = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixIsInvertible(ByVal hMatrix As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsMatrixInvertible(hMatrix, tmpResult)
    GDIPlus_MatrixIsInvertible = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixRotate(ByVal hMatrix As Long, ByVal rotateAngle As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipRotateMatrix(hMatrix, rotateAngle, operationOrder)
    GDIPlus_MatrixRotate = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixScale(ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipScaleMatrix(hMatrix, scaleX, scaleY, operationOrder)
    GDIPlus_MatrixScale = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixShear(ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipShearMatrix(hMatrix, shearX, shearY, operationOrder)
    GDIPlus_MatrixShear = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixTransformListOfPoints(ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTransformMatrixPoints(hMatrix, ptrToFirstPointF, numOfPoints)
    GDIPlus_MatrixTransformListOfPoints = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixTranslate(ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTranslateMatrix(hMatrix, offsetX, offsetY, operationOrder)
    GDIPlus_MatrixTranslate = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddArc(ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathArc(hPath, x, y, arcWidth, arcHeight, startAngle, sweepAngle)
    GDIPlus_PathAddArc = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddBezier(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathBezier(hPath, x1, y1, x2, y2, x3, y3, x4, y4)
    GDIPlus_PathAddBezier = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddClosedCurve(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathClosedCurve2(hPath, ptrToFloatArray, numOfPoints, curveTension)
    GDIPlus_PathAddClosedCurve = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddCurve(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathCurve2(hPath, ptrToFloatArray, numOfPoints, curveTension)
    GDIPlus_PathAddCurve = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddEllipse(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathEllipse(hPath, x1, y1, rectWidth, rectHeight)
    GDIPlus_PathAddEllipse = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddLine(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathLine(hPath, x1, y1, x2, y2)
    GDIPlus_PathAddLine = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddLines(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathLine2(hPath, ptrToFloatArray, numOfPoints)
    GDIPlus_PathAddLines = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddPath(ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathPath(hPath, pathToAdd, connectToPreviousPoint)
    GDIPlus_PathAddPath = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddPolygon(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathPolygon(hPath, ptrToFloatArray, numOfPoints)
    GDIPlus_PathAddPolygon = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddRectangle(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathRectangle(hPath, x1, y1, rectWidth, rectHeight)
    GDIPlus_PathAddRectangle = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathClone(ByVal srcPath As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipClonePath(srcPath, GDIPlus_PathClone)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathCloseFigure(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipClosePathFigure(hPath)
    GDIPlus_PathCloseFigure = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathCreate(Optional ByVal initFillRule As GP_FillMode = GP_FM_Alternate) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCreatePath(initFillRule, GDIPlus_PathCreate)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDelete(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDeletePath(hPath)
    GDIPlus_PathDelete = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDoesPointTouchOutlineF(ByVal hPath As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal hPen As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsOutlineVisiblePathPoint(hPath, srcX, srcY, hPen, 0&, tmpResult)
    GDIPlus_PathDoesPointTouchOutlineF = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDoesPointTouchOutlineL(ByVal hPath As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal hPen As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsOutlineVisiblePathPointI(hPath, srcX, srcY, hPen, 0&, tmpResult)
    GDIPlus_PathDoesPointTouchOutlineL = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetFillRule(ByVal hPath As Long) As GP_FillMode
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathFillMode(hPath, GDIPlus_PathGetFillRule)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetPathBoundsF(ByVal hPath As Long, Optional ByVal hTransform As Long = 0, Optional ByVal hPen As Long = 0) As RECTF
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathWorldBounds(hPath, GDIPlus_PathGetPathBoundsF, hTransform, hPen)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetPathBoundsL(ByVal hPath As Long, Optional ByVal hTransform As Long = 0, Optional ByVal hPen As Long = 0) As RECTL
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathWorldBoundsI(hPath, GDIPlus_PathGetPathBoundsL, hTransform, hPen)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathIsPointInsideF(ByVal hPath As Long, ByVal srcX As Single, ByVal srcY As Single) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsVisiblePathPoint(hPath, srcX, srcY, 0&, tmpResult)
    GDIPlus_PathIsPointInsideF = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathIsPointInsideL(ByVal hPath As Long, ByVal srcX As Long, ByVal srcY As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsVisiblePathPointI(hPath, srcX, srcY, 0&, tmpResult)
    GDIPlus_PathIsPointInsideL = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathReset(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipResetPath(hPath)
    GDIPlus_PathReset = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathSetFillRule(ByVal hPath As Long, ByVal newFillRule As GP_FillMode) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetPathFillMode(hPath, newFillRule)
    GDIPlus_PathSetFillRule = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathStartFigure(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipStartPathFigure(hPath)
    GDIPlus_PathStartFigure = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathTransform(ByVal hPath As Long, ByVal hTransformMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTransformPath(hPath, hTransformMatrix)
    GDIPlus_PathTransform = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathWiden(ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipWidenPath(hPath, hPen, hTransformMatrix, allowableError)
    GDIPlus_PathWiden = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathWindingModeOutline(ByVal hPath As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipWindingModeOutline(hPath, hTransformMatrix, allowableError)
    GDIPlus_PathWindingModeOutline = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionAddRectF(ByVal dstRegion As Long, ByRef srcRectF As RECTF, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRectF = CBool(GdipCombineRegionRect(dstRegion, srcRectF, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddRectL(ByVal dstRegion As Long, ByRef srcRectL As RECTL, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRectL = CBool(GdipCombineRegionRectI(dstRegion, srcRectL, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddRegion(ByVal dstRegion As Long, ByVal srcRegion As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRegion = CBool(GdipCombineRegionRegion(dstRegion, srcRegion, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddPath(ByVal dstRegion As Long, ByVal srcPath As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddPath = CBool(GdipCombineRegionPath(dstRegion, srcPath, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionClone(ByVal srcRegion As Long, ByRef dstRegion As Long) As Boolean
    Dim tmpReturn As Long
    tmpReturn = GdipCloneRegion(srcRegion, dstRegion)
    GDIPlus_RegionClone = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionGetClipRectF(ByVal srcRegion As Long) As RECTF
    Dim tmpReturn As Long
    tmpReturn = GdipGetRegionBounds(srcRegion, m_TransformGraphics, GDIPlus_RegionGetClipRectF)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionGetClipRectI(ByVal srcRegion As Long) As RECTL
    Dim tmpReturn As Long
    tmpReturn = GdipGetRegionBoundsI(srcRegion, m_TransformGraphics, GDIPlus_RegionGetClipRectI)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionIsInfinite(ByVal srcRegion As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsInfiniteRegion(srcRegion, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionIsInfinite = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionIsEmpty(ByVal srcRegion As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsEmptyRegion(srcRegion, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionIsEmpty = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionsAreEqual(ByVal srcRegion1 As Long, ByVal srcRegion2 As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsEqualRegion(srcRegion1, srcRegion2, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionsAreEqual = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionSetEmpty(ByVal dstRegion As Long) As Boolean
    GDIPlus_RegionSetEmpty = CBool(GdipSetEmpty(dstRegion) = GP_OK)
End Function

Public Function GDIPlus_RegionSetInfinite(ByVal dstRegion As Long) As Boolean
    GDIPlus_RegionSetInfinite = CBool(GdipSetInfinite(dstRegion) = GP_OK)
End Function


