Attribute VB_Name = "Plugin_CharLS"
'***************************************************************************
'CharLS (lossless JPEG) Library Interface
'Copyright 2021-2021 by Tanner Helland
'Created: 12/September/21
'Last updated: 12/September/21
'Last update: initial build
'
'Per its documentation (available at https://github.com/team-charls/charls), char-ls is...
'
' "...a C++ implementation of the JPEG-LS standard for lossless and near-lossless image compression
' and decompression. JPEG-LS is a low-complexity image compression standard that matches JPEG 2000
' compression ratios."
'
'CharLS is BSD-licensed and actively maintained.  Fortunately for PhotoDemon, they also provide
' a robust C interface and officially supported VS solutions, with excellent x86 support.
' PD's CharLS dll is built using those solutions (with minor modifications for Windows XP support).
'
'Also worth noting that "Lossless JPG" is kind of a strange beast; see Wikipedia for details:
' https://jpeg.org/jpegls/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A few APIs return unique values (e.g. charls_get_version_string returns *char) but most functions
' return an element of this enum:
Private Enum CharLS_Return
    cls_SUCCESS = 0
    cls_INVALID_ARGUMENT = 1
    cls_PARAMETER_VALUE_NOT_SUPPORTED = 2
    cls_DESTINATION_BUFFER_TOO_SMALL = 3
    cls_SOURCE_BUFFER_TOO_SMALL = 4
    cls_INVALID_ENCODED_DATA = 5
    cls_TOO_MUCH_ENCODED_DATA = 6
    cls_INVALID_OPERATION = 7
    cls_BIT_DEPTH_FOR_TRANSFORM_NOT_SUPPORTED = 8
    cls_COLOR_TRANSFORM_NOT_SUPPORTED = 9
    cls_ENCODING_NOT_SUPPORTED = 10
    cls_UNKNOWN_JPEG_MARKER_FOUND = 11
    cls_JPEG_MARKER_START_BYTE_NOT_FOUND = 12
    cls_NOT_ENOUGH_MEMORY = 13
    cls_UNEXPECTED_FAILURE = 14
    cls_START_OF_IMAGE_MARKER_NOT_FOUND = 15
    cls_UNEXPECTED_MARKER_FOUND = 16
    cls_INVALID_MARKER_SEGMENT_SIZE = 17
    cls_DUPLICATE_START_OF_IMAGE_MARKER = 18
    cls_DUPLICATE_START_OF_FRAME_MARKER = 19
    cls_DUPLICATE_COMPONENT_ID_IN_SOF_SEGMENT = 20
    cls_UNEXPECTED_END_OF_IMAGE_MARKER = 21
    cls_INVALID_JPEGLS_PRESET_PARAMETER_TYPE = 22
    cls_JPEGLS_PRESET_EXTENDED_PARAMETER_TYPE_NOT_SUPPORTED = 23
    cls_MISSING_END_OF_SPIFF_DIRECTORY = 24
    cls_INVALID_ARGUMENT_WIDTH = 100
    cls_INVALID_ARGUMENT_HEIGHT = 101
    cls_INVALID_ARGUMENT_COMPONENT_COUNT = 102
    cls_INVALID_ARGUMENT_BITS_PER_SAMPLE = 103
    cls_INVALID_ARGUMENT_INTERLEAVE_MODE = 104
    cls_INVALID_ARGUMENT_NEAR_LOSSLESS = 105
    cls_INVALID_ARGUMENT_JPEGLS_PC_PARAMETERS = 106
    cls_INVALID_ARGUMENT_SPIFF_ENTRY_SIZE = 110
    cls_INVALID_ARGUMENT_COLOR_TRANSFORMATION = 111
    cls_INVALID_PARAMETER_WIDTH = 200
    cls_INVALID_PARAMETER_HEIGHT = 201
    cls_INVALID_PARAMETER_COMPONENT_COUNT = 202
    cls_INVALID_PARAMETER_BITS_PER_SAMPLE = 203
    cls_INVALID_PARAMETER_INTERLEAVE_MODE = 204
    cls_INVALID_PARAMETER_NEAR_LOSSLESS = 205
    cls_INVALID_PARAMETER_JPEGLS_PC_PARAMETERS = 206
End Enum

'Defines the color space options that can be used in a SPIFF header v2, as defined in ISO/IEC 10918-3, F.2.1.1
Private Enum CharLS_ColorSpace
    CHARLS_SPIFF_COLOR_SPACE_BI_LEVEL_BLACK             'Bi-level image. Each image sample is one bit: 0 = white and 1 = black.
    CHARLS_SPIFF_COLOR_SPACE_YCBCR_ITU_BT_709_VIDEO     'The color space is based on recommendation ITU-R BT.709.
    CHARLS_SPIFF_COLOR_SPACE_NONE                       'Color space interpretation of the coded sample is none of the other options.
    CHARLS_SPIFF_COLOR_SPACE_YCBCR_ITU_BT_601_1_RGB     'The color space is based on recommendation ITU-R BT.601-1. (RGB).
    CHARLS_SPIFF_COLOR_SPACE_YCBCR_ITU_BT_601_1_VIDEO   'The color space is based on recommendation ITU-R BT.601-1. (video).
    CHARLS_SPIFF_COLOR_SPACE_GRAYSCALE                  'Grayscale – This is a single component sample with interpretation as grayscale value, 0 is minimum, 2bps -1 is maximum.
    CHARLS_SPIFF_COLOR_SPACE_PHOTO_YCC                  'This is the color encoding method used in the Photo CD™ system.
    CHARLS_SPIFF_COLOR_SPACE_RGB                        'The encoded data consists of samples of (uncalibrated) R, G and B.
    CHARLS_SPIFF_COLOR_SPACE_CMY                        'The encoded data consists of samples of Cyan, Magenta and Yellow samples.
    CHARLS_SPIFF_COLOR_SPACE_CMYK                       'The encoded data consists of samples of Cyan, Magenta, Yellow and Black samples.
    CHARLS_SPIFF_COLOR_SPACE_YCCK                       'Transformed CMYK type data (same as Adobe PostScript)
    CHARLS_SPIFF_COLOR_SPACE_CIE_LAB                    'The CIE 1976 (L* a* b*) color space.
    CHARLS_SPIFF_COLOR_SPACE_BI_LEVEL_WHITE             'Bi-level image. Each image sample is one bit: 1 = white and 0 = black.
End Enum

Private Enum CharLS_ColorTransformation
    
    'No color space transformation has been applied.
    CHARLS_COLOR_TRANSFORMATION_NONE

    'Defines the reversible lossless color transformation:
    ' G = G
    ' R = R - G
    ' B = B - G
    CHARLS_COLOR_TRANSFORMATION_HP1

    'Defines the reversible lossless color transformation:
    ' G = G
    ' B = B - (R + G) / 2
    ' R = R - G
    CHARLS_COLOR_TRANSFORMATION_HP2

    'Defines the reversible lossless color transformation of Y-Cb-Cr:
    ' R = R - G
    ' B = B - G
    ' G = G + (R + B) / 4
    CHARLS_COLOR_TRANSFORMATION_HP3
    
End Enum

Private Enum CharLS_InterleaveMode
    CHARLS_INTERLEAVE_MODE_NONE 'The data is encoded and stored as component for component: RRRGGGBBB.
    CHARLS_INTERLEAVE_MODE_LINE 'The interleave mode is by line. A full line of each component is encoded before moving to the next line.
    CHARLS_INTERLEAVE_MODE_SAMPLE   'The data is encoded and stored by sample. For RGB color images this is the format like RGBRGBRGB.
End Enum

'Defines the resolution units for the VRES and HRES parameters, as defined in ISO/IEC 10918-3, F.2.1
Private Enum CharLS_ResolutionUnits
    'VRES and HRES are to be interpreted as aspect ratio.
    ' If vertical or horizontal resolutions are not known, use this option and set VRES and HRES
    ' both to 1 to indicate that pixels in the image should be assumed to be square.
    CHARLS_SPIFF_RESOLUTION_UNITS_ASPECT_RATIO

    'Units of dots/samples per inch
    CHARLS_SPIFF_RESOLUTION_UNITS_DOTS_PER_INCH

    'Units of dots/samples per centimeter
    CHARLS_SPIFF_RESOLUTION_UNITS_DOTS_PER_CENTIMETER
End Enum

Private Type CharLSFrameInfo
    frame_width As Long     'Width of the image, range [1, 65535]
    frame_height As Long    'Height of the image, range [1, 65535]
    bits_per_sample As Long 'Number of bits per sample, range [2, 16]
    component_count As Long 'Number of components contained in the frame, range [1, 255]
End Type

Private Type CharLSCodingParameters

    'Maximum possible value for any image sample in a scan.
    ' This must be greater than or equal to the actual maximum value for the components in a scan.
    maximum_sample_value As Long
    
    'First quantization threshold value for the local gradients.
    threshold1 As Long
    
    'Second quantization threshold value for the local gradients.
    threshold2 As Long
    
    'Third quantization threshold value for the local gradients.
    threshold3 As Long
    
    'Value at which the counters A, B, and N are halved.
    reset_value As Long
    
End Type

Private Type CharLSSpiffHeader
    profile_id As Long               '// P: Application profile, type I.8
    component_count As Long          '// NC: Number of color components, range [1, 255], type I.8
    height As Long                   '// HEIGHT: Number of lines in image, range [1, 4294967295], type I.32
    width As Long                    '// WIDTH: Number of samples per line, range [1, 4294967295], type I.32
    color_space As CharLS_ColorSpace '// S: Color space used by image data, type is I.8
    bits_per_sample As Long          '// BPS: Number of bits per sample, range (1, 2, 4, 8, 12, 16), type is I.8
    compression_type As Long         '// C: Type of data compression used, type is I.8
    resolution_units As CharLS_ResolutionUnits  '// R: Type of resolution units, type is I.8
    vertical_resolution As Long      '// VRES: Vertical resolution, range [1, 4294967295], type can be F or I.32
    horizontal_resolution As Long    '// HRES: Horizontal resolution, range [1, 4294967295], type can be F or I.32
End Type

'CharLS provides a few shorthand encode/decode functions, but these are deprecated and marked
' for removal (so I haven't even bothered declaring them here).
'JpegLsDecode
'JpegLsDecodeRect
'JpegLsEncode
'JpegLsReadHeader

'Modern, non-deprecated APIs use decorated names.  (Aliasing them here makes it much easier to sync
' against upstream vs rebasing a modified .def on each new release.)
Private Declare Function charls_get_error_message Lib "charls-2-x86" Alias "_charls_get_error_message@4" (ByVal charLSErrorNumber As CharLS_Return) As Long
Private Declare Function charls_get_jpegls_category Lib "charls-2-x86" Alias "_charls_get_jpegls_category@0" () As Long

'void charls_get_version_number(OUT_OPT_ int32_t* major, OUT_OPT_ int32_t* minor, OUT_OPT_ int32_t* patch)
Private Declare Sub charls_get_version_number Lib "charls-2-x86" Alias "_charls_get_version_number@12" (ByRef vMajor As Long, ByRef vMinor As Long, ByRef vPatch As Long)
'const char* charls_get_version_string()
Private Declare Function charls_get_version_string Lib "charls-2-x86" Alias "_charls_get_version_string@0" () As Long

Private Declare Function charls_jpegls_decoder_create Lib "charls-2-x86" Alias "_charls_jpegls_decoder_create@0" () As Long
Private Declare Sub charls_jpegls_decoder_destroy Lib "charls-2-x86" Alias "_charls_jpegls_decoder_destroy@4" (ByVal srcDecoder As Long)

Private Declare Function charls_jpegls_decoder_decode_to_buffer Lib "charls-2-x86" Alias "_charls_jpegls_decoder_decode_to_buffer@16" (ByVal srcDecoder As Long, ByVal ptrToDstBytes As Long, ByVal sizeOfDstArray As Long, Optional ByVal dstStride As Long = 0) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_color_transformation Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_color_transformation@8" (ByVal srcDecoder As Long, ByRef dstColorTransform As CharLS_ColorTransformation) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_destination_size Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_destination_size@12" (ByVal srcDecoder As Long, ByRef dstStride As Long, ByRef dstSizeBytes As Long) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_frame_info Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_frame_info@8" (ByVal srcDecoder As Long, ByRef dstFrameInfo As CharLSFrameInfo) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_interleave_mode Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_interleave_mode@8" (ByVal srcDecoder As Long, ByRef dstInterleaveMode As CharLS_InterleaveMode) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_near_lossless Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_near_lossless@12" (ByVal srcDecoder As Long, ByVal idxComponent As Long, ByRef dstNearLosslessValue As Long) As CharLS_Return
Private Declare Function charls_jpegls_decoder_get_preset_coding_parameters Lib "charls-2-x86" Alias "_charls_jpegls_decoder_get_preset_coding_parameters@12" (ByVal srcDecoder As Long, ByVal intReserved As Long, ByVal dstCodingParameters As CharLSCodingParameters) As CharLS_Return
Private Declare Function charls_jpegls_decoder_read_header Lib "charls-2-x86" Alias "_charls_jpegls_decoder_read_header@4" (ByVal srcDecoder As Long) As CharLS_Return
Private Declare Function charls_jpegls_decoder_read_spiff_header Lib "charls-2-x86" Alias "_charls_jpegls_decoder_read_spiff_header@12" (ByVal srcDecoder As Long, ByVal ptrDstSpiffHeader As Long, ByRef dstHeaderFound As Long) As CharLS_Return
Private Declare Function charls_jpegls_decoder_set_source_buffer Lib "charls-2-x86" Alias "_charls_jpegls_decoder_set_source_buffer@12" (ByVal srcDecoder As Long, ByVal ptrToBytes As Long, ByVal srcNumBytes As Long) As CharLS_Return

'I have *NOT* added params for these declares, because PD currently only supports .jls reading (by design).
' If you need support for .jls writing, please file an issue on GitHub.
'Private Declare Function charls_jpegls_encoder_create Lib "charls-2-x86" Alias "_charls_jpegls_encoder_create@0" () As Long
'Private Declare Function charls_jpegls_encoder_destroy Lib "charls-2-x86" Alias "_charls_jpegls_encoder_destroy@4" () As Long
'Private Declare Function charls_jpegls_encoder_encode_from_buffer Lib "charls-2-x86" Alias "_charls_jpegls_encoder_encode_from_buffer@16" () As Long
'Private Declare Function charls_jpegls_encoder_get_bytes_written Lib "charls-2-x86" Alias "_charls_jpegls_encoder_get_bytes_written@8" () As Long
'Private Declare Function charls_jpegls_encoder_get_estimated_destination_size Lib "charls-2-x86" Alias "_charls_jpegls_encoder_get_estimated_destination_size@8" () As Long
'Private Declare Function charls_jpegls_encoder_rewind Lib "charls-2-x86" Alias "_charls_jpegls_encoder_rewind@4" () As Long
'Private Declare Function charls_jpegls_encoder_set_color_transformation Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_color_transformation@8" () As Long
'Private Declare Function charls_jpegls_encoder_set_destination_buffer2 Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_destination_buffer@12" () As Long
'Private Declare Function charls_jpegls_encoder_set_frame_info Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_frame_info@8" () As Long
'Private Declare Function charls_jpegls_encoder_set_interleave_mode Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_interleave_mode@8" () As Long
'Private Declare Function charls_jpegls_encoder_set_near_lossless Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_near_lossless@8" () As Long
'Private Declare Function charls_jpegls_encoder_set_preset_coding_parameters Lib "charls-2-x86" Alias "_charls_jpegls_encoder_set_preset_coding_parameters@8" () As Long
'Private Declare Function charls_jpegls_encoder_write_spiff_entry Lib "charls-2-x86" Alias "_charls_jpegls_encoder_write_spiff_entry@16" () As Long
'Private Declare Function charls_jpegls_encoder_write_spiff_header Lib "charls-2-x86" Alias "_charls_jpegls_encoder_write_spiff_header@8" () As Long
'Private Declare Function charls_jpegls_encoder_write_standard_spiff_header Lib "charls-2-x86" Alias "_charls_jpegls_encoder_write_standard_spiff_header@20" () As Long

'Library handle will be non-zero if CharLS is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'Forcibly disable CharLS interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String
    Dim vMajor As Long, vMinor As Long, vPatch As Long
    charls_get_version_number vMajor, vMinor, vPatch
    GetVersion = vMajor & "." & vMinor & "." & vPatch
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "charls-2-x86.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load CharLS.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Public Function IsCharLSEnabled() As Boolean
    IsCharLSEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub
