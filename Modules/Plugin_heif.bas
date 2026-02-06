Attribute VB_Name = "Plugin_Heif"
'***************************************************************************
'libheif Library Interface
'Copyright 2024-2026 by Tanner Helland
'Created: 16/July/24
'Last updated: 30/May/25
'Last update: fix filetype check potentially breaking on tiny (<128b) image files
'
'Per its documentation (available at https://github.com/strukturag/libheif), libheif is...
'
'"...an ISO/IEC 23008-12:2017 HEIF and AVIF (AV1 Image File Format) file format
' decoder and encoder... HEIF and AVIF are new image file formats employing HEVC (H.265)
' or AV1 image coding, respectively, for the best compression ratios currently possible."
'
'libheif is LGPL-licensed and actively maintained.  PhotoDemon does not use its potential
' AVIF support due to x86 compatibility issues (AVIF support is 64-bit focused and x86 builds
' are not currently feasible for me to self-maintain, so I only compile with HEIF enabled).
'
'Note that all features in this module rely on the libheif binaries that ship with PhotoDemon.
' These features will not work if libheif cannot be located.  Per standard LGPL terms, you can
' supply your own libheif copies in place of PD's default ones, but libheif and all supporting
' libraries obviously need to be built as x86 libraries for this to work (not x64).
'
'Note also that there are quite a few encoding parameters supported by libheif.  Here are the
' encoding parameters reported by heif-enc (as listed via "./heif-enc.exe -P")
' Parameters for encoder `x265 HEVC encoder (3.5+39-931178347)`:
'  quality, Default = 50, [0;100]
'  lossless, Default = False
'  preset, default=slow, { ultrafast,superfast,veryfast,faster,fast,medium,slow,slower,veryslow,placebo }
'  tune, default=ssim, { psnr,ssim,grain,fastdecode }
'  tu-intra-depth, Default = 2, [1;4]
'  complexity, [0;100]
'  chroma, default=420, { 420,422,444 }
'
'PD could expose any of these to the user, but there are localization and usability
' considerations for these that I haven't fully considered (yet).  Exposing these remains TBD.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Log extra heif debug info; recommend DISABLING in production builds
Private Const HEIF_DEBUG_VERBOSE As Boolean = False

'libheif is built using vcpkg, which limits my control over calling convention.  (I don't want to
' manually build the libraries via CMake - they're complex!)  DispCallFunc is used to work around
' VB6 stdcall limitations.
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum Libheif_ProcAddress
    heif_get_version_number
    heif_check_filetype
    heif_context_alloc
    heif_context_free
    heif_context_get_image_handle
    heif_context_get_list_of_top_level_image_IDs
    heif_context_get_number_of_top_level_images
    heif_context_get_primary_image_ID
    heif_context_has_sequence
    heif_context_is_top_level_image_ID
    heif_encoder_list_parameters
    heif_encoder_release
    heif_image_get_plane
    heif_image_get_plane_readonly
    heif_image_handle_get_width
    heif_image_handle_get_height
    heif_image_handle_has_alpha_channel
    heif_image_handle_is_premultiplied_alpha
    heif_image_handle_get_color_profile_type
    heif_image_handle_get_raw_color_profile_size
    heif_image_handle_release
    heif_image_release
    heif_image_set_premultiplied_alpha
    heif_get_file_mime_type
    heif_read_main_brand
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Library handle will be non-zero if each library is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_hLibHeif As Long, m_hLibde265 As Long, m_hLibx265 As Long, m_LibAvailable As Boolean

Private Enum PD_HeifErrorCode
    heif_error_Ok = 0                   '// Everything ok, no error occurred.
    heif_error_Input_does_not_exist = 1 '// Input file does not exist.
    heif_error_Invalid_input = 2        '// Error in input file. Corrupted or invalid content.
    heif_error_Unsupported_filetype = 3 '// Input file type is not supported.
    heif_error_Unsupported_feature = 4  '// Image requires an unsupported decoder feature.
    heif_error_Usage_error = 5          '// Library API has been used in an invalid way.
    heif_error_Memory_allocation_error = 6  '// Could not allocate enough memory.
    heif_error_Decoder_plugin_error = 7 '// The decoder plugin generated an error
    heif_error_Encoder_plugin_error = 8 '// The encoder plugin generated an error
    heif_error_Encoding_error = 9       '// Error during encoding or when writing to the output
    heif_error_Color_profile_does_not_exist = 10    '// Application has asked for a color profile type that does not exist
    heif_error_Plugin_loading_error = 11    '// Error loading a dynamic plugin
End Enum

Private Enum PD_HeifSuberrorCode

    '// no further information available
    heif_suberror_Unspecified = 0
    
    '// --- Invalid_input ---
    
    '// End of data reached unexpectedly.
    heif_suberror_End_of_data = 100
    '// Size of box (defined in header) is wrong
    heif_suberror_Invalid_box_size = 101
    '// Mandatory 'ftyp' box is missing
    heif_suberror_No_ftyp_box = 102
    heif_suberror_No_idat_box = 103
    heif_suberror_No_meta_box = 104
    heif_suberror_No_hdlr_box = 105
    heif_suberror_No_hvcC_box = 106
    heif_suberror_No_pitm_box = 107
    heif_suberror_No_ipco_box = 108
    heif_suberror_No_ipma_box = 109
    heif_suberror_No_iloc_box = 110
    heif_suberror_No_iinf_box = 111
    heif_suberror_No_iprp_box = 112
    heif_suberror_No_iref_box = 113
    heif_suberror_No_pict_handler = 114
    '// An item property referenced in the 'ipma' box is not existing in the 'ipco' container.
    heif_suberror_Ipma_box_references_nonexisting_property = 115
    '// No properties have been assigned to an item.
    heif_suberror_No_properties_assigned_to_item = 116
    '// Image has no (compressed) data
    heif_suberror_No_item_data = 117
    '// Invalid specification of image grid (tiled image)
    heif_suberror_Invalid_grid_data = 118
    '// Tile-images in a grid image are missing
    heif_suberror_Missing_grid_images = 119
    heif_suberror_Invalid_clean_aperture = 120
    '// Invalid specification of overlay image
    heif_suberror_Invalid_overlay_data = 121
    '// Overlay image completely outside of visible canvas area
    heif_suberror_Overlay_image_outside_of_canvas = 122
    heif_suberror_Auxiliary_image_type_unspecified = 123
    heif_suberror_No_or_invalid_primary_item = 124
    heif_suberror_No_infe_box = 125
    heif_suberror_Unknown_color_profile_type = 126
    heif_suberror_Wrong_tile_image_chroma_format = 127
    heif_suberror_Invalid_fractional_number = 128
    heif_suberror_Invalid_image_size = 129
    heif_suberror_Invalid_pixi_box = 130
    heif_suberror_No_av1C_box = 131
    heif_suberror_Wrong_tile_image_pixel_depth = 132
    heif_suberror_Unknown_NCLX_color_primaries = 133
    heif_suberror_Unknown_NCLX_transfer_characteristics = 134
    heif_suberror_Unknown_NCLX_matrix_coefficients = 135
    '// Invalid specification of region item
    heif_suberror_Invalid_region_data = 136
    '// Image has no ispe property
    heif_suberror_No_ispe_property = 137
    heif_suberror_Camera_intrinsic_matrix_undefined = 138
    heif_suberror_Camera_extrinsic_matrix_undefined = 139
    '// Invalid JPEG 2000 codestream - usually a missing marker
    heif_suberror_Invalid_J2K_codestream = 140
    heif_suberror_No_vvcC_box = 141
    '// icbr is only needed in some situations, this error is for those cases
    heif_suberror_No_icbr_box = 142
    '// Decompressing generic compression or header compression data failed (e.g. bitstream corruption)
    heif_suberror_Decompression_invalid_data = 150
    
    '// --- Memory_allocation_error ---
    
    '// A security limit preventing unreasonable memory allocations was exceeded by the input file.
    '// Please check whether the file is valid. If it is, contact us so that we could increase the
    '// security limits further.
    heif_suberror_Security_limit_exceeded = 1000
    '// There was an error from the underlying compression / decompression library.
    '// One possibility is lack of resources (e.g. memory).
    heif_suberror_Compression_initialisation_error = 1001
    
    '// --- Usage_error ---
    
    '// An item ID was used that is not present in the file.
    heif_suberror_Nonexisting_item_referenced = 2000    '// also used for Invalid_input
    '// An API argument was given a NULL pointer, which is not allowed for that function.
    heif_suberror_Null_pointer_argument = 2001
    '// Image channel referenced that does not exist in the image
    heif_suberror_Nonexisting_image_channel_referenced = 2002
    '// The version of the passed plugin is not supported.
    heif_suberror_Unsupported_plugin_version = 2003
    '// The version of the passed writer is not supported.
    heif_suberror_Unsupported_writer_version = 2004
    '// The given (encoder) parameter name does not exist.
    heif_suberror_Unsupported_parameter = 2005
    '// The value for the given parameter is not in the valid range.
    heif_suberror_Invalid_parameter_value = 2006
    '// Error in property specification
    heif_suberror_Invalid_property = 2007
    '// Image reference cycle found in iref
    heif_suberror_Item_reference_cycle = 2008
    
    '// --- Unsupported_feature ---
    
    '// Image was coded with an unsupported compression method.
    heif_suberror_Unsupported_codec = 3000
    '// Image is specified in an unknown way, e.g. as tiled grid image (which is supported)
    heif_suberror_Unsupported_image_type = 3001
    heif_suberror_Unsupported_data_version = 3002
    '// The conversion of the source image to the requested chroma / colorspace is not supported.
    heif_suberror_Unsupported_color_conversion = 3003
    heif_suberror_Unsupported_item_construction_method = 3004
    heif_suberror_Unsupported_header_compression_method = 3005
    '// Generically compressed data used an unsupported compression method
    heif_suberror_Unsupported_generic_compression_method = 3006
    
    '// --- Encoder_plugin_error ---
    heif_suberror_Unsupported_bit_depth = 4000
    
    '// --- Encoding_error ---
    heif_suberror_Cannot_write_output_data = 5000
    heif_suberror_Encoder_initialization = 5001
    heif_suberror_Encoder_encoding = 5002
    heif_suberror_Encoder_cleanup = 5003
    heif_suberror_Too_many_regions = 5004
    
    '// --- Plugin loading error ---
    heif_suberror_Plugin_loading_error = 6000   '// a specific plugin file cannot be loaded
    heif_suberror_Plugin_is_not_loaded = 6001   '// trying to remove a plugin that is not loaded
    heif_suberror_Cannot_read_plugin_directory = 6002   '// error while scanning the directory for plugins
    heif_suberror_No_matching_decoder_installed = 6003  '// no decoder found for that compression format

End Enum

Private Enum heif_chroma
    heif_chroma_undefined = 99
    heif_chroma_monochrome = 0
    heif_chroma_420 = 1
    heif_chroma_422 = 2
    heif_chroma_444 = 3
    heif_chroma_interleaved_RGB = 10
    heif_chroma_interleaved_RGBA = 11
    heif_chroma_interleaved_RRGGBB_BE = 12    '// HDR, big endian.
    heif_chroma_interleaved_RRGGBBAA_BE = 13  '// HDR, big endian.
    heif_chroma_interleaved_RRGGBB_LE = 14    '// HDR, little endian.
    heif_chroma_interleaved_RRGGBBAA_LE = 15  '// HDR, little endian.
End Enum

Private Enum heif_colorspace
    heif_colorspace_undefined = 99
    
    '// heif_colorspace_YCbCr should be used with one of these heif_chroma values:
    '// * heif_chroma_444
    '// * heif_chroma_422
    '// * heif_chroma_420
    heif_colorspace_YCbCr = 0
    
    '// heif_colorspace_RGB should be used with one of these heif_chroma values:
    '// * heif_chroma_444 (for planar RGB)
    '// * heif_chroma_interleaved_RGB
    '// * heif_chroma_interleaved_RGBA
    '// * heif_chroma_interleaved_RRGGBB_BE
    '// * heif_chroma_interleaved_RRGGBBAA_BE
    '// * heif_chroma_interleaved_RRGGBB_LE
    '// * heif_chroma_interleaved_RRGGBBAA_LE
    heif_colorspace_RGB = 1
    
    '// heif_colorspace_monochrome should only be used with heif_chroma = heif_chroma_monochrome
    heif_colorspace_monochrome = 2
End Enum

Private Enum heif_channel
    heif_channel_Y = 0
    heif_channel_Cb = 1
    heif_channel_Cr = 2
    heif_channel_R = 3
    heif_channel_G = 4
    heif_channel_B = 5
    heif_channel_Alpha = 6
    heif_channel_interleaved = 10
End Enum

'libheif known compression formats.
Private Enum heif_compression_format
    'Unspecified / undefined compression format.
    ' This is used to mean "no match" or "any decoder" for some parts of the
    ' API. It does not indicate a specific compression format.
    heif_compression_undefined = 0
    
    'HEVC compression, used for HEIC images.
    ' This is equivalent to H.265.
    heif_compression_HEVC = 1
    
    'AVC compression. (Currently unused in libheif.)
    ' The compression is defined in ISO/IEC 14496-10. This is equivalent to H.264.
    ' The encapsulation is defined in ISO/IEC 23008-12:2022 Annex E.
    heif_compression_AVC = 2
    
    'JPEG compression.
    ' The compression format is defined in ISO/IEC 10918-1. The encapsulation
    ' of JPEG is specified in ISO/IEC 23008-12:2022 Annex H.
    heif_compression_JPEG = 3
    
    'AV1 compression, used for AVIF images.
    ' The compression format is provided at https://aomediacodec.github.io/av1-spec/
    ' The encapsulation is defined in https://aomediacodec.github.io/av1-avif/
    heif_compression_AV1 = 4
    
    'VVC compression. (Currently unused in libheif.)
    ' The compression format is defined in ISO/IEC 23090-3. This is equivalent to H.266.
    ' The encapsulation is defined in ISO/IEC 23008-12:2022 Annex L.
    heif_compression_VVC = 5
    
    'EVC compression. (Currently unused in libheif.)
    ' The compression format is defined in ISO/IEC 23094-1. This is equivalent to H.266.
    ' The encapsulation is defined in ISO/IEC 23008-12:2022 Annex M.
    heif_compression_EVC = 6
    
    'JPEG 2000 compression.
    ' The encapsulation of JPEG 2000 is specified in ISO/IEC 15444-16:2021.
    ' The core encoding is defined in ISO/IEC 15444-1, or ITU-T T.800.
    heif_compression_JPEG2000 = 7
    
    'Uncompressed encoding.
    ' This is defined in ISO/IEC 23001-17:2023 (Final Draft International Standard).
    heif_compression_uncompressed = 8
    
    'Mask image encoding.
    ' See ISO/IEC 23008-12:2022 Section 6.10.2
    heif_compression_mask = 9
    
    'High Throughput JPEG 2000 (HT-J2K) compression.
    ' The encapsulation of HT-J2K is specified in ISO/IEC 15444-16:2021.
    ' The core encoding is defined in ISO/IEC 15444-15, or ITU-T T.814.
    heif_compression_HTJ2K = 10
End Enum

Private Type heif_error
    heif_error_code As PD_HeifErrorCode         '// main error category
    heif_suberror_code As PD_HeifSuberrorCode   '// more detailed error code
    pCharMessage As Long    '// textual error message (is always defined, you do not have to check for NULL)
End Type

Private Enum heif_filetype_result
    heif_filetype_no
    heif_filetype_yes_supported     '// it is heif and can be read by libheif
    heif_filetype_yes_unsupported   '// it is heif, but cannot be read by libheif
    heif_filetype_maybe             'not sure whether it is an heif, try detection with more input data
End Enum

'Some libheif functions return a custom 12-byte type.  This cannot be easily handled via DispCallFunc,
' so I've written an interop DLL that marshals the cdecl return for me.
Private Declare Function PD_heif_init Lib "PDHelper_win32" () As heif_error
Private Declare Sub PD_heif_deinit Lib "PDHelper_win32" ()
Private Declare Function PD_heif_context_read_from_file Lib "PDHelper_win32" (ByVal heif_context As Long, ByVal pCharFilename As Long, ByVal pHeifReadingOptions As Long) As heif_error
Private Declare Function PD_heif_context_get_image_handle Lib "PDHelper_win32" (ByVal ptr_heif_context As Long, ByVal in_heif_item_id As Long, ByRef ptr_heif_image_handle As Long) As heif_error
Private Declare Function PD_heif_context_get_primary_image_handle Lib "PDHelper_win32" (ByVal ptr_heif_context As Long, ByRef ptr_heif_image_handle As Long) As heif_error
Private Declare Function PD_heif_context_get_primary_image_ID Lib "PDHelper_win32" (ByVal in_heif_context As Long, ByRef dst_heif_item_id As Long) As heif_error
Private Declare Function PD_heif_decode_image Lib "PDHelper_win32" (ByVal in_heif_image_handle As Long, ByRef pp_out_heif_image As Long, ByVal in_heif_colorspace As heif_colorspace, ByVal in_heif_chroma As heif_chroma, ByVal p_heif_decoding_options As Long) As heif_error
Private Declare Function PD_heif_image_handle_get_preferred_decoding_colorspace Lib "PDHelper_win32" (ByVal in_heif_image_handle As Long, ByRef out_heif_colorspace As Long, ByRef out_heif_chroma As Long) As heif_error
Private Declare Function PD_heif_image_handle_get_raw_color_profile Lib "PDHelper_win32" (ByVal in_heif_image_handle As Long, ByVal ptr_out_data As Long) As heif_error
Private Declare Function PD_heif_context_get_encoder_for_format Lib "PDHelper_win32" (ByVal p_heif_context As Long, ByVal heif_compression_format As heif_compression_format, ByVal pp_heif_encoder As Long) As heif_error
Private Declare Function PD_heif_encoder_set_lossy_quality Lib "PDHelper_win32" (ByVal p_heif_encoder As Long, ByVal enc_quality As Long) As heif_error
Private Declare Function PD_heif_encoder_set_lossless Lib "PDHelper_win32" (ByVal p_heif_encoder As Long, ByVal enc_enable As Long) As heif_error
Private Declare Function PD_heif_image_create Lib "PDHelper_win32" (ByVal image_width As Long, ByVal image_height As Long, ByVal in_heif_colorspace As heif_colorspace, ByVal in_heif_chroma As heif_chroma, ByVal pp_out_heif_image As Long) As heif_error
Private Declare Function PD_heif_image_add_plane Lib "PDHelper_win32" (ByVal p_heif_image As Long, ByVal in_heif_channel As heif_channel, ByVal in_width As Long, ByVal in_height As Long, ByVal in_bit_depth As Long) As heif_error
Private Declare Function PD_heif_context_encode_image Lib "PDHelper_win32" (ByVal p_heif_context As Long, ByVal p_heif_image As Long, ByVal p_heif_encoder As Long, ByVal p_heif_encoding_options As Long, ByRef pp_out_heif_image_handle As Long) As heif_error
Private Declare Function PD_heif_context_set_primary_image Lib "PDHelper_win32" (ByVal p_heif_context As Long, ByVal p_heif_image_handle As Long) As heif_error
Private Declare Function PD_heif_context_write_to_file Lib "PDHelper_win32" (ByVal p_heif_context As Long, ByVal p_char_filename As Long) As heif_error

'As a workaround for bugs in libheif v1.7.6 and later (see the PreviewHEIF function for details),
' a temp file is used during previews.
Private m_tmpFile As String, m_utfFilename() As Byte, m_utfLen As Long
    
'Forcibly disable plugin interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetHandle_LibHeif() As Long
    GetHandle_LibHeif = m_hLibHeif
End Function

Public Function GetHandle_Libde265() As Long
    GetHandle_Libde265 = m_hLibde265
End Function

Public Function GetHandle_Libx265() As Long
    GetHandle_Libx265 = m_hLibx265
End Function

Public Function GetVersion() As String
    
    If (m_hLibHeif = 0) Or (Not m_LibAvailable) Then Exit Function
        
    'Byte version numbers get packed into a long
    Dim versionAsInt(0 To 3) As Byte
    
    Dim tmpLong As Long
    tmpLong = CallCDeclW(heif_get_version_number, vbLong)
    PutMem4 VarPtr(versionAsInt(0)), tmpLong
    
    'Want to ensure we retrieved the correct values?  Use this:
    GetVersion = versionAsInt(3) & "." & versionAsInt(2) & "." & versionAsInt(1) & "." & versionAsInt(0)
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Because of a current reliance on vcpkg for compilation, libheif only works on Win 7+
    If (Not OS.IsWin7OrLater) Then
        PDDebug.LogAction "skipping libheif initialization (OS not supported)"
        Exit Function
    End If
    
    'Initialize all required libraries
    Dim strLibPath As String
    
    'Load dependencies first
    strLibPath = pathToDLLFolder & "libde265.dll"
    m_hLibde265 = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libx265.dll"
    m_hLibx265 = VBHacks.LoadLib(strLibPath)
    
    'The main library can now resolve dependencies correctly...
    strLibPath = pathToDLLFolder & "libheif.dll"
    m_hLibHeif = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_hLibHeif <> 0) And (m_hLibde265 <> 0) And (m_hLibx265 <> 0)
    InitializeEngine = m_LibAvailable
    
    'If we initialized the library successfully, preload some proc addresses
    If InitializeEngine Then
    
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(heif_get_version_number) = GetProcAddress(m_hLibHeif, "heif_get_version_number")
        m_ProcAddresses(heif_check_filetype) = GetProcAddress(m_hLibHeif, "heif_check_filetype")
        m_ProcAddresses(heif_context_alloc) = GetProcAddress(m_hLibHeif, "heif_context_alloc")
        m_ProcAddresses(heif_context_free) = GetProcAddress(m_hLibHeif, "heif_context_free")
        m_ProcAddresses(heif_context_get_image_handle) = GetProcAddress(m_hLibHeif, "heif_context_get_image_handle")
        m_ProcAddresses(heif_context_get_list_of_top_level_image_IDs) = GetProcAddress(m_hLibHeif, "heif_context_get_list_of_top_level_image_IDs")
        m_ProcAddresses(heif_context_get_number_of_top_level_images) = GetProcAddress(m_hLibHeif, "heif_context_get_number_of_top_level_images")
        m_ProcAddresses(heif_context_get_primary_image_ID) = GetProcAddress(m_hLibHeif, "heif_context_get_primary_image_ID")
        m_ProcAddresses(heif_context_has_sequence) = GetProcAddress(m_hLibHeif, "heif_context_has_sequence")
        m_ProcAddresses(heif_context_is_top_level_image_ID) = GetProcAddress(m_hLibHeif, "heif_context_is_top_level_image_ID")
        m_ProcAddresses(heif_encoder_list_parameters) = GetProcAddress(m_hLibHeif, "heif_encoder_list_parameters")
        m_ProcAddresses(heif_encoder_release) = GetProcAddress(m_hLibHeif, "heif_encoder_release")
        m_ProcAddresses(heif_image_get_plane) = GetProcAddress(m_hLibHeif, "heif_image_get_plane")
        m_ProcAddresses(heif_image_get_plane_readonly) = GetProcAddress(m_hLibHeif, "heif_image_get_plane_readonly")
        m_ProcAddresses(heif_image_handle_get_width) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_width")
        m_ProcAddresses(heif_image_handle_get_height) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_height")
        m_ProcAddresses(heif_image_handle_has_alpha_channel) = GetProcAddress(m_hLibHeif, "heif_image_handle_has_alpha_channel")
        m_ProcAddresses(heif_image_handle_is_premultiplied_alpha) = GetProcAddress(m_hLibHeif, "heif_image_handle_is_premultiplied_alpha")
        m_ProcAddresses(heif_image_handle_get_color_profile_type) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_color_profile_type")
        m_ProcAddresses(heif_image_handle_get_raw_color_profile_size) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_raw_color_profile_size")
        m_ProcAddresses(heif_image_handle_release) = GetProcAddress(m_hLibHeif, "heif_image_handle_release")
        m_ProcAddresses(heif_image_release) = GetProcAddress(m_hLibHeif, "heif_image_release")
        m_ProcAddresses(heif_image_set_premultiplied_alpha) = GetProcAddress(m_hLibHeif, "heif_image_set_premultiplied_alpha")
        m_ProcAddresses(heif_get_file_mime_type) = GetProcAddress(m_hLibHeif, "heif_get_file_mime_type")
        m_ProcAddresses(heif_read_main_brand) = GetProcAddress(m_hLibHeif, "heif_read_main_brand")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
        'HEIF support also requires some helper functions in a twinBasic-built external library; check that next
        If (InitializeEngine And PluginManager.IsPluginCurrentlyEnabled(CCP_PDHelper)) Then
            
            Dim initError As heif_error
            initError = PD_heif_init()
            
            'Convert the returned pointer to an error object
            InitializeEngine = (initError.heif_error_code = heif_error_Ok)
            m_LibAvailable = InitializeEngine
            If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "libheif helper library returned: " & Strings.StringFromCharPtr(initError.pCharMessage, False) & " for heif_init"
            
            If InitializeEngine Then
                If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "libheif initialized successfully."
            Else
                PDDebug.LogAction "WARNING: libheif initialization failed."
            End If
            
        Else
            PDDebug.LogAction "WARNING: failed to initialize helper library."
            m_LibAvailable = False
        End If
        
    Else
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load libheif.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & strLibPath & ")"
    End If
    
End Function

'Check a file to see if it's a valid heif image.  Note that PD will return "true" even if libheif
' returns "maybe"; this is necessary to work around some valid heif files whose contents are not
' guaranteed valid until libheif actually attempts to parse the full image.
Public Function IsFileHeif(ByRef srcFile As String) As Boolean
    
    IsFileHeif = False
    
    If m_LibAvailable Then
        
        'libheif asks for at least 12-bytes, but we can pass more to increase chance of success
        
        'Open a stream on the target file and load the first 128 bytes
        Dim cStream As pdStream
        Set cStream = New pdStream
        
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            Dim ptrPeek As Long, ptrSizeAvailable As Long
            ptrSizeAvailable = Files.FileLenW(srcFile)
            If (ptrSizeAvailable > 128) Then ptrSizeAvailable = 128
            ptrPeek = cStream.Peek_PointerOnly(0, ptrSizeAvailable)
            
            Dim fType As heif_filetype_result
            fType = CallCDeclW(heif_check_filetype, vbLong, ptrPeek, ptrSizeAvailable)
            IsFileHeif = (fType = heif_filetype_yes_supported)
            
            If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif_check_filetype returned " & fType
            
            'If the filetype is "maybe heif", the decoder needs more bytes before it can make a firm determination.
            ' Feed it more (up to an arbitrary limit) and try again.
            If (fType = heif_filetype_maybe) Then
                
                If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "Testing more bytes to see if file is valid..."
                
                Dim MAX_SIZE_TO_TEST As Long
                MAX_SIZE_TO_TEST = 1024000
                If (MAX_SIZE_TO_TEST > Files.FileLenW(srcFile)) Then MAX_SIZE_TO_TEST = Files.FileLenW(srcFile)
                ptrPeek = cStream.Peek_PointerOnly(0, MAX_SIZE_TO_TEST)
                
                fType = CallCDeclW(heif_check_filetype, vbLong, ptrPeek, MAX_SIZE_TO_TEST)
                If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "After checking " & CStr(MAX_SIZE_TO_TEST) & " bytes, new heif_check_filetype result is " & fType
                
                'If the file type is still a "maybe", attempt a full load.  (Note that I've gotten this even
                ' when scanning all the way to the end of some valid HEIC files, so it appears to be an
                ' expected and perfectly acceptable return value.)
                If (fType = heif_filetype_maybe) Or (fType = heif_filetype_yes_unsupported) Then
                    fType = heif_filetype_yes_supported
                    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "Will attempt to load anyway..."
                End If
                
                IsFileHeif = (fType = heif_filetype_yes_supported)
                
            'If the file is heif but is marked as unsupported, that's okay - this seems to happen occasionally
            ' with valid heif files that contain something like an image sequence (of which we don't necessarily
            ' need to load all frames).  Treat the file as potentially supported, and allow the loader to have a
            ' go at it.
            ElseIf (fType = heif_filetype_yes_unsupported) Then
                
                PDDebug.LogAction "FYI mimetype is " & Strings.StringFromCharPtr(CallCDeclW(heif_get_file_mime_type, vbLong, ptrPeek, 128&), False)
                
                Dim fSubType As Long
                fSubType = CallCDeclW(heif_read_main_brand, vbLong, ptrPeek, 128&)
                
                'Endianness needs to be reversed before decoding
                Dim revBytes(0 To 3) As Byte
                VBHacks.CopyMemoryStrict VarPtr(revBytes(0)), VarPtr(fSubType), 4
                VBHacks.SwapEndianness32 revBytes
                PDDebug.LogAction "File brand is listed as " & Strings.StringFromCharPtr(VarPtr(revBytes(0)), False, 4, True) & ". PD will attempt to load anyway..."
                
                'Allow the decoder to have a swing at the file
                fType = heif_filetype_yes_supported
                IsFileHeif = (fType = heif_filetype_yes_supported)
                
            End If
            
        End If
        
        cStream.StopStream True
        
    End If
    
End Function

Public Function IsLibheifEnabled() As Boolean
    IsLibheifEnabled = m_LibAvailable
End Function

Public Function LoadHeifImage(ByRef srcFilename As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByRef fileAlreadyValidated As Boolean = False, Optional ByVal previewOnly As Boolean = False) As Boolean
    
    Const FUNC_NAME As String = "LoadHeifImage"
    LoadHeifImage = False
    
    'Validate as necessary
    If (Not fileAlreadyValidated) Then
        If (Not Plugin_Heif.IsFileHeif(srcFilename)) Then Exit Function
    End If
    
    'Create a new heif context
    Dim ctxHeif As Long
    ctxHeif = CallCDeclW(heif_context_alloc, vbLong)
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif context created successfully (" & CStr(ctxHeif) & ")"
    
    'Attempt to load
    Dim filenameAsUTF8() As Byte, numChars As Long
    If (Not Strings.UTF8FromString(srcFilename, filenameAsUTF8, numChars)) Then GoTo LoadFailed
    If (numChars <= 0) Then GoTo LoadFailed
    
    Dim hReturn As heif_error
    hReturn = PD_heif_context_read_from_file(ctxHeif, VarPtr(filenameAsUTF8(0)), 0&)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo LoadFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: PD_heif_context_read_from_file returned successfully"
    
    'ctxHeif now contains a decoder for the target filename.
    
    'Before preceding, pause to prep some UI bits.  (We only do this on non-batch-process operations.)
    Dim updateUI As Boolean
    updateUI = (Not previewOnly) And (Macros.GetMacroStatus() <> MacroBATCH) And (Macros.GetMacroStatus() <> MacroPLAYBACK)
    
    'See how many top-level images exist inside the file.  (This ignores thumbnails, sub-tiles, etc.)
    Dim numImages As Long
    numImages = CallCDeclW(heif_context_get_number_of_top_level_images, vbLong, ctxHeif)
    PDDebug.LogAction "Number of images in file: " & numImages
    If (numImages <= 0) Then
        InternalError FUNC_NAME, "no images in file"
        GoTo LoadFailed
    End If
    
    'Extract a list of image IDs; we use these to access individual images inside the file
    Dim listOfImageIDs() As Long
    ReDim listOfImageIDs(0 To numImages - 1) As Long
    
    Dim numImagesFailsafe As Long
    numImagesFailsafe = CallCDeclW(heif_context_get_list_of_top_level_image_IDs, vbLong, ctxHeif, VarPtr(listOfImageIDs(0)), numImages)
    If (numImagesFailsafe <> numImages) Then
        InternalError FUNC_NAME, "mismatched image count"
        GoTo LoadFailed
    End If
    
    'We also want to know the ID of the primary image.  Its dimensions will be used for the
    ' image object's dimensions, and we'll also auto-set this frame as the active layer.
    Dim idPrimaryImage As Long
    hReturn = PD_heif_context_get_primary_image_ID(ctxHeif, idPrimaryImage)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo LoadFailed
    End If
    
    Dim idxPrimaryImage As Long: idxPrimaryImage = 0
    
    Dim i As Long
    For i = 0 To numImages - 1
        If (listOfImageIDs(i) = idPrimaryImage) Then
            idxPrimaryImage = i
            If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "Primary image is at idx " & i
        End If
    Next i
    
    If updateUI Then
        
        'We'll update the progress bar twice on each image:
        ' 1) Once after decoding
        ' 2) A second time after swizzling RGB (which is slower because we have to do it in VB code)
        ProgressBars.SetProgBarMax numImages * 2
        ProgressBars.SetProgBarVal 0
        
    End If
    
    'UPDATE NOV 2025: libheif has added new APIs for animation sequences.  These can exist alongside
    ' image series, and can be accessed separately.  I've yet to find any sample images to test this on,
    ' but here's how you can query state:
    If HEIF_DEBUG_VERBOSE Then
        Dim heifHasSequence As Boolean
        heifHasSequence = (CallCDeclW(heif_context_has_sequence, vbLong, ctxHeif) <> 0)
        PDDebug.LogAction "heifHasSequence: " & heifHasSequence
        'TODO someday: implement fully, see https://github.com/strukturag/libheif/wiki/Reading-and-Writing-Sequences
    End If
    
    'Ensure we have a target image to work with
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    
    'We now want to iterate each top-level image in turn, loading it to a unique layer in the target image.
    Dim idxImage As Long
    For idxImage = 0 To numImages - 1
        
        If updateUI Then ProgressBars.SetProgBarVal (idxImage * 2)
        
        'Get a handle to this image
        Dim hImg As Long
        hReturn = PD_heif_context_get_image_handle(ctxHeif, listOfImageIDs(idxImage), hImg)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            hImg = 0
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo LoadFailed
        End If
        
        If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif_context_get_image_handle returned successfully"
        
        'Retrieve frame dimensions and assign to the parent image as relevant
        Dim imgWidth As Long, imgHeight As Long, imgHasAlpha As Boolean, imgAlphaPremultiplied As Boolean
        imgWidth = CallCDeclW(heif_image_handle_get_width, vbLong, hImg)
        imgHeight = CallCDeclW(heif_image_handle_get_height, vbLong, hImg)
        imgHasAlpha = (CallCDeclW(heif_image_handle_has_alpha_channel, vbLong, hImg) <> 0)
        imgAlphaPremultiplied = (CallCDeclW(heif_image_handle_is_premultiplied_alpha, vbLong, hImg) <> 0)
        If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "heif stats are (w,h,a,pa): " & imgWidth & ", " & imgHeight & ", " & imgHasAlpha & ", " & imgAlphaPremultiplied
        
        If (idxImage = 0) Or (idxImage = idxPrimaryImage) Then
            
            dstImage.Width = imgWidth
            dstImage.Height = imgHeight
            dstImage.SetDPI 96, 96  'libheif doesn't provide a native way to get/set resolution (at present)
            
            dstImage.SetOriginalAlpha imgHasAlpha
            If imgHasAlpha Then
                dstImage.SetOriginalColorDepth 32
            Else
                dstImage.SetOriginalColorDepth 24
            End If
            
            'Also pull monochrome tagging from the image handle
            Dim imgColorSpace As heif_colorspace, imgChromaSpace As heif_chroma
            hReturn = PD_heif_image_handle_get_preferred_decoding_colorspace(hImg, imgColorSpace, imgChromaSpace)
            If (hReturn.heif_error_code <> heif_error_Ok) Then
                InternalErrorHeif FUNC_NAME, hReturn
                GoTo LoadFailed
            End If
            
            dstImage.SetOriginalGrayscale (imgChromaSpace = heif_chroma_monochrome) Or (imgColorSpace = heif_colorspace_monochrome)
            
        End If
        
        Dim imgWidthBytes As Long
        imgWidthBytes = imgWidth * 4
        
        '// decode the image and convert colorspace to RGB, saved as 24bit interleaved
        Dim hImgDecoded As Long
        hReturn = PD_heif_decode_image(hImg, hImgDecoded, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, ByVal 0&)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            hImgDecoded = 0
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo LoadFailed
        End If
        
        If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: PD_heif_decode_image returned successfully"
        
        'Retrieve a pointer to the decoded pixel data, as well as the stride (libheif doesn't necessarily guarantee
        ' a specific line alignment).
        Dim imgStride As Long, pData As Long
        pData = CallCDeclW(heif_image_get_plane_readonly, vbLong, hImgDecoded, heif_channel_interleaved, VarPtr(imgStride))
        If (pData = 0) Then
            InternalError FUNC_NAME, "bad pixel pointer"
            GoTo LoadFailed
        End If
        
        If updateUI Then ProgressBars.SetProgBarVal (idxImage * 2) + 1
        
        'Prepare a destination DIB
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        dstDIB.CreateBlank imgWidth, imgHeight, 32, 0, 0
        dstDIB.SetInitialAlphaPremultiplicationState imgAlphaPremultiplied
        
        'Iterate lines and copy the data directly into the target DIB
        Dim dstPixels() As Byte, dstSA1D As SafeArray1D, srcPixels() As Byte, srcSA1D As SafeArray1D
        
        Dim x As Long, y As Long
        For y = 0 To imgHeight - 1
            dstDIB.WrapArrayAroundScanline dstPixels, dstSA1D, y
            VBHacks.WrapArrayAroundPtr_Byte srcPixels, srcSA1D, pData + (imgStride * y), imgStride
        For x = 0 To imgWidthBytes - 1 Step 4
            
            'Manually un-interleave to fix pixel order
            dstPixels(x) = srcPixels(x + 2)
            dstPixels(x + 1) = srcPixels(x + 1)
            dstPixels(x + 2) = srcPixels(x)
            dstPixels(x + 3) = srcPixels(x + 3)
            
        Next x
        Next y
        
        VBHacks.UnwrapArrayFromPtr_Byte srcPixels
        dstDIB.UnwrapArrayFromDIB dstPixels
        
        'heif alpha channels can be premultiplied
        If (Not imgAlphaPremultiplied) Then dstDIB.SetAlphaPremultiplication True
        
        'Next, let's see if color management is relevant to this image.
        Dim embeddedProfile As pdICCProfile
        
        Dim colorProfileType As Long
        colorProfileType = CallCDeclW(heif_image_handle_get_color_profile_type, vbLong, hImg)
        If (colorProfileType <> 0) Then
            
            Dim colorProfileName As String
            colorProfileName = StrReverse(Strings.StringFromCharPtr(VarPtr(colorProfileType), False, 4, True))
            PDDebug.LogAction "Found color profile: " & colorProfileName
            
            'nclx is an old QuickTime color profile format.  We only want to handle full ICC profiles
            ' (which have two separate IDs in libheif, the distinction between which is irrelevant to us:
            ' https://github.com/strukturag/libheif/issues/119)
            If (colorProfileName = "prof") Or (colorProfileName = "rICC") Then
                
                Set embeddedProfile = New pdICCProfile
                
                Dim psize As Long, profData() As Byte
                psize = CallCDeclW(heif_image_handle_get_raw_color_profile_size, vbLong, hImg)
                
                'Retrieve the ICC profile, and ignore errors (if any)
                If (psize > 0) Then
                    ReDim profData(0 To psize - 1) As Byte
                    hReturn = PD_heif_image_handle_get_raw_color_profile(hImg, VarPtr(profData(0)))
                    If (hReturn.heif_error_code = heif_error_Ok) Then embeddedProfile.LoadICCFromPtr psize, VarPtr(profData(0))
                End If
                
            End If
            
            'Apply the color profile to the loaded data
            If (Not embeddedProfile Is Nothing) Then
            If embeddedProfile.HasICCData() And ColorManagement.UseEmbeddedICCProfiles() Then
                
                If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "Applying color profile to image..."
                
                Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, tmpTransform As pdLCMSTransform
                Set srcProfile = New pdLCMSProfile
                
                'Ignore monochrome profiles (we can't apply them to pixels that have already been expanded to RGB)
                If srcProfile.CreateFromPDICCObject(embeddedProfile) And (srcProfile.GetColorSpace <> cmsSigGray) Then
                    
                    dstDIB.SetAlphaPremultiplication False
                    
                    'For now, do a hard-convert into sRGB format
                    Set dstProfile = New pdLCMSProfile
                    dstProfile.CreateSRGBProfile
                    
                    Dim srcFormat As LCMS_PIXEL_FORMAT
                    srcFormat = TYPE_BGRA_8
                    
                    Dim flgTransform As LCMS_TRANSFORM_FLAGS
                    flgTransform = cmsFLAGS_COPY_ALPHA
                    Set tmpTransform = New pdLCMSTransform
                    tmpTransform.CreateTwoProfileTransform srcProfile, dstProfile, srcFormat, TYPE_BGRA_8, INTENT_PERCEPTUAL, flgTransform
                    tmpTransform.ApplyTransformToArbitraryMemoryEx dstDIB.GetDIBPointer, dstDIB.GetDIBPointer, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, dstDIB.GetDIBStride, dstDIB.GetDIBStride, 0&, 0&
                    
                    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "ICC profile applied."
                    Set tmpTransform = Nothing
                    Set srcProfile = Nothing
                    
                    dstDIB.SetAlphaPremultiplication True
                    
                End If
                
            '/embedded profile exists and was retrieved successfully
            End If
            End If
        
        End If
        
        'With pixel data processed, we can now move the contents into a pdLayer object.
        
        'Prep a new layer object and initialize it
        Dim newLayerID As Long, tmpLayer As pdLayer
        newLayerID = dstImage.CreateBlankLayer()
        Set tmpLayer = dstImage.GetLayerByID(newLayerID)
        
        'We need a base layer name for each page.  In a single-frame image, default to the source filename.
        ' In a multi-frame image, use incrementing frame numbers.
        Dim baseLayerName As String
        If (numImages = 1) Then
            baseLayerName = Files.FileGetName(srcFilename, True)
        Else
            baseLayerName = g_Language.TranslateMessage("Frame %1", idxImage + 1)
        End If
        
        tmpLayer.InitializeNewLayer PDL_Image, baseLayerName, dstDIB, True
        
        'Make the primary frame visible, but no others.
        tmpLayer.SetLayerVisibility (idxImage = idxPrimaryImage)
        tmpLayer.SetLayerBlendMode BM_Normal
        
        'Before exiting, release everything allocated in reverse order.
        If (hImgDecoded <> 0) Then
            CallCDeclW heif_image_release, vbEmpty, hImgDecoded
            hImgDecoded = 0
        Else
            PDDebug.LogAction "WARNING: hImgDecodedBase is 0"
        End If
        If (hImg <> 0) Then
            CallCDeclW heif_image_handle_release, vbEmpty, hImg
            hImg = 0
        Else
            PDDebug.LogAction "WARNING: hImgBase is 0"
        End If
        
    Next idxImage
    
    'Set the primary frame as the active one
    dstImage.SetActiveLayerByIndex idxPrimaryImage
    
    'If a color profile was applied, tag the pdImage accordingly
    If (Not embeddedProfile Is Nothing) And ColorManagement.UseEmbeddedICCProfiles() Then
        
        If embeddedProfile.HasICCData() Then
            
            Dim profHash As String
            profHash = ColorManagement.AddProfileToCache(embeddedProfile, True, False, False)
            dstImage.SetColorProfile_Original profHash
            
            'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
            ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            dstDIB.CreateBlank 16, 16, 32, 0
            dstDIB.SetColorManagementState cms_ProfileConverted
            
            'IMPORTANT NOTE: at present, the destination image - by the time we're done with it - will have been
            ' hard-converted to sRGB, so we don't want to associate the destination DIB with its source profile.
            ' Instead, note that it is currently sRGB.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            
        End If
            
    End If
    
    'Free any remaining objects
    If (ctxHeif <> 0) Then
        CallCDeclW heif_context_free, vbEmpty, ctxHeif
        ctxHeif = 0
    End If
    
    'Unload any changes made to the primary app UI
    If updateUI Then ProgressBars.ReleaseProgressBar
    
    LoadHeifImage = True
    
    Exit Function
    
LoadFailed:
    
    PDDebug.LogAction "Critical error during heif load; exiting now..."
    
    'Free any allocated resources before exiting
    If (hImgDecoded <> 0) Then CallCDeclW heif_image_release, vbEmpty, hImgDecoded
    If (hImg <> 0) Then CallCDeclW heif_image_handle_release, vbEmpty, hImg
    If (ctxHeif <> 0) Then CallCDeclW heif_context_free, vbEmpty, ctxHeif
    
    'Unload any changes made to the primary app UI
    If updateUI Then ProgressBars.ReleaseProgressBar
    
    LoadHeifImage = False
    
End Function

'Round-trip a pdDIB image through libheif to preview compression results
Public Function PreviewHEIF(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef srcOptions As String) As Boolean

    Const FUNC_NAME As String = "PreviewHEIF"
    PreviewHEIF = False
    
    'Retrieve and validate parameters from incoming string.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcOptions
    
    Dim exportLossless As Boolean, exportQuality As Single
    exportLossless = cParams.GetBool("heif-lossless", False, True)
    exportQuality = cParams.GetLong("heif-lossy-quality", 90, True)
    
    If (exportQuality < 0) Then exportQuality = 0
    If (exportQuality > 100) Then exportQuality = 100
    
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'If lossless compression is being used, exit now (no preview required)
    If exportLossless Then
        dstDIB.CreateFromExistingDIB srcDIB
        If (Not srcDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
        PreviewHEIF = True
        Exit Function
    Else
        dstDIB.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, 0, 0
        dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
    End If
    
    'All remaining code handles lossy compression only
    
    'Create a new heif context
    Dim ctxHeif As Long
    ctxHeif = CallCDeclW(heif_context_alloc, vbLong)
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif context created successfully (" & CStr(ctxHeif) & ")"
    
    Dim hReturn As heif_error
    
    'Get a HEIF encoder
    Dim pHeifEncoder As Long
    hReturn = PD_heif_context_get_encoder_for_format(ctxHeif, heif_compression_HEVC, VarPtr(pHeifEncoder))
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        pHeifEncoder = 0
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    'Set encoder properties
    hReturn = PD_heif_encoder_set_lossless(pHeifEncoder, 0)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    hReturn = PD_heif_encoder_set_lossy_quality(pHeifEncoder, exportQuality)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif encoder ready"
    
    'Create a base heif image object (note that this does NOT allocate pixel data yet)
    Dim hImg As Long
    hReturn = PD_heif_image_create(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, VarPtr(hImg))
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        hImg = 0
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif image created"
    
    'Allocate the heif image
    hReturn = PD_heif_image_add_plane(hImg, heif_channel_interleaved, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 8)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    CallCDeclW heif_image_set_premultiplied_alpha, vbEmpty, hImg, IIf(srcDIB.GetAlphaPremultiplication, 1&, 0&)
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif pixels allocated"
    
    'Retrieve a pointer to the allocated image, as well as the stride (libheif doesn't necessarily guarantee
    ' a specific line alignment).
    Dim imgStride As Long, pData As Long
    pData = CallCDeclW(heif_image_get_plane, vbLong, hImg, heif_channel_interleaved, VarPtr(imgStride))
    If (pData = 0) Then
        InternalError FUNC_NAME, "bad pixel pointer"
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif image pointer retrieved"
    
    'Fill the destination image line-by-line
    Dim imgHeight As Long, imgWidthBytes As Long
    imgHeight = srcDIB.GetDIBHeight
    imgWidthBytes = srcDIB.GetDIBStride
    
    Dim dstPixels() As Byte, dstSA1D As SafeArray1D, srcPixels() As Byte, srcSA1D As SafeArray1D
    
    Dim x As Long, y As Long
    For y = 0 To imgHeight - 1
        srcDIB.WrapArrayAroundScanline srcPixels, srcSA1D, y
        VBHacks.WrapArrayAroundPtr_Byte dstPixels, dstSA1D, pData + (imgStride * y), imgStride
    For x = 0 To imgWidthBytes - 1 Step 4
        
        'Manually un-interleave to fix pixel order
        dstPixels(x) = srcPixels(x + 2)
        dstPixels(x + 1) = srcPixels(x + 1)
        dstPixels(x + 2) = srcPixels(x)
        dstPixels(x + 3) = srcPixels(x + 3)
        
    Next x
    Next y
    
    VBHacks.UnwrapArrayFromPtr_Byte dstPixels
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: copied pixel data to heif image"
    
    'Encode the image, and unlike a normal save, retrieve a handle to the encoded image.
    ' CHANGE OF PLANS AUGUST 2024: libheif crashes when loading an image handle that did
    ' not originate from a fully encoded file (see https://github.com/strukturag/libheif/issues/1168).
    ' That breaks this whole smart plan to round-trip encoding without requiring a redundant
    ' "write actual file to memory/disk" step.
    '
    'TODO: track the GitHub issue and rework this when a fix is available.
    'Dim hImgEncoded As Long
    'hReturn = PD_heif_context_encode_image(ctxHeif, hImg, pHeifEncoder, 0&, hImgEncoded)
    hReturn = PD_heif_context_encode_image(ctxHeif, hImg, pHeifEncoder, 0&, ByVal 0&)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "HEIF encoding complete"
    
    'Free the encoder
    CallCDeclW heif_encoder_release, vbEmpty, pHeifEncoder
    pHeifEncoder = 0
    
    'For now, dump to a temp file
    If (m_utfLen = 0) Then
        m_tmpFile = OS.UniqueTempFilename(customExtension:="heif")
        Strings.UTF8FromString m_tmpFile, m_utfFilename, m_utfLen
    End If
    
    Files.FileDeleteIfExists m_tmpFile
    
    hReturn = PD_heif_context_write_to_file(ctxHeif, VarPtr(m_utfFilename(0)))
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    'Free the raw image?  The docs are ambiguous on whether we can do it earlier...
    CallCDeclW heif_image_release, vbEmpty, hImg
    hImg = 0
    
    'Releast this context then create a new one
    CallCDeclW heif_context_free, vbEmpty, ctxHeif
    ctxHeif = CallCDeclW(heif_context_alloc, vbLong)
    
    hReturn = PD_heif_context_read_from_file(ctxHeif, VarPtr(m_utfFilename(0)), 0&)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    'Get a handle to the primary image
    Dim hImgAfter As Long
    hReturn = PD_heif_context_get_primary_image_handle(ctxHeif, hImgAfter)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    'Decode the image.
    ' (Refer to the above heif_context_encode_image to see a better way to handle this, pending fixes in libheif;
    ' in anticipation of that bug-fix, I've left the required line of code below.)
    Dim hImgDecoded As Long
    hReturn = PD_heif_decode_image(hImgAfter, hImgDecoded, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, ByVal 0&)
    'hReturn = PD_heif_decode_image(hImgEncoded, hImgDecoded, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, ByVal 0&)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        hImgDecoded = 0
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: encoded image decoded"
    
    'Free the pre-decoding image
    CallCDeclW heif_image_handle_release, vbEmpty, hImgAfter
    hImgAfter = 0
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: PD_heif_decode_image returned successfully"
    
    'Retrieve a pointer to the decoded pixel data, as well as the stride (libheif doesn't necessarily guarantee
    ' a specific line alignment).
    pData = CallCDeclW(heif_image_get_plane_readonly, vbLong, hImgDecoded, heif_channel_interleaved, VarPtr(imgStride))
    If (pData = 0) Then
        InternalError FUNC_NAME, "bad pixel pointer"
        GoTo PreviewFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: pointer to decoded image ready"
    
    dstDIB.SetInitialAlphaPremultiplicationState False
    
    'Iterate lines and copy the data directly into the target DIB
    For y = 0 To imgHeight - 1
        dstDIB.WrapArrayAroundScanline dstPixels, dstSA1D, y
        VBHacks.WrapArrayAroundPtr_Byte srcPixels, srcSA1D, pData + (imgStride * y), imgStride
    For x = 0 To imgWidthBytes - 1 Step 4
        
        'Manually un-interleave to fix pixel order
        dstPixels(x) = srcPixels(x + 2)
        dstPixels(x + 1) = srcPixels(x + 1)
        dstPixels(x + 2) = srcPixels(x)
        dstPixels(x + 3) = srcPixels(x + 3)
        
    Next x
    Next y
    
    VBHacks.UnwrapArrayFromPtr_Byte srcPixels
    dstDIB.UnwrapArrayFromDIB dstPixels
    dstDIB.SetAlphaPremultiplication True
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: decoded pixels copied to target DIB"
    
    'Before exiting, release the decoded image object
    If (hImgDecoded <> 0) Then
        CallCDeclW heif_image_release, vbEmpty, hImgDecoded
        hImgDecoded = 0
    End If
    
    'Free the context
    CallCDeclW heif_context_free, vbEmpty, ctxHeif
    ctxHeif = 0
    
    'Kill the temp file
    Files.FileDeleteIfExists m_tmpFile
    
    PreviewHEIF = True
    
    Exit Function
    
PreviewFailed:
    PreviewHEIF = False
    InternalError FUNC_NAME, "VB error # " & Err.Number

    'Free any allocated resources before exiting
    If (hImg <> 0) Then CallCDeclW heif_image_release, vbEmpty, hImg
    If (hImgDecoded <> 0) Then CallCDeclW heif_image_release, vbEmpty, hImgDecoded
    If (hImgAfter <> 0) Then CallCDeclW heif_image_handle_release, vbEmpty, hImgAfter
    If (pHeifEncoder <> 0) Then CallCDeclW heif_encoder_release, vbEmpty, pHeifEncoder
    If (ctxHeif <> 0) Then CallCDeclW heif_context_free, vbEmpty, ctxHeif
    
    Files.FileDeleteIfExists m_tmpFile
    
End Function

'Save an arbitrary pdImage object to a standalone HEIF file.
Public Function SaveHEIF_ToFile(ByRef srcImage As pdImage, ByRef srcOptions As String, ByRef dstFile As String) As Boolean

    Const FUNC_NAME As String = "SaveHEIF_ToFile"
    SaveHEIF_ToFile = False
    
    'Retrieve and validate parameters from incoming string.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcOptions
    
    Dim exportLossless As Boolean, exportQuality As Single, exportMultipage As Boolean
    exportLossless = cParams.GetBool("heif-lossless", False, True)
    exportQuality = cParams.GetLong("heif-lossy-quality", 90, True)
    exportMultipage = cParams.GetBool("heif-multiframe", False, True)
    
    If (exportQuality < 0) Then exportQuality = 0
    If (exportQuality > 100) Then exportQuality = 100
    
    'Retrieve the composited pdImage object (if exporting as a composite image)
    Dim finalDIB As pdDIB
    If (Not exportMultipage) Or (srcImage.GetNumOfLayers <= 1) Then srcImage.GetCompositedImage finalDIB, False
    
    'Create a new heif context
    Dim ctxHeif As Long
    ctxHeif = CallCDeclW(heif_context_alloc, vbLong)
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif context created successfully (" & CStr(ctxHeif) & ")"
    
    Dim hReturn As heif_error
    
    'Get a HEIF encoder
    Dim pHeifEncoder As Long
    hReturn = PD_heif_context_get_encoder_for_format(ctxHeif, heif_compression_HEVC, VarPtr(pHeifEncoder))
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        pHeifEncoder = 0
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo SaveFailed
    End If
    
    'Set encoder properties
    If exportLossless Then
        hReturn = PD_heif_encoder_set_lossless(pHeifEncoder, 1)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
    Else
        hReturn = PD_heif_encoder_set_lossless(pHeifEncoder, 0)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
        hReturn = PD_heif_encoder_set_lossy_quality(pHeifEncoder, exportQuality)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
    End If
    
    'When exporting a multi-frame image, the handle to the encoded image representing the active layer
    ' (if any) will be stored here.  You do need to check the case of handle = 0, in case the active
    ' layer is e.g. invisible.
    Dim hActiveFrame As Long
    
    'This function can iterate layers and export each one as a unique frame in the target HEIF.
    Dim idxStart As Long, idxEnd As Long
    If exportMultipage Then
        idxStart = 0
        idxEnd = srcImage.GetNumOfLayers - 1
    Else
        idxStart = -1
        idxEnd = -1
    End If
    
    Dim i As Long
    For i = idxStart To idxEnd
        
        'During multipage export, grab a copy of each target layer in turn
        If (i >= 0) Then
            
            If (finalDIB Is Nothing) Then Set finalDIB = New pdDIB
            
            'Account for affine transforms in the current layer, as necessary
            If srcImage.GetLayerByIndex(i).AffineTransformsActive(True) Then
                srcImage.GetLayerByIndex(i).GetAffineTransformedDIB finalDIB, 0, 0
            Else
                finalDIB.CreateFromExistingDIB srcImage.GetLayerByIndex(i).GetLayerDIB
            End If
            
        End If
        
        'Construct a heif image
        Dim hImg As Long
        hReturn = PD_heif_image_create(finalDIB.GetDIBWidth, finalDIB.GetDIBHeight, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, VarPtr(hImg))
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            hImg = 0
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
        
        'Allocate the heif image
        hReturn = PD_heif_image_add_plane(hImg, heif_channel_interleaved, finalDIB.GetDIBWidth, finalDIB.GetDIBHeight, 8)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
        
        CallCDeclW heif_image_set_premultiplied_alpha, vbEmpty, hImg, IIf(finalDIB.GetAlphaPremultiplication, 1&, 0&)
        
        'Retrieve a pointer to the allocated image, as well as the stride (libheif doesn't necessarily guarantee
        ' a specific line alignment).
        Dim imgStride As Long, pData As Long
        pData = CallCDeclW(heif_image_get_plane, vbLong, hImg, heif_channel_interleaved, VarPtr(imgStride))
        If (pData = 0) Then
            InternalError FUNC_NAME, "bad pixel pointer"
            GoTo SaveFailed
        End If
        
        'Fill the destination image line-by-line
        Dim imgHeight As Long, imgWidthBytes As Long
        imgHeight = finalDIB.GetDIBHeight
        imgWidthBytes = finalDIB.GetDIBStride
        
        Dim dstPixels() As Byte, dstSA1D As SafeArray1D, srcPixels() As Byte, srcSA1D As SafeArray1D
        
        Dim x As Long, y As Long
        For y = 0 To imgHeight - 1
            finalDIB.WrapArrayAroundScanline srcPixels, srcSA1D, y
            VBHacks.WrapArrayAroundPtr_Byte dstPixels, dstSA1D, pData + (imgStride * y), imgStride
        For x = 0 To imgWidthBytes - 1 Step 4
            
            'Manually un-interleave to fix pixel order
            dstPixels(x) = srcPixels(x + 2)
            dstPixels(x + 1) = srcPixels(x + 1)
            dstPixels(x + 2) = srcPixels(x)
            dstPixels(x + 3) = srcPixels(x + 3)
            
        Next x
        Next y
        
        VBHacks.UnwrapArrayFromPtr_Byte dstPixels
        finalDIB.UnwrapArrayFromDIB srcPixels
        
        'Encode the image.  Note that the final parameter of this call can return a handle to the encoded object.
        Dim hImgEncoded As Long
        hReturn = PD_heif_context_encode_image(ctxHeif, hImg, pHeifEncoder, 0&, hImgEncoded)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalErrorHeif FUNC_NAME, hReturn
            GoTo SaveFailed
        End If
        
        'Now that the image is encoded, we can free the raw image.
        CallCDeclW heif_image_release, vbEmpty, hImg
        hImg = 0
        
        'In multi-frame images, we need to manually set the active frame (if any)
        If (i >= 0) Then
            
            'If the active frame is visible, it can cause problems, so ensure a failsafe handle is grabbed
            If (hActiveFrame = 0) Then hActiveFrame = hImgEncoded
            
            'With the backup case covered, use the currently active layer (if any) as the active frame
            If (i = srcImage.GetActiveLayerIndex) Then hActiveFrame = hImgEncoded
            
        End If
        
        If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "HEIF encoding complete"
        
NextLayer:
    Next i
    
    'Free the encoder
    CallCDeclW heif_encoder_release, vbEmpty, pHeifEncoder
    pHeifEncoder = 0
    
    'If encoding multi-frame images, set the active frame now
    If (hActiveFrame <> 0) Then
        hReturn = PD_heif_context_set_primary_image(ctxHeif, hActiveFrame)
        If (hReturn.heif_error_code <> heif_error_Ok) Then
            InternalError FUNC_NAME, "heif_context_set_primary_image failed; base layer will be used instead"
        End If
    End If
    
    'Dump to file
    Dim utfFilename() As Byte, utfLen As Long
    Files.FileDeleteIfExists dstFile
    Strings.UTF8FromString dstFile, utfFilename, utfLen
    hReturn = PD_heif_context_write_to_file(ctxHeif, VarPtr(utfFilename(0)))
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo SaveFailed
    End If
    
    'Free the context
    CallCDeclW heif_context_free, vbEmpty, ctxHeif
    ctxHeif = 0
    
    SaveHEIF_ToFile = True
    
    Exit Function
    
SaveFailed:
    SaveHEIF_ToFile = False
    InternalError FUNC_NAME, "VB error # " & Err.Number
    
    'Free any allocated resources before exiting
    If (hImg <> 0) Then CallCDeclW heif_image_release, vbEmpty, hImg
    If (pHeifEncoder <> 0) Then CallCDeclW heif_encoder_release, vbEmpty, pHeifEncoder
    If (ctxHeif <> 0) Then CallCDeclW heif_context_free, vbEmpty, ctxHeif
    
End Function

Public Sub ReleaseEngine()
    
    'For extra safety, free in reverse order from loading
    If PluginManager.IsPluginCurrentlyEnabled(CCP_PDHelper) Then
        PD_heif_deinit  'Start by calling the internal ref-counted deinitialize function
    End If
    If (m_hLibHeif <> 0) Then
        VBHacks.FreeLib m_hLibHeif
        m_hLibHeif = 0
    End If
    If (m_hLibde265 <> 0) Then
        VBHacks.FreeLib m_hLibde265
        m_hLibde265 = 0
    End If
    If (m_hLibx265 <> 0) Then
        VBHacks.FreeLib m_hLibx265
        m_hLibx265 = 0
    End If
        
    m_LibAvailable = False
    
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As Libheif_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        If VarType(pa(i)) = vbString Then vTemp(i) = StrPtr(pa(i))
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function

Private Sub InternalErrorHeif(ByRef pdFuncName As String, ByRef heifErrorObj As heif_error)
    PDDebug.LogAction "WARNING! libheif error in PD function: " & pdFuncName
    PDDebug.LogAction "libheif error #" & heifErrorObj.heif_error_code & ": " & Strings.StringFromCharPtr(heifErrorObj.pCharMessage, False)
End Sub

Private Sub InternalError(ByRef pdFuncName As String, ByRef errText As String)
    PDDebug.LogAction "WARNING! error in PD function """ & pdFuncName & """: " & errText
End Sub
