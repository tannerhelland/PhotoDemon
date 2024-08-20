Attribute VB_Name = "Plugin_Heif"
'***************************************************************************
'libheif Library Interface
'Copyright 2024-2024 by Tanner Helland
'Created: 16/July/24
'Last updated: 20/August/24
'Last update: finish proof-of-concept image importing
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Log extra heif debug info; recommend DISABLING in production builds
Private Const HEIF_DEBUG_VERBOSE As Boolean = True

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
    heif_image_handle_get_width
    heif_image_handle_get_height
    heif_image_handle_has_alpha_channel
    heif_image_handle_is_premultiplied_alpha
    heif_image_handle_release
    heif_image_release
    heif_image_get_plane_readonly
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

Private Type heif_image_handle
    hImg As Long
    hContext As Long
End Type

'Some libheif functions return a custom 12-byte type.  This cannot be easily handled via DispCallFunc,
' so I've written an interop DLL that marshals the cdecl return for me.
Private m_hHelper As Long
Private Declare Function PD_heif_init Lib "PDHelper_win32" () As heif_error
Private Declare Sub PD_heif_deinit Lib "PDHelper_win32" ()
Private Declare Function PD_heif_context_read_from_file Lib "PDHelper_win32" (ByVal heif_context As Long, ByVal pCharFilename As Long, ByVal pHeifReadingOptions As Long) As heif_error
Private Declare Function PD_heif_context_get_primary_image_handle Lib "PDHelper_win32" (ByVal ptr_heif_context As Long, ByRef ptr_heif_image_handle As Long) As heif_error
Private Declare Function PD_heif_decode_image Lib "PDHelper_win32" (ByVal in_heif_image_handle As Long, ByRef pp_out_heif_image As Long, ByVal in_heif_colorspace As heif_colorspace, ByVal in_heif_chroma As heif_chroma, ByVal p_heif_decoding_options As Long) As heif_error

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
        m_ProcAddresses(heif_image_handle_get_width) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_width")
        m_ProcAddresses(heif_image_handle_get_height) = GetProcAddress(m_hLibHeif, "heif_image_handle_get_height")
        m_ProcAddresses(heif_image_handle_has_alpha_channel) = GetProcAddress(m_hLibHeif, "heif_image_handle_has_alpha_channel")
        m_ProcAddresses(heif_image_handle_is_premultiplied_alpha) = GetProcAddress(m_hLibHeif, "heif_image_handle_is_premultiplied_alpha")
        m_ProcAddresses(heif_image_handle_release) = GetProcAddress(m_hLibHeif, "heif_image_handle_release")
        m_ProcAddresses(heif_image_release) = GetProcAddress(m_hLibHeif, "heif_image_release")
        m_ProcAddresses(heif_image_get_plane_readonly) = GetProcAddress(m_hLibHeif, "heif_image_get_plane_readonly")
        m_ProcAddresses(heif_get_file_mime_type) = GetProcAddress(m_hLibHeif, "heif_get_file_mime_type")
        m_ProcAddresses(heif_read_main_brand) = GetProcAddress(m_hLibHeif, "heif_read_main_brand")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
        'Now attempt to initialize the helper library
        strLibPath = pathToDLLFolder & "PDHelper_win32.dll"
        m_hHelper = VBHacks.LoadLib(strLibPath)
        InitializeEngine = (m_hHelper <> 0)
        
        If InitializeEngine Then
            
            Dim initError As heif_error
            initError = PD_heif_init()
            
            'Convert the returned pointer to an error object
            InitializeEngine = (initError.heif_error_code = heif_error_Ok)
            m_LibAvailable = InitializeEngine
            If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "libheif helper library returned: " & Strings.StringFromCharPtr(initError.pCharMessage, False) & " for heif_init"
            
            If InitializeEngine Then
                PDDebug.LogAction "libheif initialized successfully."
            Else
                PDDebug.LogAction "libheif initialization failed."
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

'Check a file to see if it's a valid heif image
Public Function IsFileHeif(ByRef srcFile As String) As Boolean
    
    IsFileHeif = False
    
    If m_LibAvailable Then
        
        'libheif asks for at least 12-bytes, but we can pass more to increase chance of success
        
        'Open a stream on the target file and load the first 128 bytes
        Dim cStream As pdStream
        Set cStream = New pdStream
        
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
            
            Dim ptrPeek As Long
            ptrPeek = cStream.Peek_PointerOnly(0, 128)
            
            Dim fType As heif_filetype_result
            fType = CallCDeclW(heif_check_filetype, vbLong, ptrPeek, 128&)
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
                ' when scanning all the way to the end of some valid HEIC files, so it's a perfectly acceptable
                ' return value.)
                If (fType = heif_filetype_maybe) Then
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

Public Function LoadHeifImage(ByRef srcFilename As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByRef fileAlreadyValidated As Boolean = False) As Boolean
    
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
    
    'ctxHeif now contains a decoder for the target filename.  The context can be queried for various image properties,
    ' but for this intial build, let's just jump ahead to loading the image as quickly as possible.
    
    'Get a handle to the primary image in the file (e.g. not a thumbnail, frame > 0, etc)
    Dim hImg As Long, hImgBase As heif_image_handle
    hImg = VarPtr(hImgBase)
    hReturn = PD_heif_context_get_primary_image_handle(ctxHeif, hImg)
    
    If (hReturn.heif_error_code <> heif_error_Ok) Then
        InternalErrorHeif FUNC_NAME, hReturn
        GoTo LoadFailed
    End If
    
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "FYI: heif_context_get_primary_image_handle returned successfully"
    
    'Retrieve image dimensions from the base image object
    Dim imgWidth As Long, imgHeight As Long, imgHasAlpha As Boolean, imgAlphaPremultiplied As Boolean
    imgWidth = CallCDeclW(heif_image_handle_get_width, vbLong, hImg)
    imgHeight = CallCDeclW(heif_image_handle_get_height, vbLong, hImg)
    imgHasAlpha = (CallCDeclW(heif_image_handle_has_alpha_channel, vbLong, hImg) <> 0)
    imgAlphaPremultiplied = (CallCDeclW(heif_image_handle_is_premultiplied_alpha, vbLong, hImg) <> 0)
    If HEIF_DEBUG_VERBOSE Then PDDebug.LogAction "heif stats are (w,h,a,pa): " & imgWidth & ", " & imgHeight & ", " & imgHasAlpha & ", " & imgAlphaPremultiplied
    
    Dim imgWidthBytes As Long
    imgWidthBytes = imgWidth * 4
    
    '// decode the image and convert colorspace to RGB, saved as 24bit interleaved
    'heif_image* img;
    'heif_decode_image(handle, &img, heif_colorspace_RGB, heif_chroma_interleaved_RGB, nullptr);
    Dim hImgDecoded As Long, hImgDecodedBase As Long
    hImgDecoded = VarPtr(hImgDecodedBase)
    hReturn = PD_heif_decode_image(hImg, hImgDecoded, heif_colorspace_RGB, heif_chroma_interleaved_RGBA, ByVal 0&)
    If (hReturn.heif_error_code <> heif_error_Ok) Then
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
    
    'Prepare the destination DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank imgWidth, imgHeight, 32, initialAlpha:=255
    dstDIB.SetInitialAlphaPremultiplicationState imgAlphaPremultiplied
    
    'Iterate lines and copy the data directly into the target DIB
    Dim dstPixels() As Byte, dstSA1D As SafeArray1D, srcPixels() As Byte, srcSA1D As SafeArray1D
    Dim r As Long, g As Long, b As Long, a As Long
    
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
    dstDIB.SetInitialAlphaPremultiplicationState False
    If (Not imgAlphaPremultiplied) Then dstDIB.SetAlphaPremultiplication True Else dstDIB.SetInitialAlphaPremultiplicationState True
    
    'TODO
    'Set additional image properties
    'dstImage.SetOriginalColorDepth numComponents * 4
    'dstImage.SetOriginalGrayscale False
    'dstImage.SetOriginalAlpha (numComponents = 4)
    
    'Before exiting, release everything allocated in reverse order
    If (hImgDecoded <> 0) Then
        CallCDeclW heif_image_release, vbEmpty, hImgDecoded
        hImgDecoded = 0
    End If
    If (hImg <> 0) Then
        CallCDeclW heif_image_handle_release, vbEmpty, hImg
        hImg = 0
    End If
    If (ctxHeif <> 0) Then
        CallCDeclW heif_context_free, vbEmpty, ctxHeif
        ctxHeif = 0
    End If
    
    LoadHeifImage = True
    
    Exit Function
    
LoadFailed:
    
    PDDebug.LogAction "Critical error during heif load; exiting now..."
    
    'Free any allocated resources before exiting
    If (hImgDecoded <> 0) Then CallCDeclW heif_image_release, vbEmpty, hImgDecoded
    If (hImg <> 0) Then CallCDeclW heif_image_handle_release, vbEmpty, hImg
    If (ctxHeif <> 0) Then CallCDeclW heif_context_free, vbEmpty, ctxHeif
    LoadHeifImage = False
    
End Function

Public Sub ReleaseEngine()
    
    'Start by calling the internal ref-counted deinitialize function
    PD_heif_deinit
    
    'For extra safety, free in reverse order from loading
    If (m_hHelper <> 0) Then
        VBHacks.FreeLib m_hHelper
        m_hHelper = 0
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
