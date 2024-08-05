Attribute VB_Name = "Plugin_Heif"
'***************************************************************************
'libheif Library Interface
'Copyright 2024-2024 by Tanner Helland
'Created: 16/July/24
'Last updated: 16/July/24
'Last update: initial build
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

'Some libheif functions return a custom 12-byte type.  This cannot be easily handled via DispCallFunc, so I've written
' an interop DLL that handles marshalling the cdecl return for me.
Private m_hHelper As Long
Private Declare Function PD_heif_init Lib "PDHelper_win32" () As heif_error
Private Declare Sub PD_heif_deinit Lib "PDHelper_win32" ()

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
            'Debug.Print initError.heif_error_code, initError.heif_suberror_code, initError.pCharMessage
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
            
        End If
        
        cStream.StopStream True
        
    End If
    
End Function

Public Function IsLibheifEnabled() As Boolean
    IsLibheifEnabled = m_LibAvailable
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
