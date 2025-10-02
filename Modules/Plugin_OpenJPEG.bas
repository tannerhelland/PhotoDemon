Attribute VB_Name = "Plugin_OpenJPEG"
'***************************************************************************
'OpenJPEG (JPEG-2000) Library Interface
'Copyright 2025-2025 by Tanner Helland
'Created: 19/September/25
'Last updated: 19/September/25
'Last update: initial build
'
'Per its documentation (available at https://www.openjpeg.org/), OpenJPEG is...
'
' "...an open-source JPEG 2000 codec written in C language. It has been developed in order to promote
'  the use of JPEG 2000, a still-image compression standard from the Joint Photographic Experts Group
'  (JPEG). Since may 2015, it is officially recognized by ISO/IEC and ITU-T as a JPEG 2000 Reference
'  Software."
'
'OpenJPEG is BSD-licensed and actively maintained.
'
'PhotoDemon originally used OpenJPEG via FreeImage, but FreeImage's abandonment made it impossible
' to continue down that route.  As such, in 2025 I wrote a new, direct OpenJPEG interface from scratch.
' (The only one of its kind in VB6, AFAIK.)
'
'This interface was designed against v2.5.3 of the library (released 09 Dec 2024).  It should work fine
' with any version that maintains ABI compatibility with that release.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Enable at DEBUG-TIME for verbose logging
Private Const J2K_DEBUG_VERBOSE As Boolean = True

'Library handle will be non-zero if CharLS is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'#define OPJ_PATH_LEN 4096 /**< Maximum allowed size for filenames */
Private Const OPJ_PATH_LEN As Long = 4096

Public Enum OPJ_CODEC_FORMAT
    OPJ_CODEC_UNKNOWN = -1  '/**< place-holder */
    OPJ_CODEC_J2K = 0       '/**< JPEG-2000 codestream : read/write */
    OPJ_CODEC_JPT = 1       '/**< JPT-stream (JPEG 2000, JPIP) : read only */
    OPJ_CODEC_JP2 = 2       '/**< JP2 file format : read/write */
    OPJ_CODEC_JPP = 3       '/**< JPP-stream (JPEG 2000, JPIP) : to be coded */
    OPJ_CODEC_JPX = 4       '/**< JPX file format (JPEG 2000 Part-2) : to be coded */
End Enum

Private Enum OPJ_COLOR_SPACE
    OPJ_CLRSPC_UNKNOWN = -1     '/**< not supported by the library */
    OPJ_CLRSPC_UNSPECIFIED = 0  '/**< not specified in the codestream */
    OPJ_CLRSPC_SRGB = 1         '/**< sRGB */
    OPJ_CLRSPC_GRAY = 2         '/**< grayscale */
    OPJ_CLRSPC_SYCC = 3         '/**< YUV */
    OPJ_CLRSPC_EYCC = 4         '/**< e-YCC */
End Enum


'/**
' * Decompression parameters
' * */
Private Type opj_dparameters

'    /**
'    Set the number of highest resolution levels to be discarded.
'    The image resolution is effectively divided by 2 to the power of the number of discarded levels.
'    The reduce factor is limited by the smallest total number of decomposition levels among tiles.
'    if != 0, then original dimension divided by 2^(reduce);
'    if == 0 or not used, image is decoded to the full resolution
'    */
    cp_reduce As Long   'OPJ_UINT32
'    /**
'    Set the maximum number of quality layers to decode.
'    If there are less quality layers than the specified number, all the quality layers are decoded.
'    if != 0, then only the first "layer" layers are decoded;
'    if == 0 or not used, all the quality layers are decoded
'    */
    cp_layer As Long    'OPJ_UINT32

'    /**@name command line decoder parameters (not used inside the library) */
'    /*@{*/
'    /** input file name */
    p_infile(OPJ_PATH_LEN) As Byte  'infile[OPJ_PATH_LEN];
'    /** output file name */
    p_outfile(OPJ_PATH_LEN) As Byte 'outfile[OPJ_PATH_LEN];
'    /** input file format 0: J2K, 1: JP2, 2: JPT */
    decod_format As Long    'int
'    /** output file format 0: PGX, 1: PxM, 2: BMP */
    cod_format As Long  'int

'    /** Decoding area left boundary */
    DA_x0 As Long   'OPJ_UINT32 ;
'    /** Decoding area right boundary */
    DA_x1 As Long   'OPJ_UINT32 ;
'    /** Decoding area up boundary */
    DA_y0 As Long   'OPJ_UINT32 ;
'    /** Decoding area bottom boundary */
    DA_y1 As Long   'OPJ_UINT32 ;
'    /** Verbose mode */
    m_verbose As Long   'OPJ_BOOL (explicitly typedef'd as int 0/1);

'    /** tile number of the decoded tile */
    tile_index As Long  'OPJ_UINT32 ;
'    /** Nb of tile to decode */
    nb_tile_to_decode As Long   'OPJ_UINT32 ;

'    /*@}*/

'    /* UniPG>> */ /* NOT YET USED IN THE V2 VERSION OF OPENJPEG */
'    /**@name JPWL decoding parameters */
'    /*@{*/
'    /** activates the JPWL correction capabilities */
    jpwl_correct As Long    'OPJ_BOOL ;
'    /** expected number of components */
    jpwl_exp_comps As Long  'int ;
'    /** maximum number of tiles */
    jpwl_max_tiles As Long  'int ;
'    /*@}*/
'    /* <<UniPG */

    Flags As Long   'unsigned int ;

End Type    'opj_dparameters_t;

'/**
' * Defines a single image component
' * */
Private Type opj_image_comp
    '/** XRsiz: horizontal separation of a sample of ith component with respect to the reference grid */
    dx As Long
    '/** YRsiz: vertical separation of a sample of ith component with respect to the reference grid */
    dy As Long
    '/** data width */
    w As Long
    '/** data height */
    h As Long
    '/** x component offset compared to the whole image */
    x0 As Long
    '/** y component offset compared to the whole image */
    y0 As Long
    '/** precision: number of bits per component per pixel */
    prec As Long
    '/** obsolete: use prec instead */
    opj_bpp As Long
    '/** signed (1) / unsigned (0) */
    sgnd As Long
    '/** number of decoded resolution */
    resno_decoded As Long
    '/** number of division by 2 of the out image compared to the original size of image */
    factor As Long
    '/** image component data */
    p_data As Long
    '/** alpha channel */
    Alpha As Integer
    'safe_padding As Integer
End Type


'/**
' * Defines image data and characteristics
' * */
Private Type opj_image
    '/** XOsiz: horizontal offset from the origin of the reference grid to the left side of the image area */
    x0 As Long
    '/** YOsiz: vertical offset from the origin of the reference grid to the top side of the image area */
    y0 As Long
    '/** Xsiz: width of the reference grid */
    x1 As Long
    '/** Ysiz: height of the reference grid */
    y1 As Long
    '/** number of components in the image */
    numcomps As Long
    '/** color space: sRGB, Greyscale or YUV */
    color_space As OPJ_COLOR_SPACE
    '/** image components */
    pComps As Long
    '/** 'restricted' ICC profile */
    pIccProfile As Long
    '/** size of ICC profile */
    icc_profile_len As Long
End Type

'OpenJPEG supports callbacks for messages, warnings, and errors, but these require cdecl functions.
' I use a twinbasic-built wrapper with delegates to handle these.
Private Declare Sub PD_GetAddrJP2KCallbacks Lib "PDHelper_win32" (ByVal dstInfoCallbackIn As Long, ByVal dstWarnCallbackIn As Long, ByVal dstErrCallbackIn As Long, ByRef dstInfoCallback As Long, ByRef dstWarnCallback As Long, ByRef dstErrCallback As Long)

'Official OpenJPEG builds use stdcall
Private Declare Function opj_version Lib "openjp2" Alias "_opj_version@0" () As Long
Private Declare Sub opj_set_default_decoder_parameters Lib "openjp2" Alias "_opj_set_default_decoder_parameters@4" (ByVal p_parameters As Long)
Private Declare Function opj_create_decompress Lib "openjp2" Alias "_opj_create_decompress@4" (ByVal OPJ_CODEC_FORMAT As Long) As Long
Private Declare Function opj_setup_decoder Lib "openjp2" Alias "_opj_setup_decoder@8" (ByVal p_codec As Long, ByVal p_parameters As Long) As Long
Private Declare Sub opj_set_info_handler Lib "openjp2" Alias "_opj_set_info_handler@12" (ByVal p_codec As Long, ByVal opj_msg_callback As Long, ByVal p_user_data As Long)
Private Declare Sub opj_set_warning_handler Lib "openjp2" Alias "_opj_set_warning_handler@12" (ByVal p_codec As Long, ByVal opj_msg_callback As Long, ByVal p_user_data As Long)
Private Declare Sub opj_set_error_handler Lib "openjp2" Alias "_opj_set_error_handler@12" (ByVal p_codec As Long, ByVal opj_msg_callback As Long, ByVal p_user_data As Long)
Private Declare Function opj_decoder_set_strict_mode Lib "openjp2" Alias "_opj_decoder_set_strict_mode@8" (ByVal p_codec As Long, ByVal strict As Long) As Long
Private Declare Sub opj_destroy_codec Lib "openjp2" Alias "_opj_destroy_codec@4" (ByVal p_codec As Long)
Private Declare Function opj_codec_set_threads Lib "openjp2" Alias "_opj_codec_set_threads@8" (ByVal p_codec As Long, ByVal num_threads As Long) As Long
Private Declare Function opj_stream_create_default_file_stream Lib "openjp2" Alias "_opj_stream_create_default_file_stream@8" (ByVal p_fname As Long, ByVal p_is_read_stream As Long) As Long
Private Declare Sub opj_stream_destroy Lib "openjp2" Alias "_opj_stream_destroy@4" (ByVal p_stream As Long)
Private Declare Function opj_read_header Lib "openjp2" Alias "_opj_read_header@12" (ByVal p_stream As Long, ByVal p_codec As Long, ByVal pp_image As Long) As Long
Private Declare Sub opj_image_destroy Lib "openjp2" Alias "_opj_image_destroy@4" (ByVal p_image As Long)
Private Declare Function opj_decode Lib "openjp2" Alias "_opj_decode@12" (ByVal p_decompressor As Long, ByVal p_stream As Long, ByVal p_image As Long) As Long
Private Declare Function opj_end_decompress Lib "openjp2" Alias "_opj_end_decompress@8" (ByVal p_codec As Long, ByVal p_stream As Long) As Long

/**
 * Creates an abstract stream. This function does nothing except allocating memory and initializing the abstract stream.
 *
 * @param   p_is_input      if set to true then the stream will be an input stream, an output stream else.
 *
 * @return  a stream object.
*/
OPJ_API opj_stream_t* OPJ_CALLCONV opj_stream_default_create(
    OPJ_BOOL p_is_input);

/**
 * Creates an abstract stream. This function does nothing except allocating memory and initializing the abstract stream.
 *
 * @param   p_buffer_size  FIXME DOC
 * @param   p_is_input      if set to true then the stream will be an input stream, an output stream else.
 *
 * @return  a stream object.
*/
OPJ_API opj_stream_t* OPJ_CALLCONV opj_stream_create(OPJ_SIZE_T p_buffer_size,
        OPJ_BOOL p_is_input);

/**
 * Destroys a stream created by opj_create_stream. This function does NOT close the abstract stream. If needed the user must
 * close its own implementation of the stream.
 *
 * @param   p_stream    the stream to destroy.
 */
OPJ_API void OPJ_CALLCONV opj_stream_destroy(opj_stream_t* p_stream);

/**
 * Sets the given function to be used as a read function.
 * @param       p_stream    the stream to modify
 * @param       p_function  the function to use a read function.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_read_function(opj_stream_t* p_stream,
        opj_stream_read_fn p_function);

/**
 * Sets the given function to be used as a write function.
 * @param       p_stream    the stream to modify
 * @param       p_function  the function to use a write function.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_write_function(opj_stream_t* p_stream,
        opj_stream_write_fn p_function);

/**
 * Sets the given function to be used as a skip function.
 * @param       p_stream    the stream to modify
 * @param       p_function  the function to use a skip function.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_skip_function(opj_stream_t* p_stream,
        opj_stream_skip_fn p_function);

/**
 * Sets the given function to be used as a seek function, the stream is then seekable,
 * using SEEK_SET behavior.
 * @param       p_stream    the stream to modify
 * @param       p_function  the function to use a skip function.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_seek_function(opj_stream_t* p_stream,
        opj_stream_seek_fn p_function);

/**
 * Sets the given data to be used as a user data for the stream.
 * @param       p_stream    the stream to modify
 * @param       p_data      the data to set.
 * @param       p_function  the function to free p_data when opj_stream_destroy() is called.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_user_data(opj_stream_t* p_stream,
        void * p_data, opj_stream_free_user_data_fn p_function);

/**
 * Sets the length of the user data for the stream.
 *
 * @param p_stream    the stream to modify
 * @param data_length length of the user_data.
*/
OPJ_API void OPJ_CALLCONV opj_stream_set_user_data_length(
    opj_stream_t* p_stream, OPJ_UINT64 data_length);


'Current image, if any
Private m_j2kImage As opj_image

'What follows are PD-specific structs for importing J2K data
Private Type PD_OpjNotes
    finalWidth As Long
    finalHeight As Long
    numComponents As Long
    imgHasAlpha As Boolean
    idxAlphaChannel As Integer
    isNot8Bit As Boolean
    hasSubsampling As Boolean
    isChannelSubsampled() As Boolean
End Type

'PD-specific details re: the current image.  J2K images have a lot of storage details that are complicated for
' PD to handle (like different subsampling on each channel).  Passing those details between functions is made
' easier by module-level storage.
Private m_OpjNotes As PD_OpjNotes
    
'Forcibly disable CharLS interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String

    If (m_LibHandle <> 0) And m_LibAvailable Then
        Dim ptrVersion As Long
        ptrVersion = opj_version()
        GetVersion = Strings.StringFromCharPtr(ptrVersion, False)
    End If
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim strLibPath As String
    strLibPath = pathToDLLFolder & "openjp2.dll"
    m_LibHandle = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If (Not InitializeEngine) Then PDDebug.LogAction "WARNING!  LoadLibraryW failed to load OpenJPEG.  Last DLL error: " & Err.LastDllError
    
End Function

Public Function IsOpenJPEGEnabled() As Boolean
    IsOpenJPEGEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Private Sub InternalError(Optional ByRef funcName As String = vbNullString, Optional ByRef errString As String = vbNullString)
    PDDebug.LogAction "OpenJPEG error in " & funcName & "(): " & errString, PDM_External_Lib
End Sub

'**********************************************
' / END GENERIC 3RD-PARTY LIBRARY BOILERPLATE
'**********************************************

'Verify J2K file signature.  Doesn't require OpenJPEG.
Public Function IsFileJ2K(ByRef srcFile As String, Optional ByRef outCodecFormat As OPJ_CODEC_FORMAT) As Boolean

    Const FUNC_NAME As String = "IsFileJ2K"
    IsFileJ2K = False
    
    If Files.FileExists(srcFile) Then
        
        'Some format variations can be determined by extension only; this logic is taken from
        ' the OpenJPEG project's official implementation: https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c
        Dim srcExtension As String
        srcExtension = Files.FileGetExtension(srcFile)
        
        If Strings.StringsEqual(srcExtension, "jpt", True) Then
            outCodecFormat = OPJ_CODEC_JPT
            IsFileJ2K = True
            If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JPT stream"
            Exit Function
        End If
        
        'Pull the first 12 bytes only
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
            Dim bFirst12() As Byte
            If cStream.ReadBytes(bFirst12, 12, True) Then
                
                'Various different signatures are valid, based on the container used.
                ' (Magic numbers taken from https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c#L532)
                '#define JP2_RFC3745_MAGIC "\x00\x00\x00\x0c\x6a\x50\x20\x20\x0d\x0a\x87\x0a"
                '#define JP2_MAGIC "\x0d\x0a\x87\x0a"
                '#define J2K_CODESTREAM_MAGIC "\xff\x4f\xff\x51"
                Const JP2_RFC3745_MAGIC_1 As Long = &HC000000
                Const JP2_RFC3745_MAGIC_2 As Long = &H2020506A
                Const JP2_RFC3745_MAGIC_3 As Long = &HA870A0D
                 
                If VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_RFC3745_MAGIC_1), 4) Then
                    If VBHacks.MemCmp(VarPtr(bFirst12(4)), VarPtr(JP2_RFC3745_MAGIC_2), 4) Then
                        IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(8)), VarPtr(JP2_RFC3745_MAGIC_3), 4)
                        outCodecFormat = OPJ_CODEC_JP2
                        If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2_RFC3745"
                    End If
                End If
                
                If (Not IsFileJ2K) Then
                    
                    Const JP2_MAGIC As Long = &HA870A0D
                    Const J2K_CODESTREAM_MAGIC As Long = &H51FF4FFF
                    
                    IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_MAGIC), 4)
                    outCodecFormat = OPJ_CODEC_JP2
                    If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2 stream"
                    If (Not IsFileJ2K) Then
                        IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(J2K_CODESTREAM_MAGIC), 4)
                        outCodecFormat = OPJ_CODEC_J2K
                        If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is J2K stream"
                    End If
                    
                End If
                
            End If
        End If
        
        Set cStream = Nothing
        
    End If
    
End Function

'Load a J2K image.  Will validate the file prior to loading.  Requires OpenJPEG.
Public Function LoadJ2K(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "LoadJ2K"
    LoadJ2K = False
    
    'Failsafe check; this function is pointless if OpenJPEG doesn't exist
    If (Not Plugin_OpenJPEG.IsOpenJPEGEnabled()) Then Exit Function
    
    'Failsafe check; validate file signature (hopefully the caller did this, but you never know)
    Dim srcCodec As OPJ_CODEC_FORMAT
    If (Not Plugin_OpenJPEG.IsFileJ2K(srcFile, srcCodec)) Then Exit Function
    
    'Still here?  This file passed basic validation.
    
    'Initialize a default J2K decoder
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder itself..."
    Dim pDecoder As Long
    pDecoder = opj_create_decompress(srcCodec)
    If (pDecoder = 0) Then
        InternalError FUNC_NAME, "opj_create_Decompress failed"
        Exit Function
    End If
    
    'Initialize function wrappers for our constructed decoder.
    ' (We do this via a twinBasic-built helper library, which allows for CDecl callbacks
    '  without needing embedded assembly shenanigans.)
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing callbacks..."
    Dim hInfo As Long, hWarning As Long, hError As Long
    PD_GetAddrJP2KCallbacks AddressOf HandlerInfo, AddressOf HandlerWarning, AddressOf HandlerError, hInfo, hWarning, hError
    
    opj_set_info_handler pDecoder, hInfo, 0&
    opj_set_warning_handler pDecoder, hWarning, 0&
    opj_set_error_handler pDecoder, hError, 0&
    
    'Decoders support variable behavior via a "decoder parameter" struct.
    ' Populate a parameter struct with default values.
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    'If you want to set custom decoding parameters, do it here.
    ' (For now, PD uses default decoding params.)
    
    'Load our decoder with whatever decoding parameters we've decided on
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    Dim retOJ As Long
    retOJ = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOJ = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    'Decoders can use a "strict" mode, where incomplete streams are disallowed.
    ' (Non-strict mode tells the decoder to simply decode as much as they can, and stop when they
    '  run out of room - this may allow *some* files to be partially recovered.)
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "setting strict mode to OFF..."
    If (opj_decoder_set_strict_mode(pDecoder, 0&) <> 1&) Then
        InternalError FUNC_NAME, "failed to set strictness mode"
        GoTo SafeCleanup
    End If
    
    'Allow the decoder to use as many logical threads as it wants
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "allowing multithreaded decode..."
    If (opj_codec_set_threads(pDecoder, OS.LogicalCoreCount()) = 1&) Then
        If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Allowing J2K decoder to use " & OS.LogicalCoreCount() & " cores"
    Else
        InternalError FUNC_NAME, "couldn't set thread count; single-thread mode will be used"
    End If
    
    'Prep a generic reader against the target file.
    ' (Note that the decoder parameters struct has a place for filename, but that is only used
    '  *internally* by the library and is of no use here.)
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    'Convert the source filename to UTF-8 before assigning
    Dim utf8Filename() As Byte
    Strings.UTF8FromString srcFile, utf8Filename, 0&
    
    Dim pStream As Long
    pStream = opj_stream_create_default_file_stream(VarPtr(utf8Filename(0)), 1&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't create stream on target file; load abandoned"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Reading header..."
    
    Dim pImage As Long
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Header read successfully"
    VBHacks.CopyMemoryStrict VarPtr(m_j2kImage), pImage, LenB(m_j2kImage)
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction m_j2kImage.x0 & ", " & m_j2kImage.y0 & ", " & m_j2kImage.x1 & ", " & m_j2kImage.y1 & ", " & m_j2kImage.numcomps & " " & GetNameOfOpjColorSpace(m_j2kImage.color_space) & " components"
    
    'Finish decoding the rest of the image
    If (opj_decode(pDecoder, pStream, pImage) <> 1&) Then
        InternalError FUNC_NAME, "failed to decode image"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Decoded full image successfully"
    
    'Read to the end of the file to pull any other useful info (e.g. metadata)
    If (opj_end_decompress(pDecoder, pStream) <> 1&) Then
        InternalError FUNC_NAME, "failed to read to end of file"
        'Attempt to load image anyway
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Reached EOF successfully"
    
    'We can now pull channel data from the supplied image pointer.
    
    Dim numComponents As Long
    numComponents = m_j2kImage.numcomps
    
    If (numComponents <= 0) Then
        PDDebug.LogAction "Invalid channel count: " & numComponents
        GoTo SafeCleanup
    End If
    
    Dim imgChannels() As opj_image_comp
    ReDim imgChannels(0 To numComponents - 1) As opj_image_comp
    
    Dim sizeOfChannel As Long, sizeOfChannelAligned As Long
    sizeOfChannelAligned = LenB(imgChannels(0))
    sizeOfChannel = Len(imgChannels(0))
    
    'To simplify reading data from an arbitrary pointer, construct a pdStream object against it.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=sizeOfChannelAligned * numComponents, baseFilePointerOffset:=m_j2kImage.pComps, optimizeAccess:=OptimizeSequentialAccess) Then
        
        Dim i As Long
        For i = 0 To numComponents - 1
            cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
            VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), cStream.ReadBytes_PointerOnly(sizeOfChannel), sizeOfChannel
            
            If J2K_DEBUG_VERBOSE Then
                PDDebug.LogAction "Channel #" & CStr(i + 1) & " info: "
                With imgChannels(i)
                    PDDebug.LogAction .x0 & ", " & .y0 & ", " & .w & ", " & .h & ", " & .prec & ", " & .Alpha
                    PDDebug.LogAction .p_data & ", " & .dx & ", " & .dy & ", " & .factor & ", " & .sgnd
                End With
            End If
        
        Next i
        
    Else
        PDDebug.LogAction "Failed to initialize stream against component pointer."
        GoTo SafeCleanup
    End If
    
    'We are done with our stream object
    cStream.StopStream True
    
    'With channel headers assembled, we now need to iterate channels and copy their contents into a dedicated pdDIB object.
    
    'Prep the target object.  For size, we ideally want to use the size in the header, which may not match the
    ' size of all/any individual components.  (J2K images are challenging this way.)
    '
    'For this initial draft, we are going to do a few different validations.
    '
    'First, we're going to figure out what color space to use for the embedded data.
    Dim targetColorSpace As OPJ_COLOR_SPACE
    targetColorSpace = DetermineColorHandling(m_j2kImage.color_space, numComponents, imgChannels)
    If (targetColorSpace <> OPJ_CLRSPC_GRAY) And (targetColorSpace <> OPJ_CLRSPC_SRGB) And (targetColorSpace <> OPJ_CLRSPC_SYCC) Then
        PDDebug.LogAction "Unknown color space or component count.  Abandoning load."
        GoTo SafeCleanup
    End If
    
    'TODO: handle YCC variations; they're valid and not too bad to handle (theoretically)
    
    If m_OpjNotes.hasSubsampling Then
        PDDebug.LogAction "NOTE: this image uses subsampling; PD can't load it (yet)."
        GoTo SafeCleanup
    End If
    
    'Non-8-bpp color depth handling is TODO
    If m_OpjNotes.isNot8Bit Then
        PDDebug.LogAction "Only 8-bit-per-channel J2K images are currently supported"
        GoTo SafeCleanup
    End If
    
    'With channels and color space successfully flagged, we can now load the image.
    
    'Prep the destination image
    Set dstDIB = New pdDIB
    dstDIB.CreateBlank m_OpjNotes.finalWidth, m_OpjNotes.finalHeight, 32, RGB(255, 255, 255), 255
    
    'Load pixel data, with handling separated by color type
    If (targetColorSpace = OPJ_CLRSPC_GRAY) Then
    
    ElseIf (targetColorSpace = OPJ_CLRSPC_SRGB) Or (targetColorSpace = OPJ_CLRSPC_SYCC) Then
        
        Dim targetWidth As Long, targetHeight As Long
        Dim channelSizeEstimate As Long
        Dim srcRs() As Long, srcGs() As Long, srcBs() As Long, srcAs() As Long
        Dim srcRSA As SafeArray1D, srcGSA As SafeArray1D, srcBSA As SafeArray1D, srcASA As SafeArray1D
        Dim copyAlpha As Boolean
                
        Dim r As Long, g As Long, b As Long, yccY As Long, yccB As Long, yccR As Long
        Dim x As Long, y As Long
        Dim dstPixels() As Byte, dstSA As SafeArray1D
        Dim saOffset As Long
        
        If (imgChannels(0).prec = 8) Then
            
            'Size estimate obviously needs to change by bit-depth TODO
            targetWidth = m_OpjNotes.finalWidth
            targetHeight = m_OpjNotes.finalHeight
            
            channelSizeEstimate = targetWidth * targetHeight * 4
            
            VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(1).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcBs, srcBSA, imgChannels(2).p_data, channelSizeEstimate
            
            copyAlpha = m_OpjNotes.imgHasAlpha
            If copyAlpha Then VBHacks.WrapArrayAroundPtr_Long srcAs, srcASA, imgChannels(3).p_data, channelSizeEstimate
            Debug.Print targetWidth, targetHeight, channelSizeEstimate
            For y = 0 To targetHeight - 1
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth '* 4
                'Debug.Print saOffset
            If (targetColorSpace = OPJ_CLRSPC_SRGB) Then
                For x = 0 To targetWidth - 1
                    dstPixels(x * 4) = srcBs(saOffset + x)
                    dstPixels(x * 4 + 1) = srcGs(saOffset + x)
                    dstPixels(x * 4 + 2) = srcRs(saOffset + x)
                    If copyAlpha Then dstPixels(x * 4 + 3) = srcAs(saOffset + x)
                Next x
            
            'YCC to RGB conversion taken from OpenJPEG itself: https://github.com/uclouvain/openjpeg/blob/e7453e398b110891778d8da19209792c69ca7169/src/bin/common/color.c#L74
            ElseIf (targetColorSpace = OPJ_CLRSPC_SYCC) Then
                For x = 0 To targetWidth - 1
                    yccY = srcRs(saOffset + x)
                    yccB = srcGs(saOffset + x) - 127
                    yccR = srcBs(saOffset + x) - 127
                    r = yccY + 1.402 * yccR
                    If (r < 0) Then r = 0
                    If (r > 255) Then r = 255
                    g = yccY - (0.344 * yccB + 0.714 * yccR)
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                    b = yccY + (1.772 * yccB)
                    If (b < 0) Then b = 0
                    If (b > 255) Then b = 255
                    dstPixels(x * 4) = b
                    dstPixels(x * 4 + 1) = g
                    dstPixels(x * 4 + 2) = r
                    If copyAlpha Then dstPixels(x * 4 + 3) = srcAs(saOffset + x)
                Next x
            End If
            Next y
            
        ElseIf (imgChannels(0).prec = 16) Then
        
            'Size estimate obviously needs to change by bit-depth TODO
            targetWidth = m_OpjNotes.finalWidth
            targetHeight = m_OpjNotes.finalHeight
            
            channelSizeEstimate = targetWidth * targetHeight * 4
            
            VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(1).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcBs, srcBSA, imgChannels(2).p_data, channelSizeEstimate
            
            copyAlpha = m_OpjNotes.imgHasAlpha
            If copyAlpha Then VBHacks.WrapArrayAroundPtr_Long srcAs, srcASA, imgChannels(3).p_data, channelSizeEstimate
            
            For y = 0 To targetHeight - 1
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth
            For x = 0 To targetWidth - 1
                dstPixels(x * 4) = srcBs(saOffset + x) \ 256
                dstPixels(x * 4 + 1) = srcGs(saOffset + x) \ 256
                dstPixels(x * 4 + 2) = srcRs(saOffset + x) \ 256
                If copyAlpha Then dstPixels(x * 4 + 3) = srcAs(saOffset + x) \ 256
            Next x
            Next y
            
        End If
        
        'Unwrap all temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        VBHacks.UnwrapArrayFromPtr_Long srcGs
        VBHacks.UnwrapArrayFromPtr_Long srcBs
        If copyAlpha Then VBHacks.UnwrapArrayFromPtr_Long srcAs
        dstDIB.UnwrapArrayFromDIB dstPixels
        dstDIB.SetAlphaPremultiplication True
        
        LoadJ2K = True
        
    End If
    
    'For now cleanup and exit
    GoTo SafeCleanup
    
    Exit Function
    
SafeCleanup:
    If (Not cStream Is Nothing) Then cStream.StopStream True
    If (pImage <> 0) Then opj_image_destroy pImage
    If (pStream <> 0) Then opj_stream_destroy pStream
    If (pDecoder <> 0) Then opj_destroy_codec pDecoder
    
End Function

'Figure out how to handle the source color data.  JPEG-2000 streams are extremely flexible in terms of color components
' (e.g. "undefined" color spaces and infinite color component counts are allowed, each with different dimensions).
' I don't currently intend to cover every possible combination of file attributes.  Instead, I want PD to make intelligent
' inferences about unknown data (e.g. three undefined channels are likely RGB - this is how other software handles it).
'
'If an obvious correlation with a known color space cannot be made, PD will treat the image data as grayscale and load
' the first channel only.
Private Function DetermineColorHandling(ByVal fileColorSpace As OPJ_COLOR_SPACE, ByVal numComponents As Long, ByRef imgChannels() As opj_image_comp) As OPJ_COLOR_SPACE
    
    'An "unknown" color space notifies the caller that PD is unequipped to handle this image's data.
    DetermineColorHandling = OPJ_CLRSPC_UNKNOWN
    
    'Failsafe check for component count
    If (numComponents <= 0) Then Exit Function
    
    'Check the size of the first component.  We will only deal with subsequent components if their size matches these.
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = imgChannels(0).w
    targetHeight = imgChannels(0).h
    
    'Failsafe check for component dimensions
    If (targetWidth <= 0) Or (targetHeight <= 0) Then Exit Function
    
    'Reset the module-level property tracker
    With m_OpjNotes
        .finalWidth = targetWidth
        .finalHeight = targetHeight
        .hasSubsampling = False
        .imgHasAlpha = False
        .isNot8Bit = False
        ReDim .isChannelSubsampled(0 To numComponents) As Boolean
        .idxAlphaChannel = -1
    End With
    
    'If there is only one channel in the image, color space doesn't matter - treat it as grayscale.
    If (numComponents = 1) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        With m_OpjNotes
            .hasSubsampling = False
            .imgHasAlpha = False
            .idxAlphaChannel = -1
            .isChannelSubsampled(0) = False
            .isNot8Bit = (imgChannels(0).prec <> 8)
        End With
        Exit Function
    End If
    
    'If we're still here, this image has multiple channels.  Iterate up to 4 channels and count how many
    ' have dimensions matching the first component.  (Subsampling in J2K files means each channel can have its
    ' own independent dimensions, and the caller is expected to scale all components to match in the file image.
    ' I don't currently have a straightforward way to do that, so I'm going to follow the example of most other
    ' software and simply not deal with it.)
    '
    '(TODO in the future, perhaps: add support for handling subsampled channels.)
    Dim searchDepth As Long
    searchDepth = PDMath.Min2Int(numComponents, 4)
    
    Dim i As Long
    For i = 1 To searchDepth - 1
        
        'For now, stop if we meet a subsampled channel
        If (imgChannels(i).w <> targetWidth) Or (imgChannels(i).h <> targetHeight) Then
            m_OpjNotes.hasSubsampling = True
            Exit For
        End If
        
        'Note the alpha channel index, if any.
        ' (NOTE: this implementation is poorly tested; what if a non-3rd channel is marked as alpha?
        '        is this even possible or handle-able reliably?)
        If (imgChannels(i).Alpha <> 0) Then
            m_OpjNotes.idxAlphaChannel = i
            m_OpjNotes.imgHasAlpha = True
        ElseIf (i = 3) And (m_OpjNotes.idxAlphaChannel < 0) Then
            m_OpjNotes.idxAlphaChannel = i
            m_OpjNotes.imgHasAlpha = True
        End If
        
    Next i
    
    numComponents = i
    If (numComponents < 3) Then numComponents = 1
    
    'Assign a correct color space based on channel count
    If (numComponents > 0) And (numComponents < 3) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        m_OpjNotes.imgHasAlpha = (m_OpjNotes.idxAlphaChannel >= 0)
        If (Not m_OpjNotes.imgHasAlpha) Then m_OpjNotes.numComponents = 1
    ElseIf (numComponents > 3) Then
        DetermineColorHandling = OPJ_CLRSPC_SRGB
        m_OpjNotes.imgHasAlpha = (m_OpjNotes.idxAlphaChannel >= 0)
        If m_OpjNotes.imgHasAlpha Then
            numComponents = 4
        Else
            numComponents = 3
        End If
    Else
        
        If (fileColorSpace = OPJ_CLRSPC_EYCC) Then
            DetermineColorHandling = OPJ_CLRSPC_EYCC
        ElseIf (fileColorSpace = OPJ_CLRSPC_SYCC) Then
            DetermineColorHandling = OPJ_CLRSPC_SYCC
        Else
            DetermineColorHandling = OPJ_CLRSPC_SRGB
        End If
        
        'Alpha handling is handled identically here; basically, we only mark an image as having alpha
        ' if a channel is *specifically* flagged as an alpha channel (and is one of the first 4 channels).
        m_OpjNotes.imgHasAlpha = (m_OpjNotes.idxAlphaChannel >= 0)
        If m_OpjNotes.imgHasAlpha Then
            numComponents = 4
        Else
            numComponents = 3
        End If
        
    End If
    
    'TODO: identify subsampling and flag channels accordingly.
    
End Function

Private Function GetNameOfOpjColorSpace(ByVal srcSpace As OPJ_COLOR_SPACE) As String
    
    Select Case srcSpace
        
        '/**< not supported by the library */
        Case OPJ_CLRSPC_UNKNOWN
            GetNameOfOpjColorSpace = "unknown"
        '/**< not specified in the codestream */
        Case OPJ_CLRSPC_UNSPECIFIED
            GetNameOfOpjColorSpace = "unspecified"
        '/**< sRGB */
        Case OPJ_CLRSPC_SRGB
            GetNameOfOpjColorSpace = "sRGB"
        '/**< grayscale */
        Case OPJ_CLRSPC_GRAY
            GetNameOfOpjColorSpace = "grayscale"
        '/**< YUV */
        Case OPJ_CLRSPC_SYCC
            GetNameOfOpjColorSpace = "YUV"
        '/**< e-YCC */
        Case OPJ_CLRSPC_EYCC
            GetNameOfOpjColorSpace = "e-YCC"
    End Select
    
End Function

Private Sub HandlerInfo(ByVal pMsg As Long, ByVal pUserData As Long)
    PDDebug.LogAction "openJPEG Info: " & Strings.StringFromCharPtr(pMsg, False), PDM_External_Lib
End Sub

Private Sub HandlerWarning(ByVal pMsg As Long, ByVal pUserData As Long)
    PDDebug.LogAction "openJPEG Warning: " & Strings.StringFromCharPtr(pMsg, False), PDM_External_Lib
End Sub

Private Sub HandlerError(ByVal pMsg As Long, ByVal pUserData As Long)
    PDDebug.LogAction "openJPEG Error: " & Strings.StringFromCharPtr(pMsg, False), PDM_External_Lib
End Sub

