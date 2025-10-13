Attribute VB_Name = "Plugin_OpenJPEG"
'***************************************************************************
'OpenJPEG (JPEG-2000) Library Interface
'Copyright 2025-2025 by Tanner Helland
'Created: 19/September/25
'Last updated: 13/October/25
'Last update: add import support for a bunch of esoteric JP2 formats
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
'PhotoDemon originally used OpenJPEG via FreeImage, but FreeImage has been abandoned so a new solution
' was needed.  So in 2025 I wrote a new, direct OpenJPEG interface from scratch.
'
'This interface was designed against v2.5.4 of the library (released 20 Sep 2025).  It should work fine
' with any version that maintains ABI compatibility with that release.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Enable at DEBUG-TIME for verbose logging
Private Const JP2_DEBUG_VERBOSE As Boolean = False

'To strictly enforce the spec (and decrease chances of OpenJPEG crashes on malformed images), set this to TRUE.
' I currently set it to FALSE in production builds, to allow many more "in the wild" images to actually load.
Private Const JP2_FORCE_STRICT_DECODING As Boolean = False

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
Private Declare Sub PD_GetAddrJP2FileCallbacks Lib "PDHelper_win32" (ByVal dstCbRead As Long, ByVal dstCbWrite As Long, ByRef dstCbSeek As Long, ByVal dstCbSkip As Long, ByRef outCbRead As Long, ByRef outCbWrite As Long, ByRef outCbSeek As Long, ByRef outCbSkip As Long)

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
Private Declare Function opj_read_header Lib "openjp2" Alias "_opj_read_header@12" (ByVal p_stream As Long, ByVal p_codec As Long, ByVal pp_image As Long) As Long
Private Declare Sub opj_image_destroy Lib "openjp2" Alias "_opj_image_destroy@4" (ByVal p_image As Long)
Private Declare Function opj_decode Lib "openjp2" Alias "_opj_decode@12" (ByVal p_decompressor As Long, ByVal p_stream As Long, ByVal p_image As Long) As Long
Private Declare Function opj_end_decompress Lib "openjp2" Alias "_opj_end_decompress@8" (ByVal p_codec As Long, ByVal p_stream As Long) As Long

Private Declare Function opj_stream_create_default_file_stream Lib "openjp2" Alias "_opj_stream_create_default_file_stream@8" (ByVal p_fname As Long, ByVal p_is_read_stream As Long) As Long
Private Declare Function opj_stream_default_create Lib "openjp2" Alias "_opj_stream_default_create@4" (ByVal bool_p_is_input As Long) As Long
Private Declare Function opj_stream_create Lib "openjp2" Alias "_opj_stream_create@8" (ByVal p_buffer_size As Long, ByVal bool_p_is_input As Long) As Long
Private Declare Sub opj_stream_destroy Lib "openjp2" Alias "_opj_stream_destroy@4" (ByVal p_stream As Long)

Private Declare Sub opj_stream_set_read_function Lib "openjp2" Alias "_opj_stream_set_read_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_write_function Lib "openjp2" Alias "_opj_stream_set_write_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_skip_function Lib "openjp2" Alias "_opj_stream_set_skip_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_seek_function Lib "openjp2" Alias "_opj_stream_set_seek_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_user_data Lib "openjp2" Alias "_opj_stream_set_user_data@12" (ByVal p_stream As Long, ByVal p_data As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_user_data_length Lib "openjp2" Alias "_opj_stream_set_user_data_length@12" (ByVal p_stream As Long, ByVal data_length As Currency)

'Current image, if any
Private m_jp2Image As opj_image

'What follows are PD-specific structs for importing JP2 data
Private Type PD_OpjNotes
    finalWidth As Long
    finalHeight As Long
    numComponents As Long
    imgHasAlpha As Boolean
    idxAlphaChannel As Integer
    isAtLeast8Bit As Boolean
    hasSubsampling As Boolean
    isChannelSubsampled() As Boolean
    channelSsWidth() As Long        'Subsampled width/height, in pixels, of channel at index [n]
    channelSsHeight() As Long
End Type

'PD-specific details re: the current image.  JP2 images have a lot of storage details that are complicated for
' PD to handle (like different subsampling on each channel).  Passing those details between functions is made
' easier by module-level storage.
Private m_OpjNotes As PD_OpjNotes

'This stream reads/write the actual JP2 data
Private m_Stream As pdStream
    
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

'Verify JPEG-2000 file signature.  Doesn't require OpenJPEG.
Public Function IsFileJP2(ByRef srcFile As String, Optional ByRef outCodecFormat As OPJ_CODEC_FORMAT) As Boolean

    Const FUNC_NAME As String = "IsFileJP2"
    IsFileJP2 = False
    
    If Files.FileExists(srcFile) Then
        
        'Some format variations can be determined by extension only; this logic is taken from
        ' the OpenJPEG project's official implementation: https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c
        Dim srcExtension As String
        srcExtension = Files.FileGetExtension(srcFile)
        
        If Strings.StringsEqual(srcExtension, "jpt", True) Then
            outCodecFormat = OPJ_CODEC_JPT
            IsFileJP2 = True
            If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "File is JPT stream"
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
                        IsFileJP2 = VBHacks.MemCmp(VarPtr(bFirst12(8)), VarPtr(JP2_RFC3745_MAGIC_3), 4)
                        outCodecFormat = OPJ_CODEC_JP2
                        If IsFileJP2 And JP2_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2_RFC3745"
                    End If
                End If
                
                If (Not IsFileJP2) Then
                    
                    Const JP2_MAGIC As Long = &HA870A0D
                    Const J2K_CODESTREAM_MAGIC As Long = &H51FF4FFF
                    
                    IsFileJP2 = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_MAGIC), 4)
                    outCodecFormat = OPJ_CODEC_JP2
                    If IsFileJP2 And JP2_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2 stream"
                    If (Not IsFileJP2) Then
                        IsFileJP2 = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(J2K_CODESTREAM_MAGIC), 4)
                        outCodecFormat = OPJ_CODEC_J2K
                        If IsFileJP2 And JP2_DEBUG_VERBOSE Then PDDebug.LogAction "File is J2K stream"
                    End If
                    
                End If
                
            End If
        End If
        
        Set cStream = Nothing
        
    End If
    
End Function

'Load a JPEG-2000 image.  Will validate the file prior to loading.  Requires OpenJPEG.
Public Function LoadJP2(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "LoadJP2"
    LoadJP2 = False
    
    'Failsafe check; this function is pointless if OpenJPEG doesn't exist
    If (Not Plugin_OpenJPEG.IsOpenJPEGEnabled()) Then Exit Function
    
    'Failsafe check; validate file signature (hopefully the caller did this, but you never know)
    Dim srcCodec As OPJ_CODEC_FORMAT
    If (Not Plugin_OpenJPEG.IsFileJP2(srcFile, srcCodec)) Then Exit Function
    
    'Still here?  This file passed basic validation.
    
    'Initialize a default JPEG-2000 decoder
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder itself..."
    Dim pDecoder As Long
    pDecoder = opj_create_decompress(srcCodec)
    If (pDecoder = 0) Then
        InternalError FUNC_NAME, "opj_create_Decompress failed"
        Exit Function
    End If
    
    'Initialize function wrappers for our constructed decoder.
    ' (We do this via a twinBasic-built helper library, which allows for CDecl callbacks
    '  without needing embedded assembly shenanigans.)
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing callbacks..."
    Dim hInfo As Long, hWarning As Long, hError As Long
    PD_GetAddrJP2KCallbacks AddressOf HandlerInfo, AddressOf HandlerWarning, AddressOf HandlerError, hInfo, hWarning, hError
    
    opj_set_info_handler pDecoder, hInfo, 0&
    opj_set_warning_handler pDecoder, hWarning, 0&
    opj_set_error_handler pDecoder, hError, 0&
    
    'Decoders support variable behavior via a "decoder parameter" struct.
    ' Populate a parameter struct with default values.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    'If you want to set custom decoding parameters, do it here.
    ' (For now, PD uses default decoding params.)
    
    'Load our decoder with whatever decoding parameters we've decided on
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    Dim retOJ As Long
    retOJ = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOJ = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    'Decoders can use a "strict" mode, where incomplete streams are disallowed.
    ' (Non-strict mode tells the decoder to simply decode as much as they can, and stop when they
    '  run out of room - this may allow *some* files to be partially recovered.)
    Dim strictModeValue As Long
    If JP2_FORCE_STRICT_DECODING Then strictModeValue = 1& Else strictModeValue = 0&
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "setting strict mode to " & UCase$(CStr(JP2_FORCE_STRICT_DECODING)) & "..."
    If (opj_decoder_set_strict_mode(pDecoder, strictModeValue) <> 1&) Then
        InternalError FUNC_NAME, "failed to set strictness mode"
        GoTo SafeCleanup
    End If
    
    'Allow the decoder to use as many logical threads as it wants
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "allowing multithreaded decode..."
    If (opj_codec_set_threads(pDecoder, OS.LogicalCoreCount()) = 1&) Then
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Allowing OpenJPEG to use " & OS.LogicalCoreCount() & " cores"
    Else
        InternalError FUNC_NAME, "couldn't set thread count; single-thread mode will be used"
    End If
    
    'Prep a generic reader against the target file.
    ' (Note that the decoder parameters struct has a place for filename, but that is only used
    '  *internally* by the library and is of no use here.)
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    'TODO: use our I/O functions instead of OpenJPEG's (because theirs don't support Unicode on Windows)
    Set m_Stream = New pdStream
    If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        InternalError FUNC_NAME, "couldn't start pdStream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "pdStream initialized OK..."
    
    Dim pStream As Long
    pStream = opj_stream_default_create(1)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't start jp2 stream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Blank jp2 stream initialized OK..."
    
    Dim hRead As Long, hWrite As Long, hSkip As Long, hSeek As Long
    PD_GetAddrJP2FileCallbacks AddressOf JP2_ReadProcDelegate, AddressOf JP2_WriteProcDelegate, AddressOf JP2_SkipProcDelegate, AddressOf JP2_SeekProcDelegate, hRead, hWrite, hSeek, hSkip
    
    opj_stream_set_user_data pStream, 0&, 0&
    opj_stream_set_user_data_length pStream, Files.FileLenW(srcFile) \ 10000
    opj_stream_set_read_function pStream, hRead
    opj_stream_set_write_function pStream, hWrite
    opj_stream_set_skip_function pStream, hSkip
    opj_stream_set_seek_function pStream, hSeek
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Callbacks assigned OK..."
    
    Dim pImage As Long
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Header read successfully"
    VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), pImage, LenB(m_jp2Image)
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction m_jp2Image.x0 & ", " & m_jp2Image.y0 & ", " & m_jp2Image.x1 & ", " & m_jp2Image.y1 & ", " & m_jp2Image.numcomps & " " & GetNameOfOpjColorSpace(m_jp2Image.color_space) & " components"
    
    'Finish decoding the rest of the image
    If (opj_decode(pDecoder, pStream, pImage) <> 1&) Then
        InternalError FUNC_NAME, "failed to decode image"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Decoded full image successfully"
    
    'Read to the end of the file to pull any other useful info (e.g. metadata)
    If (opj_end_decompress(pDecoder, pStream) <> 1&) Then
        InternalError FUNC_NAME, "failed to read to end of file"
        'Attempt to load image anyway
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Reached EOF successfully"
    
    'We can now pull channel data from the supplied image pointer.
    
    Dim numComponents As Long
    numComponents = m_jp2Image.numcomps
    
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
    If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=sizeOfChannelAligned * numComponents, baseFilePointerOffset:=m_jp2Image.pComps, optimizeAccess:=OptimizeSequentialAccess) Then
        
        Dim i As Long
        For i = 0 To numComponents - 1
            cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
            VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), cStream.ReadBytes_PointerOnly(sizeOfChannel), sizeOfChannel
            
            If JP2_DEBUG_VERBOSE Then
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
    ' size of all/any individual components.  (JPEG-2000 images are challenging this way.)
    '
    'For this initial draft, we are going to do a few different validations.
    '
    'First, we're going to figure out what color space to use for the embedded data.
    Dim targetColorSpace As OPJ_COLOR_SPACE
    targetColorSpace = DetermineColorHandling(m_jp2Image.color_space, numComponents, imgChannels)
    If (targetColorSpace <> OPJ_CLRSPC_GRAY) And (targetColorSpace <> OPJ_CLRSPC_SRGB) And (targetColorSpace <> OPJ_CLRSPC_SYCC) Then
        PDDebug.LogAction "Unknown color space or component count.  Abandoning load."
        GoTo SafeCleanup
    End If
    
    'Subsampling requires us to use custom indices for each channel, based on its subsampling factor
    Dim subSamplingActive As Boolean
    subSamplingActive = m_OpjNotes.hasSubsampling
    
    Dim gSSfactorX As Single, bSSfactorX As Single, aSSfactorX As Single
    Dim gSSfactorY As Single, bSSfactorY As Single, aSSfactorY As Single
    Dim gWidth As Long, bWidth As Long, aWidth As Long
    
    If subSamplingActive And (numComponents > 1) Then
        
        'Calculate indices into each color channel.  (Note that we use RGBA indices, but channels may represent
        ' other color data - that's okay.)
        If (numComponents >= 2) Then
            gSSfactorX = CDbl(m_OpjNotes.channelSsWidth(1)) / CDbl(m_OpjNotes.finalWidth)
            gSSfactorY = CDbl(m_OpjNotes.channelSsHeight(1)) / CDbl(m_OpjNotes.finalHeight)
            gWidth = m_OpjNotes.channelSsWidth(1)
        End If
        
        If (numComponents >= 3) Then
            bSSfactorX = CDbl(m_OpjNotes.channelSsWidth(2)) / CDbl(m_OpjNotes.finalWidth)
            bSSfactorY = CDbl(m_OpjNotes.channelSsHeight(2)) / CDbl(m_OpjNotes.finalHeight)
            bWidth = m_OpjNotes.channelSsWidth(2)
        End If
        
        If (numComponents >= 4) Then
            aSSfactorX = CDbl(m_OpjNotes.channelSsWidth(3)) / CDbl(m_OpjNotes.finalWidth)
            aSSfactorY = CDbl(m_OpjNotes.channelSsHeight(3)) / CDbl(m_OpjNotes.finalHeight)
            aWidth = m_OpjNotes.channelSsWidth(3)
        End If
        
    End If
    
    'Non-8-bpp color depth handling is TODO
    If (Not m_OpjNotes.isAtLeast8Bit) Then
        PDDebug.LogAction "Only 8+ bit-per-channel JP2 images are currently supported"
        GoTo SafeCleanup
    End If
    
    'With channels and color space successfully flagged, we can now load the image.
    
    'Prep the destination image
    Set dstDIB = New pdDIB
    dstDIB.CreateBlank m_OpjNotes.finalWidth, m_OpjNotes.finalHeight, 32, RGB(255, 255, 255), 255

    Dim targetWidth As Long, targetHeight As Long
    Dim channelSizeEstimate As Long
    Dim srcRs() As Long, srcGs() As Long, srcBs() As Long, srcAs() As Long
    Dim srcRSA As SafeArray1D, srcGSA As SafeArray1D, srcBSA As SafeArray1D, srcASA As SafeArray1D
    Dim copyAlpha As Boolean
            
    Dim r As Long, g As Long, b As Long, a As Long, yccY As Long, yccB As Long, yccR As Long
    Dim x As Long, y As Long
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    Dim saOffset As Long, xOffset As Long, hdrDivisor As Long
    
    'Data can be signed, meaning that e.g. 8-bit data is represented as [-127, 128] instead of [0, 255]
    Dim rIsSigned As Boolean, gIsSigned As Boolean, bIsSigned As Boolean, aIsSigned As Boolean
    Dim rSgnComp As Long, gSgnComp As Long, bSgnComp As Long, aSgnComp As Long
    
    'Unlike other image format libraries, OpenJPEG always loads channels as ints (Longs in VB6) regardless of
    ' embedded color depth.  This is incredibly wasteful from a memory standpoint, but it does simplify
    ' handling of various bit-depths, because the source channel data is always the same size.
    targetWidth = m_OpjNotes.finalWidth
    targetHeight = m_OpjNotes.finalHeight
    channelSizeEstimate = targetWidth * targetHeight * 4    '(See above note - this line is not a typo!)
    
    'Load pixel data, with handling separated by color type
    If (targetColorSpace = OPJ_CLRSPC_GRAY) Then
        
        gIsSigned = (imgChannels(0).sgnd <> 0)
        If gIsSigned Then gSgnComp = (2 ^ imgChannels(0).prec) \ 2 Else gSgnComp = 0
        
        'Precision can technically be any value between 1 and ???? (upper limit is unclear from the spec).
        ' Note that data can also be *signed* which is currently a rare and untested state for PD handling.
        If (imgChannels(0).prec = 8) Then
            
            'Wrap a temporary array around the source channel
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(0).p_data, channelSizeEstimate
            
            'Iterate lines (data is stored top-down)
            For y = 0 To targetHeight - 1
            
                'Wrap a 1D array around the target line in in the destination image, and calculate an offset
                ' into the corresponding source channel line.
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth
                
                For x = 0 To targetWidth - 1
                    
                    g = srcGs(saOffset + x) + gSgnComp
                    
                    'Failsafe only, because some conformance test images have shown OOB gray colors.
                    ' (Well-formed images should never trigger these states.)
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                    
                    dstPixels(x * 4) = g
                    dstPixels(x * 4 + 1) = g
                    dstPixels(x * 4 + 2) = g
                    
                Next x
                
            'Proceed to next line
            Next y
            
        'In the future, we probably want to map this to a floating-point surface and propose tone-mapping
        ' (if ICC profile is missing).  For now however we perform a default linear map.
        ElseIf (imgChannels(0).prec > 8) Then
            
            'Wrap a temporary array around the source channel
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(0).p_data, channelSizeEstimate
            
            'For now, do a quick drop to 8-bit data (tone-mapping in the future is TBD)
            hdrDivisor = 2 ^ (imgChannels(0).prec - 8)
            
            'Iterate lines (data is stored top-down)
            For y = 0 To targetHeight - 1
            
                'Wrap a 1D array around the target line in in the destination image, and calculate an offset
                ' into the corresponding source channel line.
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth
                
                For x = 0 To targetWidth - 1
                    g = (srcGs(saOffset + x) + gSgnComp) \ hdrDivisor
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                    dstPixels(x * 4) = g
                    dstPixels(x * 4 + 1) = g
                    dstPixels(x * 4 + 2) = g
                Next x
                
            'Proceed to next line
            Next y
            
        End If
        
        'Unwrap all temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcGs
        dstDIB.UnwrapArrayFromDIB dstPixels
        dstDIB.SetAlphaPremultiplication True
        
        'Load complete!  (Clean-up is still required, however.)
        LoadJP2 = True
    
    ElseIf (targetColorSpace = OPJ_CLRSPC_SRGB) Or (targetColorSpace = OPJ_CLRSPC_SYCC) Then
        
        rIsSigned = (imgChannels(0).sgnd <> 0)
        gIsSigned = (imgChannels(1).sgnd <> 0)
        bIsSigned = (imgChannels(2).sgnd <> 0)
        If m_OpjNotes.imgHasAlpha Then aIsSigned = (imgChannels(3).sgnd <> 0) Else aIsSigned = False
        
        If rIsSigned Then rSgnComp = (2 ^ imgChannels(0).prec) \ 2 Else rSgnComp = 0
        If gIsSigned Then gSgnComp = (2 ^ imgChannels(1).prec) \ 2 Else gSgnComp = 0
        If bIsSigned Then bSgnComp = (2 ^ imgChannels(2).prec) \ 2 Else bSgnComp = 0
        If aIsSigned Then aSgnComp = (2 ^ imgChannels(3).prec) \ 2 Else aSgnComp = 0
        
        'Precision can technically be any value between 1 and ???? (upper limit is unclear from the spec).
        ' Note that data can also be *signed* which is currently a rare and untested state for PD handling.
        If (imgChannels(0).prec = 8) Then
            
            'Wrap temporary arrays around 3 or 4 source channels
            VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(1).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcBs, srcBSA, imgChannels(2).p_data, channelSizeEstimate
            
            copyAlpha = m_OpjNotes.imgHasAlpha
            If copyAlpha Then VBHacks.WrapArrayAroundPtr_Long srcAs, srcASA, imgChannels(3).p_data, channelSizeEstimate
            
            'Iterate lines (data is stored top-down)
            For y = 0 To targetHeight - 1
            
                'Wrap a 1D array around the target line in in the destination image, and calculate an offset
                ' into the corresponding source channel line.
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth
                
                'Further handling is separated by color type
                If (targetColorSpace = OPJ_CLRSPC_SRGB) Then
                    
                    'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                    ' we split handling accordingly.
                    If subSamplingActive Then
                        
                        For x = 0 To targetWidth - 1
                        
                            b = srcBs(Int(y * bSSfactorY) * bWidth + Int(x * bSSfactorX)) + bSgnComp
                            If (b < 0) Then b = 0
                            If (b > 255) Then b = 255
                            
                            g = srcGs(Int(y * gSSfactorY) * gWidth + Int(x * gSSfactorX)) + gSgnComp
                            If (g < 0) Then g = 0
                            If (g > 255) Then g = 255
                            
                            r = srcRs(saOffset + x) + rSgnComp
                            If (r < 0) Then r = 0
                            If (r > 255) Then r = 255
                    
                            dstPixels(x * 4) = b
                            dstPixels(x * 4 + 1) = g
                            dstPixels(x * 4 + 2) = r
                            
                            If copyAlpha Then
                                a = srcAs(Int(y * aSSfactorY) * aWidth + Int(x * aSSfactorX)) + aSgnComp
                                If (a < 0) Then a = 0
                                If (a > 255) Then a = 255
                                dstPixels(x * 4 + 3) = a
                            End If
                            
                        Next x
                        
                    Else
                    
                        For x = 0 To targetWidth - 1
                        
                            b = srcBs(saOffset + x) + bSgnComp
                            If (b < 0) Then b = 0
                            If (b > 255) Then b = 255
                            g = srcGs(saOffset + x) + gSgnComp
                            If (g < 0) Then g = 0
                            If (g > 255) Then g = 255
                            r = srcRs(saOffset + x) + rSgnComp
                            If (r < 0) Then r = 0
                            If (r > 255) Then r = 255
                            
                            dstPixels(x * 4) = b
                            dstPixels(x * 4 + 1) = g
                            dstPixels(x * 4 + 2) = r
                            
                            If copyAlpha Then
                                a = srcAs(saOffset + x) + aSgnComp
                                If (a < 0) Then a = 0
                                If (a > 255) Then a = 255
                                dstPixels(x * 4 + 3) = a
                            End If
                            
                        Next x
                        
                    End If
                
                'YCC to RGB conversion taken from OpenJPEG itself: https://github.com/uclouvain/openjpeg/blob/e7453e398b110891778d8da19209792c69ca7169/src/bin/common/color.c#L74
                ElseIf (targetColorSpace = OPJ_CLRSPC_SYCC) Then
                    
                    For x = 0 To targetWidth - 1
                        
                        'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                        ' we split handling accordingly.
                        If subSamplingActive Then
                            yccY = srcRs(saOffset + x) + rSgnComp
                            yccB = srcGs(Int(y * gSSfactorY) * gWidth + Int(x * gSSfactorX)) + gSgnComp - 127
                            yccR = srcBs(Int(y * bSSfactorY) * bWidth + Int(x * bSSfactorX)) + bSgnComp - 127
                            If copyAlpha Then dstPixels(x * 4 + 3) = srcAs(Int(y * aSSfactorY) * aWidth + Int(x * aSSfactorX)) + aSgnComp
                        Else
                            yccY = srcRs(saOffset + x) + rSgnComp
                            yccB = srcGs(saOffset + x) + gSgnComp - 127
                            yccR = srcBs(saOffset + x) + bSgnComp - 127
                            If copyAlpha Then dstPixels(x * 4 + 3) = srcAs(saOffset + x) + aSgnComp
                        End If
                        
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
                        
                    Next x
                    
                End If
                
            'Proceed to next line
            Next y
            
        'In the future, we probably want to map this to a floating-point surface and propose tone-mapping
        ' (if ICC profile is missing).  For now however we perform a default linear map.
        ElseIf (imgChannels(0).prec > 8) Then
            
            'For now, do a quick drop to 8-bit data (tone-mapping in the future is TBD)
            hdrDivisor = 2 ^ (imgChannels(0).prec - 8)
            
            'Wrap a 1D array around the target line in in the destination image, and calculate an offset
            ' into the corresponding source channel line.
            VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(1).p_data, channelSizeEstimate
            VBHacks.WrapArrayAroundPtr_Long srcBs, srcBSA, imgChannels(2).p_data, channelSizeEstimate
            
            copyAlpha = m_OpjNotes.imgHasAlpha
            If copyAlpha Then VBHacks.WrapArrayAroundPtr_Long srcAs, srcASA, imgChannels(3).p_data, channelSizeEstimate
            
            'Iterate lines (data is stored top-down)
            For y = 0 To targetHeight - 1
            
                'Wrap a 1D array around the target line in in the destination image, and calculate an offset
                ' into the corresponding source channel line.
                dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
                saOffset = y * targetWidth
                
                'Further handling is separated by color type
                If (targetColorSpace = OPJ_CLRSPC_SRGB) Then
                        
                    If subSamplingActive Then
                            
                        For x = 0 To targetWidth - 1
                        
                            b = (srcBs(Int(y * bSSfactorY) * bWidth + Int(x * bSSfactorX)) + bSgnComp) \ hdrDivisor
                            If (b < 0) Then b = 0
                            If (b > 255) Then b = 255
                            
                            g = (srcGs(Int(y * gSSfactorY) * gWidth + Int(x * gSSfactorX)) + gSgnComp) \ hdrDivisor
                            If (g < 0) Then g = 0
                            If (g > 255) Then g = 255
                            
                            r = (srcRs(saOffset + x) + rSgnComp) \ hdrDivisor
                            If (r < 0) Then r = 0
                            If (r > 255) Then r = 255
                        
                            dstPixels(x * 4) = b
                            dstPixels(x * 4 + 1) = g
                            dstPixels(x * 4 + 2) = r
                            
                            If copyAlpha Then
                                a = (srcAs(Int(y * aSSfactorY) * aWidth + Int(x * aSSfactorX)) + aSgnComp) \ hdrDivisor
                                If (a < 0) Then a = 0
                                If (a > 255) Then a = 255
                                dstPixels(x * 4 + 3) = a
                            End If
                            
                        Next x
                        
                    Else
                    
                        For x = 0 To targetWidth - 1
                            
                            b = (srcBs(saOffset + x) + bSgnComp) \ hdrDivisor
                            If (b < 0) Then b = 0
                            If (b > 255) Then b = 255
                            g = (srcGs(saOffset + x) + gSgnComp) \ hdrDivisor
                            If (g < 0) Then g = 0
                            If (g > 255) Then g = 255
                            r = (srcRs(saOffset + x) + rSgnComp) \ hdrDivisor
                            If (r < 0) Then r = 0
                            If (r > 255) Then r = 255
                            
                            dstPixels(x * 4) = b
                            dstPixels(x * 4 + 1) = g
                            dstPixels(x * 4 + 2) = r
                            
                            If copyAlpha Then
                                a = (srcAs(saOffset + x) + aSgnComp) \ hdrDivisor
                                If (a < 0) Then a = 0
                                If (a > 255) Then a = 255
                                dstPixels(x * 4 + 3) = a
                            End If
                            
                        Next x
                        
                    End If
                    
                
                'YCC to RGB conversion taken from OpenJPEG itself: https://github.com/uclouvain/openjpeg/blob/e7453e398b110891778d8da19209792c69ca7169/src/bin/common/color.c#L74
                ElseIf (targetColorSpace = OPJ_CLRSPC_SYCC) Then
                    
                    For x = 0 To targetWidth - 1
                        
                        'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                        ' we split handling accordingly.
                        If subSamplingActive Then
                            yccY = (srcRs(saOffset + x) + rSgnComp) \ hdrDivisor
                            yccB = (srcGs(Int(y * gSSfactorY) * gWidth + Int(x * gSSfactorX)) + gSgnComp) \ hdrDivisor - 127
                            yccR = (srcBs(Int(y * bSSfactorY) * bWidth + Int(x * bSSfactorX)) + bSgnComp) \ hdrDivisor - 127
                            If copyAlpha Then dstPixels(x * 4 + 3) = (srcAs(Int(y * aSSfactorY) * aWidth + Int(x * aSSfactorX)) + aSgnComp) \ hdrDivisor
                        Else
                            yccY = (srcRs(saOffset + x) + rSgnComp) \ hdrDivisor
                            yccB = (srcGs(saOffset + x) + gSgnComp) \ hdrDivisor - 127
                            yccR = (srcBs(saOffset + x) + bSgnComp) \ hdrDivisor - 127
                            If copyAlpha Then dstPixels(x * 4 + 3) = (srcAs(saOffset + x) + aSgnComp) \ hdrDivisor
                        End If
                        
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
                        
                    Next x
                    
                End If
                
            Next y
            
        End If
        
        'Unwrap all temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        VBHacks.UnwrapArrayFromPtr_Long srcGs
        VBHacks.UnwrapArrayFromPtr_Long srcBs
        If copyAlpha Then VBHacks.UnwrapArrayFromPtr_Long srcAs
        dstDIB.UnwrapArrayFromDIB dstPixels
        dstDIB.SetAlphaPremultiplication True
        
        'Load complete!  (Clean-up is still required, however.)
        LoadJP2 = True
        
    End If
    
    'For now cleanup and exit
    GoTo SafeCleanup
    
    Exit Function
    
    'Code beyond this point performs a full clean-up of all internal and external resources for the current jp2 image
SafeCleanup:
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
        Set m_Stream = Nothing
    End If
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
        .isAtLeast8Bit = True
        ReDim .isChannelSubsampled(0 To numComponents) As Boolean
        ReDim .channelSsWidth(0 To numComponents) As Long
        ReDim .channelSsHeight(0 To numComponents) As Long
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
            .isAtLeast8Bit = (imgChannels(0).prec >= 8)
        End With
        Exit Function
    End If
    
    'If we're still here, this image has multiple channels.  Iterate up to 4 channels and count how many
    ' have dimensions matching the first component.  (Subsampling in JP2 files means each channel can have its
    ' own independent dimensions, and the caller is expected to scale all components to match in the file image.)
    Dim searchDepth As Long
    searchDepth = PDMath.Min2Int(numComponents, 4)
    
    Dim i As Long
    For i = 1 To searchDepth - 1
        
        'For now, stop if we meet a subsampled channel
        If (imgChannels(i).w <> targetWidth) Or (imgChannels(i).h <> targetHeight) Then
            m_OpjNotes.hasSubsampling = True
            m_OpjNotes.isChannelSubsampled(i) = True
            m_OpjNotes.channelSsWidth(i) = imgChannels(i).w
            m_OpjNotes.channelSsHeight(i) = imgChannels(i).h
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
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "openJPEG Info: " & Strings.StringFromCharPtr(pMsg, False), PDM_External_Lib
End Sub

Private Sub HandlerWarning(ByVal pMsg As Long, ByVal pUserData As Long)
    PDDebug.LogAction "openJPEG Warning: " & Trim$(Strings.StringFromCharPtr(pMsg, False)), PDM_External_Lib
End Sub

Private Sub HandlerError(ByVal pMsg As Long, ByVal pUserData As Long)
    PDDebug.LogAction "openJPEG Error: " & Strings.StringFromCharPtr(pMsg, False), PDM_External_Lib
End Sub

Private Function JP2_ReadProcDelegate(ByVal p_buffer As Long, ByVal p_nb_bytes As Long, ByVal p_user_data As Long) As Long
    
    If (Not m_Stream Is Nothing) Then
    
        'Return -1 when EOF is reached
        If (m_Stream.GetPosition() = m_Stream.GetStreamSize()) Then
            JP2_ReadProcDelegate = -1
        
        'Otherwise, read as many bytes as we can
        Else
            
            Dim numBytesToRead As Long, numBytesLeft As Long
            numBytesToRead = p_nb_bytes
            
            numBytesLeft = m_Stream.GetStreamSize() - m_Stream.GetPosition()
            If (numBytesLeft < numBytesToRead) Then numBytesToRead = numBytesLeft
            
            JP2_ReadProcDelegate = m_Stream.ReadBytesToBarePointer(p_buffer, numBytesToRead)
            
            'Once again, return -1 for the special case of reaching EOF (should have been caught above; this is just a failsafe)
            If (JP2_ReadProcDelegate = 0) Then JP2_ReadProcDelegate = -1
            
        End If
        
    End If
    
End Function

Private Function JP2_WriteProcDelegate(ByVal p_buffer As Long, ByVal p_nb_bytes As Long, ByVal p_user_data As Long) As Long
    'Debug.Print "JP2_WriteProcDelegate", p_nb_bytes
    'TBD!
End Function

Private Function JP2_SkipProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Currency
    'Debug.Print "JP2_SkipProcDelegate", p_nb_bytes
    If (Not m_Stream Is Nothing) Then
        If m_Stream.SetPosition(p_nb_bytes \ 10000, FILE_CURRENT) Then
            JP2_SkipProcDelegate = p_nb_bytes
        Else
            JP2_SkipProcDelegate = -1
        End If
    End If
End Function

Private Function JP2_SeekProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Long
    'Debug.Print "JP2_SeekProcDelegate", p_nb_bytes
    If (Not m_Stream Is Nothing) Then
        If m_Stream.SetPosition(p_nb_bytes \ 10000, FILE_BEGIN) Then
            JP2_SeekProcDelegate = 1
        Else
            JP2_SeekProcDelegate = 0
        End If
    End If
End Function
