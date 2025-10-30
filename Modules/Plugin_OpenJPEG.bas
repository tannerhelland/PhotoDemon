Attribute VB_Name = "Plugin_OpenJPEG"
'***************************************************************************
'OpenJPEG (JPEG-2000) Library Interface
'Copyright 2025-2025 by Tanner Helland
'Created: 19/September/25
'Last updated: 30/October/25
'Last update: switch to custom-built OpenJPEG to enable support all the way back to Windows XP
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
'PhotoDemon originally used OpenJPEG via FreeImage, but FreeImage has since been abandoned so a new
' solution was needed. In 2025 I wrote a direct-to-OpenJPEG interface from scratch, custom-built for PD.
'
'This interface was designed against v2.5.4 of OpenJPEG (released 20 Sep 2025).  It should work fine
' with any version maintaining ABI compatibility with the 2.x line.
'
'The initial build of this interface passes all files in the official OpenJPEG conformance suite save one,
' which crashes due to a known OpenJPEG-caused error specific to 32-bit library builds:
' https://github.com/uclouvain/openjpeg/issues/1017
'
'I hope that means this interface is sufficiently robust for real-world usage!
'
'By my testing, PhotoDemon's current coverage of the JPEG-2000 spec is more extensive than any other
' open-source project, with identical behavior to OpenJPEG's reference implementations across a wide range
' of features and comptability settings.  PD is particularly adept at handling OpenJPEG images with
' unexpected precisions and/or combinations of esoteric features (like signed data-types).
'
'Note, however, that the JPEG-2000 format is poorly designed and very difficult for decoders to handle.
' For example, images are allowed to define their color space as "undefined" and it's the decoder's job
' to figure out how to handle this case.  OpenJPEG's authors (insanely) recommend querying color space from
' the ICC profile, if one exists, which is a bad idea for many reasons but primarily because ICC profiles
' don't - and shouldn't - need to define source color spaces because SANE FILE FORMATS PROVIDE THAT DATA
' IN THE HEADER BECAUSE EVERY SUBSEQUENT HANDLING OPERATION DEPENDS ON IT.
'
'GIMP devs encountered this same issue when attempting to improve JP2 support, also pointed out the
' insanity of it, and and ultimately decided on just asking the user to pick a (random?) color space at
' import-time if an image is self-defined as "unknown color space":
' https://github.com/uclouvain/openjpeg/issues/1103
'
'Because PD focuses heavily on automated batch processing of image files, load-time popups are an absolute
' last resort and my own experience with GIMP's import screen is that I, the user, have no fucking idea
' what color space a given JP2 image is in because that level of technical knowledge can NEVER be inferred
' by casual users on a per-image-file basis!  So basically, my current approach is to say "fuck it" and
' simply load unknown data as RGB.  This has so far produced good results on a the vast majority of
' "in the wild" test images.  Other images could theoretically be salvaged by providing a new PD-specific
' transform of e.g. YUV to RGB, but I have not added this (yet).
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
' (Note that this does expose the user to an increased of risk of crashes on malformed images, however - an
' inevitable consequence of attempting to resuscitate bad data streams.)
Private Const JP2_FORCE_STRICT_DECODING As Boolean = False

'Library handles will be non-zero if OpenJPEG is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE (not recommended except for fallback testing).
Private m_LibHandle As Long, m_LibAvailable As Boolean

'JPEG-2000 files can actually contain one of several different stream formats
Public Enum OPJ_CODEC_FORMAT
    OPJ_CODEC_UNKNOWN = -1  '/**< place-holder */
    OPJ_CODEC_J2K = 0       '/**< JPEG-2000 codestream : read/write */
    OPJ_CODEC_JPT = 1       '/**< JPT-stream (JPEG 2000, JPIP) : read only */
    OPJ_CODEC_JP2 = 2       '/**< JP2 file format : read/write */
    OPJ_CODEC_JPP = 3       '/**< JPP-stream (JPEG 2000, JPIP) : to be coded */
    OPJ_CODEC_JPX = 4       '/**< JPX file format (JPEG 2000 Part-2) : to be coded */
End Enum

#If False Then
    Private Const OPJ_CODEC_UNKNOWN = -1, OPJ_CODEC_J2K = 0, OPJ_CODEC_JPT = 1, OPJ_CODEC_JP2 = 2, OPJ_CODEC_JPP = 3, OPJ_CODEC_JPX = 4
#End If

'OpenJPEG foolishly allows "unspecified" color spaces, which is not only stupid but requires painful
' heuristics (or user-knowledge) to load data correctly.  PD generally infers color space from component count
' when necessary, defaulting to RGB or RGBA for 3- or 4-channel streams with unknown contents.
Private Enum OPJ_COLOR_SPACE
    OPJ_CLRSPC_UNKNOWN = -1     '/**< not supported by the library */
    OPJ_CLRSPC_UNSPECIFIED = 0  '/**< not specified in the codestream */
    OPJ_CLRSPC_SRGB = 1         '/**< sRGB */
    OPJ_CLRSPC_GRAY = 2         '/**< grayscale */
    OPJ_CLRSPC_SYCC = 3         '/**< YUV */
    OPJ_CLRSPC_EYCC = 4         '/**< e-YCC */
    OPJ_CLRSPC_CMYK = 5         '/**< CMYK */
End Enum

#If False Then
    Private Const OPJ_CLRSPC_UNKNOWN = -1, OPJ_CLRSPC_UNSPECIFIED = 0, OPJ_CLRSPC_SRGB = 1, OPJ_CLRSPC_GRAY = 2, OPJ_CLRSPC_SYCC = 3, OPJ_CLRSPC_EYCC = 4, OPJ_CLRSPC_CMYK = 5
#End If

'Comments on remaining structs are copied as-is from openjpeg.h

'/**
' * Decompression parameters
' * */

'#define OPJ_PATH_LEN 4096 /**< Maximum allowed size for filenames */
Private Const OPJ_PATH_LEN As Long = 4096

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

'Official OpenJPEG builds use stdcall, but without a def file, so we need to manually alias exports here
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
Private Declare Function opj_set_decoded_components Lib "openjp2" Alias "_opj_set_decoded_components@16" (ByVal p_codec As Long, ByVal numcomps As Long, ByVal p_comps_indices As Long, ByVal b_apply_color_transforms As Long) As Long

Private Declare Function opj_stream_create_default_file_stream Lib "openjp2" Alias "_opj_stream_create_default_file_stream@8" (ByVal p_fname As Long, ByVal p_is_read_stream As Long) As Long
Private Declare Function opj_stream_default_create Lib "openjp2" Alias "_opj_stream_default_create@4" (ByVal bool_p_is_input As Long) As Long
Private Declare Function opj_stream_create Lib "openjp2" Alias "_opj_stream_create@8" (ByVal p_buffer_size As Long, ByVal bool_p_is_input As Long) As Long
Private Declare Sub opj_stream_destroy Lib "openjp2" Alias "_opj_stream_destroy@4" (ByVal p_stream As Long)

'Default OpenJPEG builds assume cdecl callbacks, but I'm currently custom-building OpenJPEG with support for stdcall callbacks
' so we can handle everything from inside VB6
Private Declare Sub opj_stream_set_read_function Lib "openjp2" Alias "_opj_stream_set_read_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_write_function Lib "openjp2" Alias "_opj_stream_set_write_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_skip_function Lib "openjp2" Alias "_opj_stream_set_skip_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_seek_function Lib "openjp2" Alias "_opj_stream_set_seek_function@8" (ByVal p_stream As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_user_data Lib "openjp2" Alias "_opj_stream_set_user_data@12" (ByVal p_stream As Long, ByVal p_data As Long, ByVal p_function As Long)
Private Declare Sub opj_stream_set_user_data_length Lib "openjp2" Alias "_opj_stream_set_user_data_length@12" (ByVal p_stream As Long, ByVal data_length As Currency)

'Current OpenJPEG image header, if any
Private m_jp2Image As opj_image

'This type is a PD-specific struct used when importing JP2 data.  (PD has to make a number of decisions
' about how to handle JP2 data in a PD-compatible way; this struct will be passed around multiple functions
' as part of the decision-making process.)
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
    finalPrecision As Long
End Type

Private m_OpjNotes As PD_OpjNotes

'This pdStream object reads/write actual JP2 data, using the callbacks supplied to OpenJPEG
Private m_Stream As pdStream

'ICC profile of the embedded image, if any.  Check m_ICCLength <> 0 for profile presence.
Private m_IccBytes() As Byte, m_IccLength As Long, m_ColorProfile As pdICCProfile

'JP2 files require extraordinary amounts of memory to decode (4-bytes per channel, so 16-bytes per pixel for RGBA!)
' but we can request down-sampling during the load process to "salvage" extraordinarily huge images that are
' otherwise "unloadable" on 32-bit systems.
Private Const MAX_SIZE_COMPONENTS As Long = 120000000
Private m_SizeReduction As Long
    
'Forcibly disable library interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

'OpenJPEG exports a dedicated version-reporting function
Public Function GetVersion() As String

    If (m_LibHandle <> 0) And m_LibAvailable Then
        Dim ptrVersion As Long
        ptrVersion = opj_version()
        GetVersion = Strings.StringFromCharPtr(ptrVersion, False)
    End If
    
End Function

'Must be called before any OpenJPEG functions are used.
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
        ' OpenJPEG's official reference implementation:
        ' https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c
        Dim srcExtension As String
        srcExtension = Files.FileGetExtension(srcFile)
        
        If Strings.StringsEqual(srcExtension, "jpt", True) Then
            outCodecFormat = OPJ_CODEC_JPT
            IsFileJP2 = True
            If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "File is JPT stream"
            Exit Function
        End If
        
        'We need at least 12-bytes to make a concrete determination
        If (Files.FileLenW(srcFile) < 12) Then
            IsFileJP2 = False
            Exit Function
        End If
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
            '12 bytes are enough to make a type determination
            Dim bFirst12() As Byte
            If cStream.ReadBytes(bFirst12, 12, True) Then
                
                'Various different signatures are valid, based on the container used.
                ' (Magic numbers taken from https://github.com/uclouvain/openjpeg/blob/41c25e3827c68a39b9e20c1625a0b96e49955445/src/bin/jp2/opj_decompress.c#L532)
                '#define JP2_RFC3745_MAGIC "\x00\x00\x00\x0c\x6a\x50\x20\x20\x0d\x0a\x87\x0a"
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
                    
                    '#define JP2_MAGIC "\x0d\x0a\x87\x0a"
                    Const JP2_MAGIC As Long = &HA870A0D
                    '#define J2K_CODESTREAM_MAGIC "\xff\x4f\xff\x51"
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
    
    'JPEG-2000 images can be extremely large.  PD will attempt to load all images at their defined size,
    ' but if an image fails due to memory constraints, we will try again at 1/4 size.  This process
    ' will happen 3x (reducing dimensions by another 75% each time) before we give up entirely.
    '
    'And yes, I'm gonna use GoTo to achieve this.)
    m_SizeReduction = 1
    
AttemptDecodingWithReduction:
    
    'Reset any module-level items from previous JPEG-2000 interactions.
    Erase m_IccBytes
    m_IccLength = 0
    Set m_ColorProfile = Nothing
    Set m_Stream = New pdStream
    
    'Initialize a default JPEG-2000 decoder.  Note that this requires us to know which codec to use in advance;
    ' *that's* why we need to identify the file header concretely in a previous step (note the codec return).
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder itself..."
    Dim pDecoder As Long
    pDecoder = opj_create_decompress(srcCodec)
    If (pDecoder = 0) Then
        InternalError FUNC_NAME, "opj_create_Decompress failed"
        Exit Function
    End If
    
    'Initialize local I/O functions for our constructed decoder.
    ' (Note that this requires a custom-built version of OpenJPEG with manually added support for stdcall callbacks.)
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing callbacks..."
    opj_set_info_handler pDecoder, AddressOf HandlerInfo, 0&
    opj_set_warning_handler pDecoder, AddressOf HandlerWarning, 0&
    opj_set_error_handler pDecoder, AddressOf HandlerError, 0&
    
    'Decoders support variable behavior via a "decoder parameter" struct.
    ' Populate a parameter struct with default values.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    'If you want to set custom decoding parameters, do it here.
    ' (For now, PD uses default decoding params.)
    
    'NOTE: to reduce an excessively large image, you can set a reduction factor here (2 ^ n).
    ' This can salvage large images that won't otherwise load on 32-bit systems.
    If (m_SizeReduction <> 1) Then dParams.cp_reduce = m_SizeReduction \ 2
    
    'Load our decoding parameters into the decoder object
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    Dim retOpj As Long
    retOpj = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOpj = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    'Decoders can use a "strict" mode, where incomplete or broken JP2 streams are simply disallowed.
    '
    '(Non-strict mode tells the decoder to decode as much as they can, and stop when they
    '  reach EOF or some other meaningful marker in the file - this can allow *some* files to be partially recovered,
    '  and testing shows a fair amount of in-the-wild images require strict turned OFF to work at all.)
    Dim strictModeValue As Long
    If JP2_FORCE_STRICT_DECODING Then strictModeValue = 1& Else strictModeValue = 0&
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "setting strict mode to " & UCase$(CStr(JP2_FORCE_STRICT_DECODING)) & "..."
    If (opj_decoder_set_strict_mode(pDecoder, strictModeValue) <> 1&) Then
        InternalError FUNC_NAME, "failed to set strictness mode"
        GoTo SafeCleanup
    End If
    
    'Allow the decoder to use as many logical threads as it wants; this provides meaningful perf improvements
    ' on modern systems.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "allowing multithreaded decode..."
    If (opj_codec_set_threads(pDecoder, OS.LogicalCoreCount()) = 1&) Then
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Allowing OpenJPEG to use " & OS.LogicalCoreCount() & " cores"
    Else
        InternalError FUNC_NAME, "couldn't set thread count; single-thread mode will be used"
    End If
    
    'Prep a generic OpenJPEG-specific "stream" (read-only) against the target file.
    ' This generic object will call our I/O functions for all behaviors, but it's still required as a parameter
    ' for OpenJPEG-specific read functions.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    'OpenJPEG's built-in file stream doesn't support Unicode chars on Windows, so we must manually supply I/O callbacks.
    
    'Open a pdStream object on the target file.  (Buffer size doesn't matter here; OpenJPEG will request blocks
    ' in its own preferred size, and the pdStream class handles their size requests automatically.)
    Set m_Stream = New pdStream
    If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, optimizeAccess:=OptimizeSequentialAccess) Then
        InternalError FUNC_NAME, "couldn't start pdStream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "pdStream initialized OK..."
    
    'Start a blank OpenJPEG memory stream.  Again, this stream object won't actually touch the file -
    ' it'll simply copy over whatever chunks of file data *we* supply.)
    Dim pStream As Long
    pStream = opj_stream_default_create(1&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't start blank jp2 stream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Blank jp2 stream initialized OK..."
    
    'Pass our local I/O callbacks to OpenJPEG.
    ' (Note that this requires a custom-built version of OpenJPEG with manually added support for stdcall callbacks.)
    opj_stream_set_user_data pStream, 0&, 0&
    opj_stream_set_user_data_length pStream, Files.FileLenW(srcFile) \ 10000
    opj_stream_set_read_function pStream, AddressOf JP2_ReadProcDelegate
    opj_stream_set_write_function pStream, AddressOf JP2_WriteProcDelegate
    opj_stream_set_skip_function pStream, AddressOf JP2_SkipProcDelegate
    opj_stream_set_seek_function pStream, AddressOf JP2_SeekProcDelegate
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "I/O callbacks assigned OK..."
    
    'With the file set up, we can now attempt to read the header.
    ' (From this point forward, OpenJPEG will call our I/O callbacks as necessary to grab file bits.)
    Dim pImage As Long
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Header read successfully"
    
    'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it
    ' and conveniently access relevant struct members.
    '
    'NOTE: DO NOT RELY ON THESE INITIAL HEADER VALUES FOR ANYTHING BUT "OH, THAT'S INTERESTING" VALUE.
    ' Why? Because the contents of the header can CHANGE between now and decoding the rest of the image.
    ' OpenJPEG sometimes makes new decisions about how to handle image data, like what dimensions it will use
    ' (because component dimensions conflict with header dimensions) or how many color components exist or
    ' WHAT those components represent.  For example, file9.jp2 in the official conformance suite is defined as a
    ' grayscale image by the header, but when you decode the file, it magically turns to color.  Why?  OpenJPEG finds
    ' multiple channels and just decides to run with 'em.  So you can't trust the actual file header, OR any headers
    ' returned by OpenJPEG until AFTER all components have been decoded.
    '
    '(Note that this behavior is particularly problematic for PD, because sometimes we need to make memory
    ' decisions based on image size - like only decoding a reduced-size copy of the image if 32-bit memory limits
    ' are a concern - but we CAN'T ACTUALLY DO THAT YET because the image's size might change post-decoding.
    '
    'And yes, I'm extremely frustrated by this baffling behavior, because it's unintuitive and caused me a ton
    ' of frustration solving inexplicable crashes caused by OpenJPEG internally changing image descriptors after
    ' initially reading the header.
    VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), pImage, LenB(m_jp2Image)
    If JP2_DEBUG_VERBOSE Then
        PDDebug.LogAction "Initial header read (DO NOT USE YET; FOR REFERENCE ONLY):"
        PDDebug.LogAction m_jp2Image.x0 & ", " & m_jp2Image.y0 & ", " & m_jp2Image.x1 & ", " & m_jp2Image.y1 & ", " & m_jp2Image.numcomps & " " & GetNameOfOpjColorSpace(m_jp2Image.color_space) & " components"
    End If
    
    'If the image is too large to load on this system, reduce size to 25% of current size and try again.
    ' (We'll repeat this up to 3 times before giving up and abandoning loading entirely.)
    If ((m_jp2Image.x1 * m_jp2Image.y1 * m_jp2Image.numcomps) \ m_SizeReduction > MAX_SIZE_COMPONENTS) And (m_SizeReduction < 4) Then
        
        m_SizeReduction = m_SizeReduction * 2
        
        'Free any open objects before attempting a new load
        If (Not m_Stream Is Nothing) Then
            If m_Stream.IsOpen() Then m_Stream.StopStream True
            Set m_Stream = Nothing
        End If
        If (pImage <> 0) Then opj_image_destroy pImage
        pImage = 0
        If (pStream <> 0) Then opj_stream_destroy pStream
        pStream = 0
        If (pDecoder <> 0) Then opj_destroy_codec pDecoder
        pDecoder = 0
        
        'That's right, suckers - I'm using GoTo and no one can stop me!  MWAHAHAHAHA
        GoTo AttemptDecodingWithReduction
        
    End If
    
    'Finish decoding all pixel data.  (Expect multiple I/O callbacks here.)
    If (opj_decode(pDecoder, pStream, pImage) <> 1&) Then
        InternalError FUNC_NAME, "failed to decode image"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Decoded full image successfully"
    
    'Although pixel data is done, we still need to read to the end of the file to pull other potentially useful info
    ' (metadata, ICC profile, etc)
    If (opj_end_decompress(pDecoder, pStream) <> 1&) Then
        InternalError FUNC_NAME, "failed to read to end of file"
        'Malformed data beyond pixel encoding isn't a deal-breaker.  We'll still attempt to load the image.
    Else
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Reached EOF successfully"
    End If
    
    'Only NOW can we actually read the header, because it may have changed from previous accesses.
    '
    'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it
    ' and conveniently access relevant struct members.
    VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), pImage, LenB(m_jp2Image)
    If JP2_DEBUG_VERBOSE Then
        PDDebug.LogAction "Final header values:"
        PDDebug.LogAction m_jp2Image.x0 & ", " & m_jp2Image.y0 & ", " & m_jp2Image.x1 & ", " & m_jp2Image.y1 & ", " & m_jp2Image.numcomps & " " & GetNameOfOpjColorSpace(m_jp2Image.color_space) & " components"
    End If
    
    'If the image has an ICC profile, copy it into a module-level array.
    If (m_jp2Image.icc_profile_len > 0) And (m_jp2Image.pIccProfile <> 0) Then
        m_IccLength = m_jp2Image.icc_profile_len
        ReDim m_IccBytes(0 To m_IccLength - 1) As Byte
        VBHacks.CopyMemoryStrict VarPtr(m_IccBytes(0)), m_jp2Image.pIccProfile, m_IccLength
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Found ICC profile (" & m_IccLength & " bytes)"
    Else
        m_IccLength = 0
        Erase m_IccBytes
    End If
    
    'With the image data fully parsed, we can now pull channel data from the supplied component pointer(s).
    
    'Handling varies by component count.  PD can handle 1, 3, 4-channel data OK; for higher counts, it'll just grab
    ' the first 1/3/4 (depending on color model) and use them as-is.
    '
    '2-channel is currently treated as grayscale, but handling could be added for grayscale+alpha if I can find
    ' a relevant official image "in the wild" that uses this combination.
    Dim numComponents As Long
    numComponents = m_jp2Image.numcomps
    
    If (numComponents <= 0) Then
        PDDebug.LogAction "Invalid channel count: " & numComponents
        GoTo SafeCleanup
    End If
    
    Dim imgChannels() As opj_image_comp
    ReDim imgChannels(0 To numComponents - 1) As opj_image_comp
    
    'Next, we want to pull individual component headers.  This struct has non-aligned members
    ' (entries that don't align cleanly on 4-byte boundaries), so we need to account for this when
    ' memcpy'ing them into local structs.
    Dim sizeOfChannel As Long, sizeOfChannelAligned As Long
    sizeOfChannelAligned = LenB(imgChannels(0))
    sizeOfChannel = Len(imgChannels(0))
    
    'To simplify reading data from arbitrary pointers, use a pdStream object.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=sizeOfChannelAligned * numComponents, baseFilePointerOffset:=m_jp2Image.pComps, optimizeAccess:=OptimizeSequentialAccess) Then
        
        'Pull *all* components into local structs that we can easily traverse via VB6 code
        Dim i As Long
        For i = 0 To numComponents - 1
        
            cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
            VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), cStream.ReadBytes_PointerOnly(sizeOfChannel), sizeOfChannel
            
            'Components often have unpredictable behavior, and it required a *lot* of debugging to solve.
            ' If users encounter problems on their own images, I need to see this info to resolve crashes.
            If JP2_DEBUG_VERBOSE Then
                PDDebug.LogAction "Channel #" & CStr(i + 1) & " info: "
                With imgChannels(i)
                    PDDebug.LogAction .x0 & ", " & .y0 & ", " & .w & ", " & .h & ", " & .prec & ", " & .Alpha
                    PDDebug.LogAction .p_data & ", " & .dx & ", " & .dy & ", " & .factor & ", " & .sgnd
                End With
            End If
        
        Next i
    
    'The only time stream construction would fail is if we're passed bad (null) pointers by OpenJPEG
    Else
        PDDebug.LogAction "Failed to initialize stream against component pointer."
        GoTo SafeCleanup
    End If
    
    'We are done with that temporary stream object; release it to free memory for the next (high-consumption) step
    cStream.StopStream True
    
    'With channel headers assembled, we now need to iterate channels and copy their contents into a local image object.
    
    'First, figure out what color space to use for the embedded data.  This is necessary because the JP2 designers
    ' (insanely) allow "unknown" as a descriptor for key fields like "what color space does this image use".
    ' This forces us to make hard decisions (or as GIMP decided, query the user at load-time) about what to
    ' do with "unknown" data, because a huge fraction of wild JP2 images use "unknown" despite being absolutely
    ' normal color spaces like RGB.
    '
    'To other designers: DO NOT ALLOW "UNKNOWN" AS AN OFFICIAL VALUE IN YOUR SPEC.  THIS IS STUPID AND DEFEATS
    ' THE WHOLE POINT OF A SPECIFICATION.
    Dim targetColorSpace As OPJ_COLOR_SPACE
    targetColorSpace = DetermineColorHandling(m_jp2Image.color_space, numComponents, imgChannels)
    If (targetColorSpace <> OPJ_CLRSPC_GRAY) And (targetColorSpace <> OPJ_CLRSPC_SRGB) And (targetColorSpace <> OPJ_CLRSPC_SYCC) Then
        PDDebug.LogAction "Unknown color space or component count.  Abandoning load."
        GoTo SafeCleanup
    End If
    
    'JP2 images allow each component to have their own size.  At load-time, the caller is responsible for
    ' normalizing all sizes.  (This is handled differently by nearly every JP2 library, because the spec
    ' doesn't properly define behavior like "what resampling algorithm to use" or "how/when to round values", etc.)
    '
    'Because subsampling can impose a significant performance hit, PD tracks it as on a per-component basis,
    ' with alternate (much faster) load pathways for non-subsampled channels.  This provides a large speed
    ' improvement over e.g. FreeImage's approach.
    '
    '(NOTE: yes, even the *first* channel in the image can technically be subsampled!  We do cover this case,
    '  but only because it shows up in the official conformance suite lol.)
    Dim subSamplingActive As Boolean
    subSamplingActive = m_OpjNotes.hasSubsampling
    
    Dim rSSfactorX As Single, gSSfactorX As Single, bSSfactorX As Single, aSSfactorX As Single
    Dim rSSfactorY As Single, gSSfactorY As Single, bSSfactorY As Single, aSSfactorY As Single
    Dim rWidth As Long, gWidth As Long, bWidth As Long, aWidth As Long
    
    Dim rSSPosX() As Long, rSSPosY() As Long, gSSPosX() As Long, gSSPosY() As Long, bSSPosX() As Long, bSSPosY() As Long
    Dim aSSPosX() As Long, aSSPosY() As Long
    
    Dim x As Long, y As Long
    
    'OpenJPEG's reference library uses rounding when upsampling, but using a naive 0.5 factor can produce
    ' OOB errors on images with odd-numbered heights.  A slightly sub-0.5 rounding factor prevents this.
    ' (This produces results basically identical to the official reference library - I say "basically"
    '  because all images in the official conformance suite match, but I can't test every image in existence.)
    Dim ssRoundingFactor As Single
    ssRoundingFactor = 0.4999!
    
    'ADDED OCT 2025: if an image only has one channel, ignore the image header's defined dimensions
    ' and instead force the image to the single channel's dimensions.  This matches OpenJPEG's reference
    ' handling of this case.
    If subSamplingActive And (numComponents >= 1) Then
        
        'Calculate indices into each color channel.  (Note that we use RGBA indices, but channels may represent
        ' other color data - the spec foolishly doesn't provide a way to determine canonical color representation,
        ' and testing shows that pretty much all wild files use standard RGB/YUV order convention.)
        If (numComponents >= 1) Then
            rSSfactorX = CDbl(m_OpjNotes.channelSsWidth(0)) / CDbl(m_OpjNotes.finalWidth)
            rSSfactorY = CDbl(m_OpjNotes.channelSsHeight(0)) / CDbl(m_OpjNotes.finalHeight)
            rWidth = m_OpjNotes.channelSsWidth(0)
            
            'If subsampling is active, precalculate all offsets into the image (with correct rounding)
            ' and store the calculated positions in local LUTs.  This is faster than calculating offsets repeatedly
            ' across millions of pixels.
            ReDim rSSPosX(0 To m_OpjNotes.finalWidth - 1) As Long
            For x = 0 To m_OpjNotes.finalWidth - 1
                rSSPosX(x) = Int(x * rSSfactorX + ssRoundingFactor)
            Next x
            
            ReDim rSSPosY(0 To m_OpjNotes.finalHeight - 1) As Long
            For y = 0 To m_OpjNotes.finalHeight - 1
                rSSPosY(y) = Int(y * rSSfactorY + ssRoundingFactor) * rWidth
            Next y
            
        End If
        
        If (numComponents >= 2) Then
            gSSfactorX = CDbl(m_OpjNotes.channelSsWidth(1)) / CDbl(m_OpjNotes.finalWidth)
            gSSfactorY = CDbl(m_OpjNotes.channelSsHeight(1)) / CDbl(m_OpjNotes.finalHeight)
            gWidth = m_OpjNotes.channelSsWidth(1)
            ReDim gSSPosX(0 To m_OpjNotes.finalWidth - 1) As Long
            For x = 0 To m_OpjNotes.finalWidth - 1
                gSSPosX(x) = Int(x * gSSfactorX + ssRoundingFactor)
            Next x
            ReDim gSSPosY(0 To m_OpjNotes.finalHeight - 1) As Long
            For y = 0 To m_OpjNotes.finalHeight - 1
                gSSPosY(y) = Int(y * gSSfactorY + ssRoundingFactor) * gWidth
            Next y
        End If
        
        If (numComponents >= 3) Then
            bSSfactorX = CDbl(m_OpjNotes.channelSsWidth(2)) / CDbl(m_OpjNotes.finalWidth)
            bSSfactorY = CDbl(m_OpjNotes.channelSsHeight(2)) / CDbl(m_OpjNotes.finalHeight)
            bWidth = m_OpjNotes.channelSsWidth(2)
            ReDim bSSPosX(0 To m_OpjNotes.finalWidth - 1) As Long
            For x = 0 To m_OpjNotes.finalWidth - 1
                bSSPosX(x) = Int(x * bSSfactorX + ssRoundingFactor)
            Next x
            
            ReDim bSSPosY(0 To m_OpjNotes.finalHeight - 1) As Long
            For y = 0 To m_OpjNotes.finalHeight - 1
                bSSPosY(y) = Int(y * bSSfactorY + ssRoundingFactor) * bWidth
            Next y
        End If
        
        If (numComponents >= 4) Then
            aSSfactorX = CDbl(m_OpjNotes.channelSsWidth(3)) / CDbl(m_OpjNotes.finalWidth)
            aSSfactorY = CDbl(m_OpjNotes.channelSsHeight(3)) / CDbl(m_OpjNotes.finalHeight)
            aWidth = m_OpjNotes.channelSsWidth(3)
            ReDim aSSPosX(0 To m_OpjNotes.finalWidth - 1) As Long
            For x = 0 To m_OpjNotes.finalWidth - 1
                aSSPosX(x) = Int(x * aSSfactorX + ssRoundingFactor)
            Next x
            
            ReDim aSSPosY(0 To m_OpjNotes.finalHeight - 1) As Long
            For y = 0 To m_OpjNotes.finalHeight - 1
                aSSPosY(y) = Int(y * aSSfactorY + ssRoundingFactor) * aWidth
            Next y
        End If
        
    End If
    
    'With (up-to-four) channels successfully sized, and a determination made on color mode handling,
    ' we can now load pixel data.
    
    'Prep the destination image object.
    Set dstDIB = New pdDIB
    dstDIB.CreateBlank m_OpjNotes.finalWidth, m_OpjNotes.finalHeight, 32, RGB(255, 255, 255), 255
    
    Dim targetWidth As Long, targetHeight As Long
    Dim channelSizeEstimate As Long
    Dim srcRs() As Long, srcGs() As Long, srcBs() As Long, srcAs() As Long
    Dim srcRSA As SafeArray1D, srcGSA As SafeArray1D, srcBSA As SafeArray1D, srcASA As SafeArray1D
    Dim copyAlpha As Boolean
            
    Dim r As Long, g As Long, b As Long, a As Long, yccY As Long, yccB As Long, yccR As Long
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    Dim saOffset As Long, xOffset As Long, hdrDivisor As Long
    
    'Data in JP2 files can be signed, meaning that e.g. 8-bit data is represented as [-127, 128] instead of [0, 255].
    ' PD handles this case successfully.
    Dim rIsSigned As Boolean, gIsSigned As Boolean, bIsSigned As Boolean, aIsSigned As Boolean
    Dim rSgnComp As Long, gSgnComp As Long, bSgnComp As Long, aSgnComp As Long
    
    'Unlike other image format libraries, OpenJPEG always loads channels as 4-byte ints (Longs in VB6)
    ' regardless of embedded color depth.  This is incredibly wasteful from a memory standpoint,
    ' but it does simplify handling of various bit-depths, because the source channel data is always the
    ' same size per-pixel.
    targetWidth = m_OpjNotes.finalWidth
    targetHeight = m_OpjNotes.finalHeight
    channelSizeEstimate = targetWidth * targetHeight * 4    '(See above note - this line is not a typo!)
    
    Dim finalPrec As Long
    finalPrec = imgChannels(0).prec
    m_OpjNotes.finalPrecision = finalPrec
    
    'Load pixel data, with handling separated by color type
    If (targetColorSpace = OPJ_CLRSPC_GRAY) Then
        
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Importing image using grayscale decoder..."
        
        'Handle the sign bit universally for all channels
        gIsSigned = (imgChannels(0).sgnd <> 0)
        If gIsSigned Then rSgnComp = (2 ^ imgChannels(0).prec) \ 2 Else rSgnComp = 0
        
        'Wrap a temporary array around the source channel
        VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
        
        'Precision can technically be any value between 1 and ???? (upper limit is unclear from the spec).
        ' All images from the official conformance spec are handled correctly, but PD hasn't been tested
        ' against 32-bit unsigned data.  (All other precision+signed combinations work well!)
        
        'Sub-8pp channels
        If (finalPrec < 8) Then
            
            'Reusing variable names is stupid, but here we are!  This value is multiplied by the sub-8-bit component value
            ' to arrive at a value on the range [0, 255].
            '
            'Well, TECHNICALLY it won't be on the range [0, 255] because e.g. a 4-bit image will go from [0, 15] to [0, 240].
            ' I only do it this way because that's what the official OpenJPEG library does, and it's bad but we need to
            ' mimic their behavior for consistency.
            hdrDivisor = 2 ^ (8 - finalPrec)
        
        'Normal 8-bpp channels require no extra handling
        'ElseIf (imgChannels(0).prec = 8) Then
        
        'HDR data needs to be downsampled to 8-bpp for PD.
        ' (TODO at some future date: in the absence of an ICC profile, allow the user to tone-map the data as they wish?
        Else
            hdrDivisor = 2 ^ (finalPrec - 8)
        End If
        
        'Iterate lines (data is stored top-down)
        For y = 0 To targetHeight - 1
            
            'Wrap a 1D array around the target line in in the destination image, and calculate an offset
            ' into the corresponding source channel line.
            dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
            saOffset = y * targetWidth
            
            For x = 0 To targetWidth - 1
                
                'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                ' we split handling.
                If subSamplingActive Then
                    g = srcRs(rSSPosY(y) + rSSPosX(x))
                Else
                    g = srcRs(saOffset + x)
                End If
                
                g = g + rSgnComp
                
                'Up-sample low-precision
                If (finalPrec < 8) Then
                    g = g * hdrDivisor
                'Down-sample high-precision
                ElseIf (finalPrec > 8) Then
                    g = g \ hdrDivisor
                End If
                
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
        
        'Unwrap all temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        dstDIB.UnwrapArrayFromDIB dstPixels
        dstDIB.SetAlphaPremultiplication True
        
        'Load complete!  (Clean-up is still required, however.)
        LoadJP2 = True
        
        'End grayscale handling
    
    'Color and YCC spaces are handled together
    ElseIf (targetColorSpace = OPJ_CLRSPC_SRGB) Or (targetColorSpace = OPJ_CLRSPC_SYCC) Or (targetColorSpace = OPJ_CLRSPC_EYCC) Then
        
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Importing image using color decoder..."
        
        'Color channels can be signed or unsigned.  For example, 8-bit SIGNED data needs to be
        ' treated as if on the range [-127, 127] while signed is [0, 255].  Signed state is available
        ' in all precisions (and for all channels, including alpha) and must be handled accordingly.
        rIsSigned = (imgChannels(0).sgnd <> 0)
        gIsSigned = (imgChannels(1).sgnd <> 0)
        bIsSigned = (imgChannels(2).sgnd <> 0)
        If m_OpjNotes.imgHasAlpha Then aIsSigned = (imgChannels(3).sgnd <> 0) Else aIsSigned = False
        
        If rIsSigned Then rSgnComp = (2 ^ imgChannels(0).prec) \ 2 Else rSgnComp = 0
        If gIsSigned Then gSgnComp = (2 ^ imgChannels(1).prec) \ 2 Else gSgnComp = 0
        If bIsSigned Then bSgnComp = (2 ^ imgChannels(2).prec) \ 2 Else bSgnComp = 0
        If aIsSigned Then aSgnComp = (2 ^ imgChannels(3).prec) \ 2 Else aSgnComp = 0
        
        'Wrap temporary arrays around 3 (or 4) source channels
        VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
        VBHacks.WrapArrayAroundPtr_Long srcGs, srcGSA, imgChannels(1).p_data, channelSizeEstimate
        VBHacks.WrapArrayAroundPtr_Long srcBs, srcBSA, imgChannels(2).p_data, channelSizeEstimate
        
        copyAlpha = m_OpjNotes.imgHasAlpha
        If copyAlpha Then VBHacks.WrapArrayAroundPtr_Long srcAs, srcASA, imgChannels(3).p_data, channelSizeEstimate
        
        'Precision can technically be any value between 1 and ???? (upper limit is unclear from the spec).
        ' All images from the official conformance spec are handled correctly, but PD hasn't been tested
        ' against 32-bit unsigned data.  (All other precision+signed combinations work well!)
        
        'Sub-8pp channels
        If (finalPrec < 8) Then
            
            'Reusing variable names is stupid, but here we are!  This value is multiplied by the sub-8-bit component value
            ' to arrive at a value on the range [0, 255].
            '
            'Well, TECHNICALLY it won't be on the range [0, 255] because e.g. a 4-bit image will go from [0, 15] to [0, 240].
            ' I only do it this way because that's what the official OpenJPEG library does, and it's bad but we need to
            ' mimic their behavior for consistency.
            hdrDivisor = 2 ^ (8 - finalPrec)
        
        'Normal 8-bpp channels require no extra handling
        'ElseIf (imgChannels(0).prec = 8) Then
        
        'HDR data needs to be downsampled to 8-bpp for PD.
        ' (TODO at some future date: in the absence of an ICC profile, allow the user to tone-map the data as they wish
        Else
            hdrDivisor = 2 ^ (finalPrec - 8)
        End If
        
        'Iterate lines (data is stored top-down)
        For y = 0 To targetHeight - 1
        
            'Wrap a 1D array around the target line in in the destination image, and calculate an offset
            ' into the corresponding source channel line.
            dstDIB.WrapArrayAroundScanline dstPixels, dstSA, y
            saOffset = y * targetWidth
            
            'Further handling is separated by color type
            If (targetColorSpace = OPJ_CLRSPC_SRGB) Then
            
                For x = 0 To targetWidth - 1
                
                    'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                    ' we split handling.
                    If subSamplingActive Then
                        b = srcBs(bSSPosY(y) + bSSPosX(x))
                        g = srcGs(gSSPosY(y) + gSSPosX(x))
                        r = srcRs(rSSPosY(y) + rSSPosX(x))
                    Else
                        b = srcBs(saOffset + x)
                        g = srcGs(saOffset + x)
                        r = srcRs(saOffset + x)
                    End If
                    
                    b = b + bSgnComp
                    g = g + gSgnComp
                    r = r + rSgnComp
                    
                    'Up-sample low-precision
                    If (finalPrec < 8) Then
                        b = b * hdrDivisor
                        g = g * hdrDivisor
                        r = r * hdrDivisor
                    'Down-sample high-precision
                    ElseIf (finalPrec > 8) Then
                        b = b \ hdrDivisor
                        g = g \ hdrDivisor
                        r = r \ hdrDivisor
                    End If
                    
                    If (b < 0) Then b = 0
                    If (b > 255) Then b = 255
                    
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                    
                    If (r < 0) Then r = 0
                    If (r > 255) Then r = 255
            
                    dstPixels(x * 4) = b
                    dstPixels(x * 4 + 1) = g
                    dstPixels(x * 4 + 2) = r
                    
                    'Repeat all the above steps for the alpha channel, as relevant
                    If copyAlpha Then
                        If subSamplingActive Then
                            a = srcAs(aSSPosY(y) + aSSPosX(x))
                        Else
                            a = srcAs(saOffset + x)
                        End If
                        a = a + aSgnComp
                        If (finalPrec < 8) Then
                            a = a * hdrDivisor
                        ElseIf (finalPrec > 8) Then
                            a = a \ hdrDivisor
                        End If
                        If (a < 0) Then a = 0
                        If (a > 255) Then a = 255
                        dstPixels(x * 4 + 3) = a
                    End If
                    
                Next x
                
            'YCC to RGB conversion taken from OpenJPEG itself: https://github.com/uclouvain/openjpeg/blob/e7453e398b110891778d8da19209792c69ca7169/src/bin/common/color.c#L74
            ' TODO: find eYCC images and test conversion; it likely needs different conversion matrices,
            ' but the conformance suite doesn't use that format so I'm SOL for testing currently
            ElseIf (targetColorSpace = OPJ_CLRSPC_SYCC) Or (targetColorSpace = OPJ_CLRSPC_EYCC) Then
                
                For x = 0 To targetWidth - 1
                    
                    'Subsampling imposes a perf penalty, so to improve performance on non-subsampled images,
                    ' we split handling accordingly.
                    If subSamplingActive Then
                        yccR = srcBs(bSSPosY(y) + bSSPosX(x))
                        yccB = srcGs(gSSPosY(y) + gSSPosX(x))
                        yccY = srcRs(rSSPosY(y) + rSSPosX(x))
                    Else
                        yccR = srcBs(saOffset + x)
                        yccB = srcGs(saOffset + x)
                        yccY = srcRs(saOffset + x)
                    End If
                    
                    'Handle signed vs unsigned
                    yccY = yccY + bSgnComp
                    yccB = yccB + gSgnComp
                    yccR = yccR + rSgnComp
                    
                    'Up-sample low-precision
                    If (finalPrec < 8) Then
                        yccY = yccY * hdrDivisor
                        yccB = yccB * hdrDivisor
                        yccR = yccR * hdrDivisor
                    'Down-sample high-precision
                    ElseIf (finalPrec > 8) Then
                        yccY = yccY \ hdrDivisor
                        yccB = yccB \ hdrDivisor
                        yccR = yccR \ hdrDivisor
                    End If
                    
                    'Scale B/R to the correct range for sRGB conversion
                    yccB = yccB - 127
                    yccR = yccR - 127
                    
                    'Convert from YUV to RGB and clamp for safety
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
                    
                    'Repeat all the above steps for the alpha channel, as relevant
                    If copyAlpha Then
                        If subSamplingActive Then
                            a = srcAs(aSSPosY(y) + aSSPosX(x))
                        Else
                            a = srcAs(saOffset + x)
                        End If
                        a = a + aSgnComp
                        If (finalPrec < 8) Then
                            a = a * hdrDivisor
                        ElseIf (finalPrec > 8) Then
                            a = a \ hdrDivisor
                        End If
                        If (a < 0) Then a = 0
                        If (a > 255) Then a = 255
                        dstPixels(x * 4 + 3) = a
                    End If
                    
                Next x
                
            End If
            
        'Proceed to next line
        Next y
        
        'Unwrap all temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        VBHacks.UnwrapArrayFromPtr_Long srcGs
        VBHacks.UnwrapArrayFromPtr_Long srcBs
        If copyAlpha Then VBHacks.UnwrapArrayFromPtr_Long srcAs
        dstDIB.UnwrapArrayFromDIB dstPixels
        
        'Load complete!  (Clean-up is still required, however.)
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "JP2 pixel data processed successfully!"
        LoadJP2 = True
        
    End If
    
    'With the target DIB now successfully constructed, we can apply color management (if an embedded color profile exists).
    If LoadJP2 And (m_IccLength <> 0) And ColorManagement.UseEmbeddedICCProfiles() And _
        ((targetColorSpace = OPJ_CLRSPC_SRGB) Or (targetColorSpace = OPJ_CLRSPC_SYCC) Or (targetColorSpace = OPJ_CLRSPC_EYCC)) Then
        
        'Copy the source ICC profile into a PD-specific ICC struct
        Set m_ColorProfile = New pdICCProfile
        If m_ColorProfile.LoadICCFromPtr(m_IccLength, VarPtr(m_IccBytes(0))) Then
    
            If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Applying color profile to image..."
            
            Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, tmpTransform As pdLCMSTransform
            Set srcProfile = New pdLCMSProfile
            
            'Ignore monochrome profiles (we can't apply them to pixels that have already been expanded to RGB)
            If srcProfile.CreateFromPDICCObject(m_ColorProfile) And (srcProfile.GetColorSpace <> cmsSigGray) Then
                
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
                
                If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "ICC profile applied."
                Set tmpTransform = Nothing
                Set srcProfile = Nothing
                
                'With the profile successfully applied, add this color profile to the central cache.
                Dim profHash As String
                profHash = ColorManagement.AddProfileToCache(m_ColorProfile, True, False, False)
                If (Not dstImage Is Nothing) Then dstImage.SetColorProfile_Original profHash
                dstDIB.SetColorManagementState cms_ProfileConverted
            
                'IMPORTANT NOTE: at present, the destination image - by the time we're done with it - will have been
                ' hard-converted to sRGB, so we don't want to associate the destination DIB with its source profile.
                ' Instead, note that it is currently sRGB.
                profHash = ColorManagement.GetSRGBProfileHash()
                dstDIB.SetColorProfileHash profHash
                
            End If
        End If
    
    End If
    
    If (Not dstDIB Is Nothing) Then dstDIB.SetAlphaPremultiplication True
    
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

Public Function GetComponentCountOfLastImage() As Long
     GetComponentCountOfLastImage = m_OpjNotes.numComponents
End Function

Public Function GetPrecisionOfLastImage() As Long
    GetPrecisionOfLastImage = m_OpjNotes.finalPrecision
End Function

'Figure out how to handle the source color data.  JPEG-2000 streams are extremely flexible in terms of color components
' (e.g. "undefined" color spaces and infinite color component counts are allowed, and each channel is allowed its own
' encoding method and/or grid dimensions via subsampling).  This makes them messy to handle, and a lot of software simply
' doesn't touch data that's encoded beyond non-subsampled 8-bpp RGB.
'
'Similarly, my goal here isn't necessarily to cover every possible combination of JP2 file attributes.  Instead, I want PD
' to make intelligent inferences about unknown data (e.g. three undefined channels are likely RGB, four is RGBA) and cover
' as many likely use-cases as I can.
'
'If an obvious correlation with a known color space cannot be made, PD will treat the image data as grayscale and load
' the first channel only.  This typically allows *something* to be recovered from the file.
Private Function DetermineColorHandling(ByVal fileColorSpace As OPJ_COLOR_SPACE, ByVal numComponents As Long, ByRef imgChannels() As opj_image_comp) As OPJ_COLOR_SPACE
    
    'An "unknown" color space notifies the caller that PD is unequipped to handle this image's data.
    ' ("Unknown" is an extremely common state of wild JP2 images, and PD will attempt to reassign that
    '  constant to something useable based on simple heuristics.)
    DetermineColorHandling = OPJ_CLRSPC_UNKNOWN
    
    'Failsafe check for component count (should have been validated by caller)
    If (numComponents <= 0) Then Exit Function
    
    'Some software only checks the size of the first component, and uses that as the size of the image.
    ' PD tries to use the header-defined image size instead (but note that the first channel *can* be subsampled,
    ' which is unlike other formats!)
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = m_jp2Image.x1 - m_jp2Image.x0
    targetHeight = m_jp2Image.y1 - m_jp2Image.y0
    
    If (m_SizeReduction <> 1) Then
        targetWidth = targetWidth \ m_SizeReduction
        targetHeight = targetHeight \ m_SizeReduction
    End If
    
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
    ' (Also, to match the reference OpenJPEG implementation, subsampling is ignored and the channel
    ' dimensions are forcibly used as the final image dimensions, regardless of what the image header
    ' actually says.)
    If (numComponents = 1) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        With m_OpjNotes
            .hasSubsampling = False
            .imgHasAlpha = False
            .idxAlphaChannel = -1
            .isChannelSubsampled(0) = False
            .isAtLeast8Bit = (imgChannels(0).prec >= 8)
            .channelSsWidth(0) = imgChannels(0).w
            .channelSsHeight(0) = imgChannels(0).h
            .finalWidth = .channelSsWidth(0)
            .finalHeight = .channelSsHeight(0)
        End With
        Exit Function
    End If
    
    'If we're still here, this image has multiple channels.  Iterate up to 4 channels and track specific channel data,
    ' including channel dimensions.  (Subsampling in JP2 files means each channel can have its own independent dimensions,
    ' and the caller is expected to scale all components to match in the final image.)
    Dim searchDepth As Long
    searchDepth = PDMath.Min2Int(numComponents, 4)
    
    Dim i As Long
    For i = 0 To searchDepth - 1
        
        'Flag subsampled channels
        If (imgChannels(i).w <> targetWidth) Or (imgChannels(i).h <> targetHeight) Then
            m_OpjNotes.hasSubsampling = True
            m_OpjNotes.isChannelSubsampled(i) = True
        End If
        
        'Track channel width/height independently (only relevant when a channel is subsampled)
        m_OpjNotes.channelSsWidth(i) = imgChannels(i).w
        m_OpjNotes.channelSsHeight(i) = imgChannels(i).h
        
        'Note the alpha channel index, if any.
        ' (NOTE: this implementation assumes the alpha channel is always channel 4 in a 4+ channel image, but technically
        '        channels can be "flagged" as alpha - if this happens, what does it mean?  IDK without test images.)
        If (i = 3) Then
            m_OpjNotes.idxAlphaChannel = i
            m_OpjNotes.imgHasAlpha = True
        End If
        
    Next i
    
    numComponents = i
    
    'For now, treat 2-component images as 1-component grayscale.
    If (numComponents < 3) Then numComponents = 1
    
    'Assign a correct color space based on channel count
    If (numComponents > 0) And (numComponents < 3) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        m_OpjNotes.imgHasAlpha = (m_OpjNotes.idxAlphaChannel >= 0)
        If (Not m_OpjNotes.imgHasAlpha) Then m_OpjNotes.numComponents = 1
    
    'CMYK was recently added as a potential JP2 color space, but I have not found any conformance images
    ' using this space so it's currently UNTESTED.
    ElseIf (numComponents > 3) Then
    
        If (fileColorSpace <> OPJ_CLRSPC_SRGB) And (fileColorSpace <> OPJ_CLRSPC_SYCC) And (fileColorSpace <> OPJ_CLRSPC_EYCC) Then
            PDDebug.LogAction "WARNING: 4-channel image found but color space is " & GetNameOfOpjColorSpace(fileColorSpace) & "; using sRGB pathway"
            DetermineColorHandling = OPJ_CLRSPC_SRGB
        Else
            DetermineColorHandling = fileColorSpace
        End If
        
        m_OpjNotes.imgHasAlpha = (m_OpjNotes.idxAlphaChannel >= 0)
        If m_OpjNotes.imgHasAlpha Then
            numComponents = 4
        Else
            numComponents = 3
        End If
    
    '3-component spaces can be RGB or YCC; we handle both (EYCC is currently untested pending test images, FYI)
    Else
        
        If (fileColorSpace = OPJ_CLRSPC_EYCC) Then
            DetermineColorHandling = OPJ_CLRSPC_EYCC
        ElseIf (fileColorSpace = OPJ_CLRSPC_SYCC) Then
            DetermineColorHandling = OPJ_CLRSPC_SYCC
        Else
            DetermineColorHandling = OPJ_CLRSPC_SRGB
        End If
        
        '3-component images never have alpha, regardless of how channels in the file are flagged.
        m_OpjNotes.imgHasAlpha = False
        
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
        '/**< CMYK */
        Case OPJ_CLRSPC_CMYK
            GetNameOfOpjColorSpace = "CMYK"
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

'OpenJPEG does not support wide chars in its default Windows I/O functions,
' so we need to supply our own callbacks and use them for all I/O behavior.
' (As a nice bonus, this also improves performance because we use memory mapped I/O which can
'  greatly improve throughput.)
Private Function JP2_ReadProcDelegate(ByVal p_buffer As Long, ByVal p_nb_bytes As Long, ByVal p_user_data As Long) As Long
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Read requested for " & p_nb_bytes
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
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Write requested for " & p_nb_bytes
    'Debug.Print "JP2_WriteProcDelegate", p_nb_bytes
    'TBD!
End Function

'Advance pointer [n] bytes in input file.
' Trial-and-error shows that this expects offsets relative to stream start, *not* current pointer.
Private Function JP2_SkipProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Currency
    
    'PD can't actually use 64-bit values (yet) for file seeks; use only the lower 4 bytes.
    ' (This workaround would not be needed in a 64-bit build.)
    Dim lowerFourSkip As Long
    VBHacks.GetMem4_Ptr VarPtr(p_nb_bytes), VarPtr(lowerFourSkip)
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Skip requested for " & lowerFourSkip
    
    If (Not m_Stream Is Nothing) Then
        If m_Stream.SetPosition(lowerFourSkip, FILE_BEGIN) Then
            JP2_SkipProcDelegate = p_nb_bytes
        Else
            JP2_SkipProcDelegate = -1
        End If
    End If
    
End Function

'Advance pointer [n] bytes in output file.
' TODO: debug this to see if seek is relative to file start or current pointer!
Private Function JP2_SeekProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Long

    'PD can't actually use 64-bit values (yet) for file seeks; use only the lower 4 bytes.
    ' (This workaround would not be needed in a 64-bit build.)
    Dim lowerFourSkip As Long
    VBHacks.GetMem4_Ptr VarPtr(p_nb_bytes), VarPtr(lowerFourSkip)
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Seek requested for " & lowerFourSkip
    
    If (Not m_Stream Is Nothing) Then
        If m_Stream.SetPosition(lowerFourSkip, FILE_BEGIN) Then
            JP2_SeekProcDelegate = p_nb_bytes
        Else
            JP2_SeekProcDelegate = -1
        End If
    End If
    
End Function
