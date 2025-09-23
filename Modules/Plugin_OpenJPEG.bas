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

Private Enum OPJ_CODEC_FORMAT
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

'Current image, if any
Private m_j2kImage As opj_image
    
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
Public Function IsFileJ2K(ByRef srcFile As String) As Boolean

    Const FUNC_NAME As String = "IsFileJ2K"
    IsFileJ2K = False
    
    If Files.FileExists(srcFile) Then
        
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
                        If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2_RFC3745"
                    End If
                End If
                
                If (Not IsFileJ2K) Then
                    
                    Const JP2_MAGIC As Long = &HA870A0D
                    Const J2K_CODESTREAM_MAGIC As Long = &H51FF4FFF
                    
                    IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(JP2_MAGIC), 4)
                    If IsFileJ2K And J2K_DEBUG_VERBOSE Then PDDebug.LogAction "File is JP2 stream"
                    If (Not IsFileJ2K) Then
                        IsFileJ2K = VBHacks.MemCmp(VarPtr(bFirst12(0)), VarPtr(J2K_CODESTREAM_MAGIC), 4)
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
    If (Not Plugin_OpenJPEG.IsFileJ2K(srcFile)) Then Exit Function
    
    'Still here?  This file passed basic validation.
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    
    'Populate a basic decoder struct
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder itself..."
    
    'Initialize a default decoder
    Dim pDecoder As Long
    pDecoder = opj_create_decompress(OPJ_CODEC_JP2)
    If (pDecoder = 0) Then
        InternalError FUNC_NAME, "opj_create_Decompress failed"
        Exit Function
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing callbacks..."
    
    'Initialize function wrappers.  (We do this via a twinBasic-built helper library.)
    Dim hInfo As Long, hWarning As Long, hError As Long
    PD_GetAddrJP2KCallbacks AddressOf HandlerInfo, AddressOf HandlerWarning, AddressOf HandlerError, hInfo, hWarning, hError
    
    opj_set_info_handler pDecoder, hInfo, 0&
    opj_set_warning_handler pDecoder, hWarning, 0&
    opj_set_error_handler pDecoder, hError, 0&
    
    'Set param settings here?
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    
    'Initialize the decoder with our param object
    Dim retOJ As Long
    retOJ = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOJ = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "setting strict mode..."
    
    'Strict mode is TBD
    If (opj_decoder_set_strict_mode(pDecoder, 0&) <> 1&) Then
        InternalError FUNC_NAME, "failed to set strict mode"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "allowing multithreaded decode..."
    
    'Allow the decoder to use as many threads as it wants
    If (opj_codec_set_threads(pDecoder, OS.LogicalCoreCount()) <> 1&) Then
        InternalError FUNC_NAME, "couldn't set thread count; single-thread mode will be used"
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    'Prep a generic reader on the target file
    Dim utf8Filename() As Byte
    Strings.UTF8FromString srcFile, utf8Filename, 0&
    'dParams.p_infile = VarPtr(utf8Filename(0))
    Dim pStream As Long
    pStream = opj_stream_create_default_file_stream(VarPtr(utf8Filename(0)), 1&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't create stream on target file; load abandoned"
        GoTo SafeCleanup
    End If
    
    If J2K_DEBUG_VERBOSE Then PDDebug.LogAction "Reading header..."
    
    Dim pImage As Long
'    Dim firstImageComp As opj_image_comp
'    m_j2kImage.p
'    pImage = VarPtr(m_j2kImage)
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    PDDebug.LogAction "read header successfully"
    VBHacks.CopyMemoryStrict VarPtr(m_j2kImage), pImage, LenB(m_j2kImage)
    PDDebug.LogAction "copied header successfully"
    PDDebug.LogAction m_j2kImage.x0 & ", " & m_j2kImage.y0 & ", " & m_j2kImage.x1 & ", " & m_j2kImage.y1 & ", " & m_j2kImage.numcomps
    
    'Finish decoding the rest of the image
    If (opj_decode(pDecoder, pStream, pImage) <> 1&) Then
        InternalError FUNC_NAME, "failed to decode image"
        GoTo SafeCleanup
    End If
    
    PDDebug.LogAction "decoded full image successfully"
    
    'Read to the end of the file to pull any other useful info (e.g. metadata)
    If (opj_end_decompress(pDecoder, pStream) <> 1&) Then
        InternalError FUNC_NAME, "failed to read to end of file"
        'Attempt to load image anyway
    End If
    
    PDDebug.LogAction "read to EOF successfully"
    
    'We can now pull channel data from the supplied image pointer.
    
    Dim NumChannels As Long
    NumChannels = m_j2kImage.numcomps
    
    If (NumChannels <= 0) Then
        PDDebug.LogAction "Invalid channel count: " & NumChannels
        GoTo SafeCleanup
    End If
    
    Dim imgChannels() As opj_image_comp
    ReDim imgChannels(0 To NumChannels - 1) As opj_image_comp
    
    Dim sizeOfChannel As Long, sizeOfChannelAligned As Long
    sizeOfChannelAligned = LenB(imgChannels(0))
    sizeOfChannel = Len(imgChannels(0))
    
    'To simplify reading data from an arbitrary pointer, construct a pdStream object against it.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=sizeOfChannelAligned * NumChannels, baseFilePointerOffset:=m_j2kImage.pComps, optimizeAccess:=OptimizeSequentialAccess) Then
        
        Dim i As Long
        For i = 0 To NumChannels - 1
            cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
            VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), cStream.ReadBytes_PointerOnly(sizeOfChannel), sizeOfChannel
            
            If J2K_DEBUG_VERBOSE Then
                PDDebug.LogAction "Channel #" & CStr(i + 1) & " info: "
                With imgChannels(i)
                    PDDebug.LogAction .x0 & ", " & .y0 & ", " & .w & ", " & .h & ", " & .prec & ", " & .Alpha
                End With
            End If
        
        Next i
        
    Else
        PDDebug.LogAction "Failed to initialize stream against component pointer."
        GoTo SafeCleanup
    End If
    
    'm_j2kImage
    
    'For now cleanup and exit
    GoTo SafeCleanup
    
    Exit Function
    
SafeCleanup:
    If (Not cStream Is Nothing) Then cStream.StopStream True
    If (pImage <> 0) Then opj_image_destroy pImage
    If (pStream <> 0) Then opj_stream_destroy pStream
    If (pDecoder <> 0) Then opj_destroy_codec pDecoder
    
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

