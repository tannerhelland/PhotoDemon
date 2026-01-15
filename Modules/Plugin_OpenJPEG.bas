Attribute VB_Name = "Plugin_OpenJPEG"
'***************************************************************************
'OpenJPEG (JPEG-2000) Library Interface
'Copyright 2025-2026 by Tanner Helland
'Created: 19/September/25
'Last updated: 11/November/25
'Last update: finalize write support
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
' solution was needed. In 2025, I wrote a direct-to-OpenJPEG interface from scratch, custom-built for PD.
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
' IN THE HEADER BECAUSE EVERY SUBSEQUENT OPERATION DEPENDS ON IT.
'
'GIMP devs encountered this same issue when attempting to improve their own JP2 support, also pointed out
' the insanity of it, and and ultimately decided to just ask the user to pick a (random?) color space at
' import-time if an image is self-defined as "unknown color space":
' https://github.com/uclouvain/openjpeg/issues/1103
'
'Because PD focuses heavily on automated batch processing of image files, load-time popups are an absolute
' last resort and my own experience with GIMP's import screen is that I, the user, have no fucking idea
' what color space a given JP2 image is in because that level of technical knowledge can NEVER be inferred
' unless you created the image yourself!  So my current approach in PD is to say "fuck it" and simply load
' unknown data as RGB.  This has so far produced good results on the vast majority of "in the wild" images.
' Images with undefined non-RGB formats could theoretically be salvaged by providing a new PD-specific
' transform of e.g. YUV to RGB that the user can access directly, but I have not added this capability (yet).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Enable at DEBUG-TIME ONLY for verbose logging.  (Note: this interface is *very* verbose.)
Private Const JP2_DEBUG_VERBOSE As Boolean = False

'To strictly enforce the spec (and decrease chances of OpenJPEG crashes on malformed images), set this to TRUE.
' I currently set it to FALSE in production builds, to allow many more "in the wild" images to actually load.
' (Note that this exposes the user to an increased of risk of crashes on malformed images, however -
'  an inevitable consequence of attempting to resuscitate bad data streams.)
Private Const JP2_FORCE_STRICT_DECODING As Boolean = False

'Library handles will be non-zero when OpenJPEG is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE (not recommended except for fallback testing).
Private m_LibHandle As Long, m_LibAvailable As Boolean

'JPEG-2000 files can actually contain one of several different stream formats.
' PhotoDemon successfully reads J2K, JPT, and JP2 containers, but only writes JP2 ones.
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
' when necessary, defaulting to RGB or RGBA for 3- or 4-channel streams with unspecified contents.
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

'Comments on remaining structs are copied as-is from v2.5 of openjpeg.h:
'https://github.com/uclouvain/openjpeg/blob/1ad9bec2c12ee445ce23e660f5e4fe870b9d5e09/src/lib/openjp2/openjpeg.h

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
' * Component parameters structure used by the opj_image_create function
' * */
Private Type opj_image_comptparm
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
End Type

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

'Compression parameters

'/**
' * Progression order
' * */
Private Enum prog_order
    OPJ_PROG_UNKNOWN = -1   ',  /**< place-holder */
    OPJ_LRCP = 0            ',  /**< layer-resolution-component-precinct order */
    OPJ_RLCP = 1            ',  /**< resolution-layer-component-precinct order */
    OPJ_RPCL = 2            ',  /**< resolution-precinct-component-layer order */
    OPJ_PCRL = 3            ',  /**< precinct-component-resolution-layer order */
    OPJ_CPRL = 4            ' /**< component-precinct-resolution-layer order */
End Enum

'/**
' * Progression order changes
' *
' */
Private Type opj_poc
    '/** Resolution num start, Component num start, given by POC */
    resno0 As Long
    compno0 As Long
    '/** Layer num end,Resolution num end, Component num end, given by POC */
    layno1 As Long
    resno1 As Long
    compno1 As Long
    '/** Layer num start,Precinct num start, Precinct num end */
    layno0 As Long
    precno0 As Long
    precno1 As Long
    '/** Progression order enum*/
    prg1 As prog_order
    prg As prog_order
    '/** Progression order string*/
    progorder(0 To 4) As Byte   'TODO: investigate padding behavior here
    '/** Tile number (starting at 1) */
    tile As Long
    '/** Start and end values for Tile width and height*/
    tx0 As Long
    tx1 As Long
    ty0 As Long
    ty1 As Long
    '/** Start value, initialised in pi_initialise_encode*/
    layS As Long
    resS As Long
    compS As Long
    prcS As Long
    '/** End value, initialised in pi_initialise_encode */
    layE As Long
    resE As Long
    compE As Long
    prcE As Long
    '/** Start and end values of Tile width and height, initialised in pi_initialise_encode*/
    txS As Long
    txE As Long
    tyS As Long
    tyE As Long
    dx As Long
    dy As Long
    '/** Temporary values for Tile parts, initialised in pi_create_encode */
    lay_t As Long
    res_t As Long
    comp_t As Long
    prc_t As Long
    tx0_t As Long
    ty0_t As Long
End Type

'/**
' * DEPRECATED: use RSIZ, OPJ_PROFILE_* and OPJ_EXTENSION_* instead
' * Digital cinema operation mode
' * */
Private Enum OPJ_CINEMA_MODE
    OPJ_OFF = 0         ',    /** Not Digital Cinema*/
    OPJ_CINEMA2K_24 = 1 ',    /** 2K Digital Cinema at 24 fps*/
    OPJ_CINEMA2K_48 = 2 ',    /** 2K Digital Cinema at 48 fps*/
    OPJ_CINEMA4K_24 = 3 '     /** 4K Digital Cinema at 24 fps*/
End Enum

'/**
' * DEPRECATED: use RSIZ, OPJ_PROFILE_* and OPJ_EXTENSION_* instead
' * Rsiz Capabilities
' * */
Private Enum RSIZ_CAPABILITIES
    OPJ_STD_RSIZ = 0    ',       /** Standard JPEG2000 profile*/
    OPJ_CINEMA2K = 3    ',       /** Profile name for a 2K image*/
    OPJ_CINEMA4K = 4    ',       /** Profile name for a 4K image*/
    OPJ_MCT = &H8100&
End Enum

'/**
' * Compression parameters
' * */
Private Type opj_cparameters
    '/** size of tile: tile_size_on = false (not in argument) or = true (in argument) */
    tile_size_on As Long
    '/** XTOsiz */
    cp_tx0 As Long
    '/** YTOsiz */
    cp_ty0 As Long
    '/** XTsiz */
    cp_tdx As Long
    '/** YTsiz */
    cp_tdy As Long
    '/** allocation by rate/distortion */
    cp_disto_alloc As Long
    '/** allocation by fixed layer */
    cp_fixed_alloc As Long
    '/** allocation by fixed quality (PSNR) */
    cp_fixed_quality As Long
    '/** fixed layer */
    p_cp_matrice As Long
    '/** comment for coding */
    p_cp_comment As Long
    '/** csty : coding style */
    csty As Long
    '/** progression order (default OPJ_LRCP) */
    progorder As prog_order
    '/** progression order changes */
    POC(0 To 31) As opj_poc
    '/** number of progression order changes (POC), default to 0 */
    numpocs As Long
    '/** number of layers */
    tcp_numlayers As Long
    '/** rates of layers - might be subsequently limited by the max_cs_size field.
    ' * Should be decreasing. 1 can be
    ' * used as last value to indicate the last layer is lossless. */
    tcp_rates(0 To 99) As Single
    '/** different psnr for successive layers. Should be increasing. 0 can be
    ' * used as last value to indicate the last layer is lossless. */
    tcp_distoratio(0 To 99) As Single
    '/** number of resolutions */
    numresolution As Long
    '/** initial code block width, default to 64 */
    cblockw_init As Long
    '/** initial code block height, default to 64 */
    cblockh_init As Long
    '/** mode switch (cblk_style) */
    Mode As Long
    '/** 1 : use the irreversible DWT 9-7, 0 : use lossless compression (default) */
    irreversible As Long
    '/** region of interest: affected component in [0..3], -1 means no ROI */
    roi_compno As Long
    '/** region of interest: upshift value */
    roi_shift As Long
    '/* number of precinct size specifications */
    res_spec As Long
    '/** initial precinct width */
    '#define OPJ_J2K_MAXRLVLS 33                 /**< Number of maximum resolution level authorized */
    prcw_init(0 To 32) As Long
    '/** initial precinct height */
    prch_init(0 To 32) As Long

    '/**@name command line encoder parameters (not used inside the library) */
    '/*@{*/
    '/** input file name */
    '#define OPJ_PATH_LEN 4096 /**< Maximum allowed size for filenames */
    infile(0 To 4095) As Byte
    '/** output file name */
    outfile(0 To 4095) As Byte
    '/** DEPRECATED. Index generation is now handled with the opj_encode_with_info() function. Set to NULL */
    index_on As Long
    '/** DEPRECATED. Index generation is now handled with the opj_encode_with_info() function. Set to NULL */
    index_opj(0 To 4095) As Byte
    '/** subimage encoding: origin image offset in x direction */
    image_offset_x0 As Long
    '/** subimage encoding: origin image offset in y direction */
    image_offset_y0 As Long
    '/** subsampling value for dx */
    subsampling_dx As Long
    '/** subsampling value for dy */
    subsampling_dy As Long
    '/** input file format 0: PGX, 1: PxM, 2: BMP 3:TIF*/
    decod_format As Long
    '/** output file format 0: J2K, 1: JP2, 2: JPT */
    cod_format As Long
    '/*@}*/

    '/* UniPG>> */ /* NOT YET USED IN THE V2 VERSION OF OPENJPEG */
    '/**@name JPWL encoding parameters */
    '/*@{*/
    '/** enables writing of EPC in MH, thus activating JPWL */
    jpwl_epc_on As Long
    '/** error protection method for MH (0,1,16,32,37-128) */
    jpwl_hprot_MH As Long
    '/** tile number of header protection specification (>=0) */
    '#define JPWL_MAX_NO_TILESPECS   16 /**< Maximum number of tile parts expected by JPWL: increase at your will */
    jpwl_hprot_TPH_tileno(0 To 15) As Long
    '/** error protection methods for TPHs (0,1,16,32,37-128) */
    jpwl_hprot_TPH(0 To 15) As Long
    '/** tile number of packet protection specification (>=0) */
    '#define JPWL_MAX_NO_PACKSPECS   16 /**< Maximum number of packet parts expected by JPWL: increase at your will */
    jpwl_pprot_tileno(0 To 15) As Long
    '/** packet number of packet protection specification (>=0) */
    jpwl_pprot_packno(0 To 15) As Long
    '/** error protection methods for packets (0,1,16,32,37-128) */
    jpwl_pprot(0 To 15) As Long
    '/** enables writing of ESD, (0=no/1/2 bytes) */
    jpwl_sens_size As Long
    '/** sensitivity addressing size (0=auto/2/4 bytes) */
    jpwl_sens_addr As Long
    '/** sensitivity range (0-3) */
    jpwl_sens_range As Long
    '/** sensitivity method for MH (-1=no,0-7) */
    jpwl_sens_MH As Long
    '/** tile number of sensitivity specification (>=0) */
    jpwl_sens_TPH_tileno(0 To 15) As Long
    '/** sensitivity methods for TPHs (-1=no,0-7) */
    jpwl_sens_TPH(0 To 15) As Long
    '/*@}*/
    '/* <<UniPG */

    '/**
    ' * DEPRECATED: use RSIZ, OPJ_PROFILE_* and MAX_COMP_SIZE instead
    ' * Digital Cinema compliance 0-not compliant, 1-compliant
    ' * */
    cp_cinema As OPJ_CINEMA_MODE
    '/**
    ' * Maximum size (in bytes) for each component.
    ' * If == 0, component size limitation is not considered
    ' * */
    max_comp_size As Long
    '/**
    ' * DEPRECATED: use RSIZ, OPJ_PROFILE_* and OPJ_EXTENSION_* instead
    ' * Profile name
    ' * */
    cp_rsiz As RSIZ_CAPABILITIES
    '/** Tile part generation*/
    tp_on As Byte
    '/** Flag for Tile part generation*/
    tp_flag As Byte
    '/** MCT (multiple component transform) */
    tcp_mct As Byte
    '/** Enable JPIP indexing*/
    jpip_on As Long
    '/** Naive implementation of MCT restricted to a single reversible array based
    '    encoding without offset concerning all the components. */
    p_mct_data As Long
    '/**
    ' * Maximum size (in bytes) for the whole codestream.
    ' * If == 0, codestream size limitation is not considered
    ' * If it does not comply with tcp_rates, max_cs_size prevails
    ' * and a warning is issued.
    ' * */
    max_cs_size As Long
    '/** RSIZ value
    '    To be used to combine OPJ_PROFILE_*, OPJ_EXTENSION_* and (sub)levels values. */
    rsiz As Integer
End Type

'Official OpenJPEG builds use stdcall but without a def file, so I had to manually alias all exports here
Private Declare Function opj_version Lib "openjp2" Alias "_opj_version@0" () As Long
Private Declare Sub opj_set_default_decoder_parameters Lib "openjp2" Alias "_opj_set_default_decoder_parameters@4" (ByVal p_parameters As Long)
Private Declare Function opj_create_decompress Lib "openjp2" Alias "_opj_create_decompress@4" (ByVal jp2_format As OPJ_CODEC_FORMAT) As Long
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
'Private Declare Function opj_set_decoded_components Lib "openjp2" Alias "_opj_set_decoded_components@16" (ByVal p_codec As Long, ByVal numcomps As Long, ByVal p_comps_indices As Long, ByVal b_apply_color_transforms As Long) As Long
Private Declare Function opj_image_create Lib "openjp2" Alias "_opj_image_create@12" (ByVal numcmpts As Long, ByVal p_cmptparms As Long, ByVal clrspc As OPJ_COLOR_SPACE) As Long
Private Declare Function opj_create_compress Lib "openjp2" Alias "_opj_create_compress@4" (ByVal jp2_format As OPJ_CODEC_FORMAT) As Long
Private Declare Sub opj_set_default_encoder_parameters Lib "openjp2" Alias "_opj_set_default_encoder_parameters@4" (ByVal p_opj_cparameters_t As Long)
Private Declare Function opj_setup_encoder Lib "openjp2" Alias "_opj_setup_encoder@12" (ByVal p_codec As Long, ByVal p_parameters As Long, ByVal p_image As Long) As Long
Private Declare Function opj_start_compress Lib "openjp2" Alias "_opj_start_compress@12" (ByVal p_codec As Long, ByVal p_image As Long, ByVal p_stream As Long) As Long
Private Declare Function opj_encode Lib "openjp2" Alias "_opj_encode@8" (ByVal p_codec As Long, ByVal p_stream As Long) As Long
Private Declare Function opj_end_compress Lib "openjp2" Alias "_opj_end_compress@8" (ByVal p_codec As Long, ByVal p_stream As Long) As Long
'Private Declare Function opj_stream_create_default_file_stream Lib "openjp2" Alias "_opj_stream_create_default_file_stream@8" (ByVal p_fname As Long, ByVal p_is_read_stream As Long) As Long
Private Declare Function opj_stream_default_create Lib "openjp2" Alias "_opj_stream_default_create@4" (ByVal bool_p_is_input As Long) As Long
'Private Declare Function opj_stream_create Lib "openjp2" Alias "_opj_stream_create@8" (ByVal p_buffer_size As Long, ByVal bool_p_is_input As Long) As Long
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

'This is a PD-specific struct used when importing JP2 data.  (PD has to make a number of decisions
' about how to handle JP2 data in a PD-compatible way; this struct will be passed around multiple
' functions as part of the decision-making process.)
Private Type PD_OpjNotes
    finalWidth As Long
    finalHeight As Long
    numComponents As Long
    imgHasAlpha As Boolean
    idxAlphaChannel As Integer
    hasSubsampling As Boolean
    isChannelSubsampled() As Boolean
    channelSsWidth() As Long        'Subsampled width/height, in pixels, of channel at index [n]
    channelSsHeight() As Long
    finalPrecision As Long
End Type

Private m_OpjNotes As PD_OpjNotes

'This pdStream object reads/write actual JP2 data, via callbacks supplied to OpenJPEG.
' (This reference may also point to an external stream supplied by the caller.)
Private m_Stream As pdStream

'ICC profile of the embedded image, if any.  Check m_ICCLength <> 0 for profile presence.
Private m_IccBytes() As Byte, m_IccLength As Long, m_ColorProfile As pdICCProfile

'JP2 files require extraordinary amounts of memory to decode (4-bytes per channel, so 16-bytes per pixel for RGBA!)
' but we can request downsampling during the load process to "salvage" extraordinarily huge images that are
' otherwise unworkable on 32-bit systems.
Private Const MAX_SIZE_COMPONENTS As Long = 120000000
Private m_SizeReduction As Long

'Caches for export; generating some OpenJPEG objects is resource-intensive, so any steps we can skip
' on back-to-back calls (e.g. when previewing export quality) is beneficial.  Free these via FreeJp2Caches().
Private m_OpjExportImg As Long, m_exportColorDepth As Long

'Set to TRUE when writing a JP2 file; FALSE when reading.  This value affects how streams handle "skip" functions;
' reading simply skips, while writing writes null-bytes.
Private m_writeMode As Boolean
    
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

'This initialization function must be called before any OpenJPEG functions are used.
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

'Manually free OpenJPEG via this function
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

'Verify JPEG-2000 file signature.  Doesn't require OpenJPEG.  Obviously requires read access on the target file.
Public Function IsFileJP2(ByRef srcFile As String, Optional ByRef outCodecFormat As OPJ_CODEC_FORMAT) As Boolean

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
        
        'Open an I/O stream on the target file (read-only)
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
            'Grab the first 12 bytes
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
    m_writeMode = False
    
    'Failsafe check; this function is pointless if OpenJPEG doesn't exist
    If (Not Plugin_OpenJPEG.IsOpenJPEGEnabled()) Then Exit Function
    
    'Failsafe check; validate file signature (hopefully the caller did this, but you never know)
    Dim srcCodec As OPJ_CODEC_FORMAT
    If (Not Plugin_OpenJPEG.IsFileJP2(srcFile, srcCodec)) Then Exit Function
    
    'Still here?  This file passed basic JP2 validation.
    
    'JPEG-2000 images can be extremely large.  PD will attempt to load all images at their defined size,
    ' but if an image fails due to memory constraints, we will try again at 1/4 size.  This process
    ' will happen 3x (reducing dimensions by another 75% each time) before we give up entirely.
    '
    '(And yes, I use GoTo to achieve this.)
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
    ' Start by populating a local parameter struct with the library's current default values.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    'If you want to set custom decoding parameters, do it here.
    ' (For now, PD mostly uses default decoding params.)
    
    'NOTE: to reduce an excessively large image, you can set a reduction factor here (2 ^ n).
    ' This can salvage large images that won't otherwise load on 32-bit systems.
    ' We only apply this if this function has failed once already due to excessive image dimensions.
    If (m_SizeReduction <> 1) Then dParams.cp_reduce = m_SizeReduction \ 2
    
    'Load our decoding parameters into the decoder object
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    Dim retOpj As Long
    retOpj = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOpj = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    'Decoders can operate in a "strict" mode, where incomplete or broken JP2 streams are disallowed.
    '
    'Conversely, "not-strict" mode tells the decoder to decode as much as they can, and stop when they
    ' reach EOF or another meaningful marker - this allows *some* files to be partially recovered,
    ' and testing shows a fair amount of in-the-wild images require strict turned OFF to work.
    '
    'As such, PD currently operates in "not-strict" mode.
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
    
    'Prep a generic OpenJPEG-specific I/O "stream" (read-only) against the target file.
    ' This generic object will call *our* I/O functions for all behaviors, but it's still required as a
    ' parameter to OpenJPEG-specific read functions.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    'OpenJPEG's built-in file stream object doesn't support Unicode chars on Windows,
    ' so we must handle I/O manually.
    
    'Open a pdStream object on the target file.  (Buffer size doesn't matter here; OpenJPEG will request blocks
    ' in its own preferred size, and the pdStream class handles their chunk requests automatically.)
    Set m_Stream = New pdStream
    If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        InternalError FUNC_NAME, "couldn't start pdStream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "pdStream initialized OK..."
    
    'Start a blank OpenJPEG memory stream.  Again, this stream object won't actually touch the file -
    ' it'll simply copy over whatever file chunks *we* supply.)
    Dim pStream As Long
    pStream = opj_stream_default_create(1&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't start blank jp2 stream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Blank jp2 stream initialized OK..."
    
    'Pass our local I/O callbacks to OpenJPEG.  It will assign these to the generic stream object we created earlier.
    ' (Note that this requires a custom-built version of OpenJPEG with manually added support for stdcall callbacks.)
    
    'NOTE: testing shows that even if you don't use the optional user-data parameter, you *must* pass file length
    ' to the "set_user_data_length" function because OpenJPEG assumes that to be the size of the target file.
    ' (Not doing this still works, but OpenJPEG throws spurious warnings and fails in Strict mode.)
    opj_stream_set_user_data pStream, 0&, 0&
    Dim actualFileLen As Currency
    VBHacks.PutMem4 VarPtr(actualFileLen) + 4, Files.FileLenW(srcFile)
    opj_stream_set_user_data_length pStream, actualFileLen
    
    opj_stream_set_read_function pStream, AddressOf JP2_ReadProcDelegate
    opj_stream_set_write_function pStream, AddressOf JP2_WriteProcDelegate
    opj_stream_set_skip_function pStream, AddressOf JP2_SkipProcDelegate
    opj_stream_set_seek_function pStream, AddressOf JP2_SeekProcDelegate
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "I/O callbacks assigned OK..."
    
    'With all I/O set up, we can now attempt to read the target file header.
    ' (From this point forward, OpenJPEG will call our I/O callbacks as necessary to grab file bits;
    '  review the debug log to see specifics.)
    Dim pImage As Long
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Header read successfully"
    
    'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it
    ' and conveniently access relevant struct members.
    '
    'NOTE: DO NOT RELY ON THESE INITIAL HEADER VALUES FOR ANYTHING BUT "OH, THAT'S INTERESTING" PURPOSES.
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
    ' are a concern - but we CAN'T ACTUALLY DO THAT YET because the image's size might change post-component-decoding.
    '
    'And yes, I'm extremely frustrated by this baffling behavior, because it's unintuitive and caused me a ton
    ' of frustration solving inexplicable crashes caused by OpenJPEG internally changing image descriptors after
    ' the header had already been returned and used locally.
    VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), pImage, LenB(m_jp2Image)
    If JP2_DEBUG_VERBOSE Then
        PDDebug.LogAction "Initial header read (DO NOT USE YET; FOR REFERENCE ONLY):"
        PDDebug.LogAction m_jp2Image.x0 & ", " & m_jp2Image.y0 & ", " & m_jp2Image.x1 & ", " & m_jp2Image.y1 & ", " & m_jp2Image.numcomps & " " & GetNameOfOpjColorSpace(m_jp2Image.color_space) & " components"
    End If
    
    'If the image is too large to load on this system, reduce size by half (in each dimension) and try again.
    ' (We'll repeat this up to 3 times before giving up and abandoning loading entirely.)
    If ((m_jp2Image.x1 * m_jp2Image.y1 * m_jp2Image.numcomps) \ m_SizeReduction > MAX_SIZE_COMPONENTS) And (m_SizeReduction < 4) Then
        
        m_SizeReduction = m_SizeReduction * 2
        
        'Free any open objects before attempting a new load operation
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
    
    'Read the rest of the file, including all components e.g. pixel data.
    ' (Expect multiple I/O callbacks here.)
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
    
    'Only NOW can we actually rely on the image header, because it may have changed from our initial access.
    
    'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it more easily.
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
    
    'With image data fully parsed, we can now pull channel data from the supplied component pointer(s).
    
    'Handling varies by component count.  PD can handle 1, 3, 4-channel data OK; for higher counts, let's just grab
    ' the first 1/3/4 (depending on color model) components and use them as-is.
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
        
        'Pull *all* components into local component structs that we can easily traverse via VB6 code
        Dim i As Long
        For i = 0 To numComponents - 1
        
            cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
            VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), cStream.ReadBytes_PointerOnly(sizeOfChannel), sizeOfChannel
            
            'Components often have unpredictable behavior, and these required a *lot* of debugging to solve.
            ' If users encounter problems on their own images, I need to see per-component info to debug.
            If JP2_DEBUG_VERBOSE Then
                PDDebug.LogAction "Channel #" & CStr(i + 1) & " info: "
                With imgChannels(i)
                    PDDebug.LogAction .x0 & ", " & .y0 & ", " & .w & ", " & .h & ", " & .prec & ", " & .Alpha
                    PDDebug.LogAction .p_data & ", " & .dx & ", " & .dy & ", " & .factor & ", " & .sgnd
                End With
            End If
        
        Next i
    
    'The only time stream construction would fail is if we're passed bad (null) component pointers by OpenJPEG
    Else
        PDDebug.LogAction "Failed to initialize stream against component pointer."
        GoTo SafeCleanup
    End If
    
    'We are done with that temporary stream object; release it to free memory for the next (high-consumption) step.
    cStream.StopStream True
    
    'With channel headers assembled, we now need to iterate channels and copy their contents into a local image object.
    
    'First, figure out what color space to use for the embedded data.  This is necessary because the JP2 designers
    ' (insanely) default to "unknown" as a descriptor for key fields like "what color space does this image use".
    ' This forces us to make hard decisions (or as GIMP decided, query the user at load-time) about what to
    ' do with "unknown" data, because a huge fraction of wild JP2 images use "unknown" despite using absolutely
    ' normal color spaces like RGBA.
    '
    'To other designers: DO NOT ALLOW "UNKNOWN" AS AN OFFICIAL VALUE IN YOUR SPEC.  THIS IS STUPID AND DEFEATS
    ' THE WHOLE POINT OF A SPECIFICATION.
    '
    'Anyway, PD handles this as its own step because it's unnecessarily complex.
    Dim targetColorSpace As OPJ_COLOR_SPACE
    targetColorSpace = DetermineColorHandling(m_jp2Image.color_space, numComponents, imgChannels)
    If (targetColorSpace <> OPJ_CLRSPC_GRAY) And (targetColorSpace <> OPJ_CLRSPC_SRGB) And (targetColorSpace <> OPJ_CLRSPC_SYCC) Then
        PDDebug.LogAction "Unknown color space or component count.  Abandoning load."
        GoTo SafeCleanup
    End If
    
    'JP2 images allow each component to have their own dimensions.  At load-time, the caller is responsible for
    ' normalizing these.  (This is handled differently by nearly every JP2 library, because the spec doesn't
    ' formally define behavior like "what resampling algorithm to use" or "how/when to round values", etc.)
    '
    'Because subsampling can impose a significant performance hit, PD tracks it as on a per-component basis,
    ' with alternate (much faster) load pathways for non-subsampled channels.  This provides a large speed
    ' improvement over e.g. FreeImage, which does a ton of coordinate math on every pixel access, even when
    ' subsampling isn't used.
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
    '  because all images in the official conformance suite match this implementation, but images with
    '  extremely large dimensions could require double-precision for full accuracy.  IDGAF about images
    '  like that in a 32-bit codebase, though)
    Dim ssRoundingFactor As Single
    ssRoundingFactor = 0.4999!
    
    'ADDED OCT 2025: if an image only has one channel, ignore the image header's defined dimensions
    ' and instead force the final image to the single component's dimensions.  This matches OpenJPEG's
    ' reference handling of this case.
    If subSamplingActive And (numComponents >= 1) Then
        
        'Calculate indices into each color channel.  (Note that we use RGBA indices, but channels may represent
        ' other color data - the spec foolishly doesn't provide a way to determine canonical color representation,
        ' and my testing shows that 99+% of wild files use standard RGB/YUV component order convention.)
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
        
        'Repeat the above steps for each remaining channel, using each channel's unique subsampled dimensions
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
    ' we can now load pixel data directly into a local RGBA buffer.
    
    'Prep the destination surface.
    Set dstDIB = New pdDIB
    dstDIB.CreateBlank m_OpjNotes.finalWidth, m_OpjNotes.finalHeight, 32, RGB(255, 255, 255), 255
    
    Dim targetWidth As Long, targetHeight As Long
    Dim channelSizeEstimate As Long
    Dim srcRs() As Long, srcGs() As Long, srcBs() As Long, srcAs() As Long
    Dim srcRSA As SafeArray1D, srcGSA As SafeArray1D, srcBSA As SafeArray1D, srcASA As SafeArray1D
    Dim copyAlpha As Boolean
            
    Dim r As Long, g As Long, b As Long, a As Long, yccY As Long, yccB As Long, yccR As Long
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    Dim saOffset As Long, hdrDivisor As Long
    
    'Data in JP2 files can be signed, meaning that e.g. 8-bit data is represented as [-127, 128] instead of [0, 255].
    ' PD handles this case successfully but requires additional variables to convert to unsigned types.
    Dim rIsSigned As Boolean, gIsSigned As Boolean, bIsSigned As Boolean, aIsSigned As Boolean
    Dim rSgnComp As Long, gSgnComp As Long, bSgnComp As Long, aSgnComp As Long
    
    'Unlike other image format libraries, OpenJPEG always loads channels as 4-byte ints (Longs in VB6)
    ' regardless of embedded color depth.  This is incredibly wasteful from a memory standpoint,
    ' but it does simplify handling of various bit-depths, because the source channel data is always the
    ' same size per-pixel, regardless of how it was actually encoded in the target file.
    targetWidth = m_OpjNotes.finalWidth
    targetHeight = m_OpjNotes.finalHeight
    channelSizeEstimate = targetWidth * targetHeight * 4    '(See above note - this line is not a typo!)
    
    Dim finalPrec As Long
    finalPrec = imgChannels(0).prec
    m_OpjNotes.finalPrecision = finalPrec
    
    'Load pixel data, with handling separated by color type
    If (targetColorSpace = OPJ_CLRSPC_GRAY) Then
        
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Importing image using grayscale decoder..."
        
        'Handle signed data correctly by pre-calculating a normalization offset
        gIsSigned = (imgChannels(0).sgnd <> 0)
        If gIsSigned Then rSgnComp = (2 ^ imgChannels(0).prec) \ 2 Else rSgnComp = 0
        
        'Wrap a temporary array around OpenJPEG's copy of the channel
        VBHacks.WrapArrayAroundPtr_Long srcRs, srcRSA, imgChannels(0).p_data, channelSizeEstimate
        
        'Precision can technically be any value between 1 and ???? (upper limit is unclear from the spec).
        ' All images from the official conformance spec are handled correctly, but PD has only been tested
        ' on color depths up-to-24-bit.  32-bit unsigned data may break due to a lack of an unsigned type
        ' in VB6.  (All other precision+signed combinations are expected to work!)
        
        'Sub-8pp channels need to be upsampled
        If (finalPrec < 8) Then
            
            'Reusing variable names is stupid, but here we are!  This value is multiplied by the sub-8-bit component value
            ' to arrive at a value on the range [0, 255].
            '
            'Well, TECHNICALLY it won't be on the range [0, 255] because e.g. a 4-bit image will go from
            ' [0, 15] to [0, 240]. I only do it this way because that's what the official OpenJPEG library does,
            ' and they're wrong but we need to mimic their behavior for consistency.
            hdrDivisor = 2 ^ (8 - finalPrec)
        
        'Normal 8-bpp channels require no extra handling
        'ElseIf (imgChannels(0).prec = 8) Then
        
        'HDR data needs to be downsampled to 8-bpp for PD.
        ' (TODO at some future date: in the absence of an ICC profile, allow the user to tone-map as they wish?)
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
                ' we split handling.  Branch prediction in modern CPUs effectively makes this "free".
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
                
                'Assign all destination channels to the calculated gray
                dstPixels(x * 4) = g
                dstPixels(x * 4 + 1) = g
                dstPixels(x * 4 + 2) = g
                
            Next x
            
        'Proceed to next line
        Next y
        
        'Unwrap all unsafe array wrappers before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        dstDIB.UnwrapArrayFromDIB dstPixels
        dstDIB.SetAlphaPremultiplication True
        
        'Load complete!  (Clean-up is still required, however.)
        LoadJP2 = True
        
        'End grayscale handling
    
    'Color and YCC spaces are handled together, because indexing is identical - only the s/eYCC conversion varies
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
                    ' we split handling.  Branch prediction in modern CPUs effectively makes this "free".
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
                    
                    'Safety against malformed data
                    If (b < 0) Then b = 0
                    If (b > 255) Then b = 255
                    
                    If (g < 0) Then g = 0
                    If (g > 255) Then g = 255
                    
                    If (r < 0) Then r = 0
                    If (r > 255) Then r = 255
                    
                    'Final assignment into the destination buffer
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
            ' TODO: find eYCC images and test conversion; it will need different conversion matrices,
            ' but the conformance suite doesn't use that format so I'm SOL for testing currently.
            ElseIf (targetColorSpace = OPJ_CLRSPC_SYCC) Or (targetColorSpace = OPJ_CLRSPC_EYCC) Then
                
                For x = 0 To targetWidth - 1
                    
                    'For detailed comments, see RGB/A section above.
                    
                    'Subsampling
                    If subSamplingActive Then
                        yccR = srcBs(bSSPosY(y) + bSSPosX(x))
                        yccB = srcGs(gSSPosY(y) + gSSPosX(x))
                        yccY = srcRs(rSSPosY(y) + rSSPosX(x))
                    Else
                        yccR = srcBs(saOffset + x)
                        yccB = srcGs(saOffset + x)
                        yccY = srcRs(saOffset + x)
                    End If
                    
                    'Signed vs unsigned
                    yccY = yccY + bSgnComp
                    yccB = yccB + gSgnComp
                    yccR = yccR + rSgnComp
                    
                    'Up-sample low-precision / down-sample high-precision
                    If (finalPrec < 8) Then
                        yccY = yccY * hdrDivisor
                        yccB = yccB * hdrDivisor
                        yccR = yccR * hdrDivisor
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
        
        'Unwrap all unsafe temporary arrays before exiting
        VBHacks.UnwrapArrayFromPtr_Long srcRs
        VBHacks.UnwrapArrayFromPtr_Long srcGs
        VBHacks.UnwrapArrayFromPtr_Long srcBs
        If copyAlpha Then VBHacks.UnwrapArrayFromPtr_Long srcAs
        dstDIB.UnwrapArrayFromDIB dstPixels
        
        'Load complete!  (Clean-up is still required, however.)
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "JP2 pixel data processed successfully!"
        LoadJP2 = True
        
    End If
    
    'With the target surface successfully constructed, we can apply color management (if an embedded color profile exists).
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
                ' Instead, note that it is already sRGB.
                profHash = ColorManagement.GetSRGBProfileHash()
                dstDIB.SetColorProfileHash profHash
                
            End If
        End If
    
    End If
    
    'Premultiply alpha before exiting
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

'Return component count of last-loaded image
Public Function GetComponentCountOfLastImage() As Long
     GetComponentCountOfLastImage = m_OpjNotes.numComponents
End Function

'Return precision (bits-per-largest-component) of last-loaded image.
Public Function GetPrecisionOfLastImage() As Long
    GetPrecisionOfLastImage = m_OpjNotes.finalPrecision
End Function

'Perform a max-speed decode from an open pdStream (*MUST* be opened and set to the desired offset by the caller) to a pdDIB object.
' Does *not* perform extra validations, and does *not* color-manage the result.  sRGB is assumed.
' (PhotoDemon uses this function internally to generate export quality previews; color-management has already occurred.)
Public Function FastDecodeFromStreamToDIB(ByRef srcStream As pdStream, ByRef dstDIB As pdDIB) As Boolean

    Const FUNC_NAME As String = "FastDecodeFromStreamToDIB"
    FastDecodeFromStreamToDIB = False
    m_writeMode = False
    
    'Failsafe check; this function is pointless if OpenJPEG doesn't exist
    If (Not Plugin_OpenJPEG.IsOpenJPEGEnabled()) Then Exit Function
    
    'Reset any module-level items from previous JPEG-2000 interactions.
    ' Note that - by design - this function will *not* fill color management structs.
    Erase m_IccBytes
    m_IccLength = 0
    Set m_ColorProfile = Nothing
    
    'Point the module-level stream object at the passed stream, and reset the stream to its start (just in case)
    Set m_Stream = srcStream
    m_Stream.SetPosition 0&, FILE_BEGIN
    
    'This function does *not* support automatic size reduction on OOM errors
    m_SizeReduction = 1
    
    'Initialize a default JP2-format JPEG-2000 decoder (the only codec supported by this function, by design).
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder itself..."
    Dim pDecoder As Long
    pDecoder = opj_create_decompress(OPJ_CODEC_JP2)
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
    
    'Populate a parameter struct with default values.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Prepping decoder params..."
    Dim dParams As opj_dparameters
    opj_set_default_decoder_parameters VarPtr(dParams)
    
    'Load decoding parameters into the decoder object
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing decoder against params..."
    Dim retOpj As Long
    retOpj = opj_setup_decoder(pDecoder, VarPtr(dParams))
    If (retOpj = 0) Then
        InternalError FUNC_NAME, "failed to set up decoder"
        GoTo SafeCleanup
    End If
    
    'Decoders can use a "strict" mode, where incomplete or broken JP2 streams are simply disallowed.
    ' This mode doesn't really matter for this function, but we'll use non-strict to match PD's regular load path.
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
    
    'Prep a generic OpenJPEG "stream" (read-only) against the target file.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Creating default file stream..."
    
    Dim pStream As Long
    pStream = opj_stream_default_create(1&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't start blank jp2 stream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Blank jp2 stream initialized OK..."
    
    'Use our local I/O callbacks instead of OpenJPEG's.
    ' (Note that this requires a custom-built version of OpenJPEG with manually added support for stdcall callbacks.)
    
    'NOTE: after testing, we *do* need to call these functions even if we don't use them.  OpenJPEG will error
    ' randomly if the length of the user data is not set to a non-zero value prior to reading the source stream
    opj_stream_set_user_data pStream, 0&, 0&
    Dim actualFileLen As Currency
    VBHacks.PutMem4 VarPtr(actualFileLen) + 4, m_Stream.GetStreamSize()
    opj_stream_set_user_data_length pStream, actualFileLen
    
    opj_stream_set_read_function pStream, AddressOf JP2_ReadProcDelegate
    opj_stream_set_write_function pStream, AddressOf JP2_WriteProcDelegate
    opj_stream_set_skip_function pStream, AddressOf JP2_SkipProcDelegate
    opj_stream_set_seek_function pStream, AddressOf JP2_SeekProcDelegate
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "I/O callbacks assigned OK..."
    
    'With the stream set up, we can now attempt to read the actual source data.
    ' (From this point forward, OpenJPEG will call our I/O callbacks as necessary to grab file bits.)
    
    'Read the header
    Dim pImage As Long
    If (opj_read_header(pStream, pDecoder, VarPtr(pImage)) <> 1&) Then
        InternalError FUNC_NAME, "failed to read header"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Header read successfully"
    
    'Decode component (pixel) data.  (Expect multiple I/O callbacks here.)
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
    
    'Only NOW can we actually rely on the header's contents, because it may have changed post-initial-load.
    
    'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it
    ' and conveniently access relevant struct members.
    VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), pImage, LenB(m_jp2Image)
    If JP2_DEBUG_VERBOSE Then
        PDDebug.LogAction "Final header values:"
        PDDebug.LogAction m_jp2Image.x0 & ", " & m_jp2Image.y0 & ", " & m_jp2Image.x1 & ", " & m_jp2Image.y1 & ", " & m_jp2Image.numcomps & " " & GetNameOfOpjColorSpace(m_jp2Image.color_space) & " components"
    End If
    
    'With the image data fully parsed, we can now pull channel data from the supplied component pointer(s).
    
    'Handling varies by component count.  PD can handle 1, 3, 4-channel data OK; for higher counts, it'll just grab
    ' the first 1/3/4 (depending on color model) and use them as-is.
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
    
    'First, figure out what color space to use top interpret the source components
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
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank m_OpjNotes.finalWidth, m_OpjNotes.finalHeight, 32, RGB(255, 255, 255), 255
    
    Dim targetWidth As Long, targetHeight As Long
    Dim channelSizeEstimate As Long
    Dim srcRs() As Long, srcGs() As Long, srcBs() As Long, srcAs() As Long
    Dim srcRSA As SafeArray1D, srcGSA As SafeArray1D, srcBSA As SafeArray1D, srcASA As SafeArray1D
    Dim copyAlpha As Boolean
            
    Dim r As Long, g As Long, b As Long, a As Long, yccY As Long, yccB As Long, yccR As Long
    Dim dstPixels() As Byte, dstSA As SafeArray1D
    Dim saOffset As Long, hdrDivisor As Long
    
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
        FastDecodeFromStreamToDIB = True
        
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
        FastDecodeFromStreamToDIB = True
        
    End If
    
    'For now cleanup and exit
    GoTo SafeCleanup
    
    Exit Function
    
    'Code beyond this point performs a full clean-up of all internal and external resources for the current jp2 image
SafeCleanup:
    If (Not m_Stream Is Nothing) Then m_Stream.SetPosition 0&, FILE_BEGIN
    Set m_Stream = Nothing
    
    If (Not cStream Is Nothing) Then cStream.StopStream True
    If (pImage <> 0) Then opj_image_destroy pImage
    If (pStream <> 0) Then opj_stream_destroy pStream
    If (pDecoder <> 0) Then opj_destroy_codec pDecoder
    
End Function

'Save a pdDIB object to an arbitrary pdStream object.  This provides flexibility in saving to file vs saving to memory,
' since OpenJPEG relies on us to supply our own I/O stream implementation anyway.
'
'FOR THIS TO WORK, THE STREAM MUST BE OPEN AND INITIALIZED **BEFORE** CALLING this function.
' This function will fail otherwise, because it doesn't know where you want the JP2 saved.
'
'ALSO: for performance reasons, this function creates (potentially large) module-level caches for storing
' original DIB pixel data, because everything has to be translated to 32-bit channels (128-bit RGBA pixels)
' prior to encoding via OpenJPEG.  After this function wraps, you MUST call FreeJp2Caches() to reclaim that
' memory.  PD implements it this way because export previews reuse the 32-bit channels between calls,
' preventing memory thrashing and greatly improving performance on low-end PCs.
'
'outputColorDepth must be one of three values:
' - 8 (grayscale, no alpha)
' - 24 (RGB, no alpha)
' - 32 (RGBA)
Public Function SavePdDIBToJp2Stream(ByRef srcDIB As pdDIB, ByRef dstStream As pdStream, ByVal saveQuality As Long, Optional ByVal outputColorDepth As Long = 32, Optional ByVal forceNewImageObject As Boolean = False) As Boolean
    
    Const FUNC_NAME As String = "SavePdDIBToJp2Stream"
    SavePdDIBToJp2Stream = False
    m_writeMode = True
    
    'Exit immediately if the destination stream isn't open and initialized
    If (dstStream Is Nothing) Then
        InternalError FUNC_NAME, "stream must be initialized"
        Exit Function
    End If
    
    If (Not dstStream.IsOpen()) Then
        InternalError FUNC_NAME, "stream must be open"
        Exit Function
    End If
    
    'Initialize a default set of encoding parameters
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Retrieving default jp2 params..."
    Dim srcParams As opj_cparameters
    opj_set_default_encoder_parameters VarPtr(srcParams)
    
    'To maximize compatibility, PD only saves single-layer JP2 images with minimal deviations from default behavior
    srcParams.tcp_numlayers = 1
    
    'Up to 512 can be used as the "compression" value, but PD currently only provides a range of [1, 256]
    ' to the user (because above ~256 compression artifacts become egregious).
    srcParams.tcp_rates(0) = CSng(CLng(saveQuality And &H3FF&))
    
    'Tell the library to use the above-defined rate as primary quality determiner
    srcParams.cp_disto_alloc = 1
    
    'A separate chunk of the compression settings applies to command-line parameters only,
    ' so specifying a magic number for JP2 format shouldn't be necessary here... but OpenJPEG
    ' has a lot of hidden behaviors so I'd prefer to err on the side of "better safe than sorry".
    srcParams.cod_format = 1
    
    'Initialize an array of component parameters (one per component).
    Dim numParams As Long
    If (outputColorDepth <= 8) Then
        numParams = 1
    ElseIf (outputColorDepth <= 24) Then
        numParams = 3
    Else
        numParams = 4
    End If
    
    'If the export color depth changes between calls, we need to generate a new backing image object
    If (m_exportColorDepth <> numParams) Then forceNewImageObject = True
    m_exportColorDepth = numParams
    
    'RGB images can use an MCT (multi-component transform) for additional file size savings.
    ' (Note that this setting only works on RGBA images if we *explicitly* mark the alpha channel flag for the correct component.)
    If (numParams >= 3) Then srcParams.tcp_mct = 1 Else srcParams.tcp_mct = 0
    
    'Next, we need to prep an OpenJPEG-specific image object.  PD will attempt to reuse the last-created image object
    ' unless explicitly told otherwise.  (This step is very expensive, so skipping it is advisable whenever possible,
    ' particularly when previewing export quality.)
    If forceNewImageObject Or (m_OpjExportImg = 0) Then
        
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Creating new jp2 backing imag..."
        
        'Free the previous image, if any
        If (m_OpjExportImg <> 0) Then
            opj_image_destroy m_OpjExportImg
            m_OpjExportImg = 0
        End If
        
        'Populate a list of image component headers with desired encoding settings
        Dim cmpParams() As opj_image_comptparm
        ReDim cmpParams(0 To numParams - 1) As opj_image_comptparm
        
        'Populate all component parameter values
        Dim i As Long
        For i = 0 To numParams - 1
            With cmpParams(i)
                
                'Subsampling is not currently supported by PD at export-time
                .dx = srcParams.subsampling_dx
                .dy = srcParams.subsampling_dy
                .w = srcDIB.GetDIBWidth
                .h = srcDIB.GetDIBHeight
                
                'Precision is currently locked at 8-bits-per-channel
                .prec = 8
                .opj_bpp = .prec    'BPP is deprecated; only .prec matters in modern OpenJPEG builds
                
                'PD only writes unsigned data
                .sgnd = 0
                
            End With
        Next i
        
        'PD only exports grayscale or sRGB images at present
        Dim dstColorSpace As OPJ_COLOR_SPACE
        If (numParams = 1) Then
            dstColorSpace = OPJ_CLRSPC_GRAY
        Else
            dstColorSpace = OPJ_CLRSPC_SRGB
        End If
        
        'We now have everything we need to initialize an OpenJPEG image object
        m_OpjExportImg = opj_image_create(numParams, VarPtr(cmpParams(0)), dstColorSpace)
        If (m_OpjExportImg = 0) Then
            InternalError FUNC_NAME, "opj_image_create failed"
            GoTo SafeCleanup
        End If
        
        'Next, we need to manually set image size inside the image object.
        ' (Because we can't easily alias m_OpjExportImage as an opj_image object,
        '  we're just gonna set these values manually, in x0, y0, x1, y1 order)
        VBHacks.PutMem4 m_OpjExportImg, 0&
        VBHacks.PutMem4 m_OpjExportImg + 4, 0&
        VBHacks.PutMem4 m_OpjExportImg + 8, srcDIB.GetDIBWidth
        VBHacks.PutMem4 m_OpjExportImg + 12, srcDIB.GetDIBHeight
        
        'Next we need to populate pixel channels.  OpenJPEG has already allocated memory for each channel,
        ' but we (obviously) have to fill them.
        
        'First, retrieve the custom opj structs that actually store component information.
        Dim imgChannels() As opj_image_comp
        ReDim imgChannels(0 To numParams - 1) As opj_image_comp
    
        'The target struct has non-aligned members (entries that don't align cleanly on 4-byte boundaries),
        ' so we need to account for this when memcpy'ing them into local structs.
        Dim sizeOfChannel As Long, sizeOfChannelAligned As Long
        sizeOfChannelAligned = LenB(imgChannels(0))
        sizeOfChannel = Len(imgChannels(0))
        
        'Copy the header struct from OpenJPEG's handle to our own local struct, so we can traverse it
        ' and conveniently access relevant struct members.
        VBHacks.CopyMemoryStrict VarPtr(m_jp2Image), m_OpjExportImg, LenB(m_jp2Image)
        
        'To simplify reading data from arbitrary pointers, use a pdStream object.
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, startingBufferSize:=sizeOfChannelAligned * numParams, baseFilePointerOffset:=m_jp2Image.pComps, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Pull *all* components into local structs that we can easily traverse via VB6 code
            For i = 0 To numParams - 1
            
                cStream.SetPosition i * sizeOfChannelAligned, FILE_BEGIN
                
                Dim pSrc As Long
                pSrc = cStream.ReadBytes_PointerOnly(sizeOfChannel)
                VBHacks.CopyMemoryStrict VarPtr(imgChannels(i)), pSrc, sizeOfChannel
                
                'If we're writing a 32-bit image, manually flag the alpha channel now, and write the modified
                ' component header back out to its original location (owned by OpenJPEG)
                If (i = 3) Then
                    imgChannels(i).Alpha = 1
                    VBHacks.CopyMemoryStrict pSrc, VarPtr(imgChannels(i)), sizeOfChannel
                End If
                
            Next i
        
        'The only time stream construction would fail is if we're passed bad (null) pointers by OpenJPEG
        Else
            PDDebug.LogAction "Failed to initialize stream against component pointer."
            GoTo SafeCleanup
        End If
        
        'We are done with the temporary stream object
        cStream.StopStream True
        
        'To simplify our life, wrap VB arrays (ints!) around the component pointers
        Dim dstR() As Long, dstG() As Long, dstB() As Long, dstA() As Long
        Dim dstSaR As SafeArray1D, dstSaG As SafeArray1D, dstSaB As SafeArray1D, dstSaA As SafeArray1D
        If (numParams = 1) Then
            VBHacks.WrapArrayAroundPtr_Long dstR, dstSaR, imgChannels(0).p_data, imgChannels(0).w * imgChannels(0).h * 4
        Else
            VBHacks.WrapArrayAroundPtr_Long dstR, dstSaR, imgChannels(0).p_data, imgChannels(0).w * imgChannels(0).h * 4
            VBHacks.WrapArrayAroundPtr_Long dstG, dstSaG, imgChannels(1).p_data, imgChannels(1).w * imgChannels(1).h * 4
            VBHacks.WrapArrayAroundPtr_Long dstB, dstSaB, imgChannels(2).p_data, imgChannels(2).w * imgChannels(2).h * 4
            If (numParams > 3) Then VBHacks.WrapArrayAroundPtr_Long dstA, dstSaA, imgChannels(3).p_data, imgChannels(3).w * imgChannels(3).h * 4
        End If
        
        'Copy pixel data
        Dim x As Long, y As Long, idxDst As Long, idxSrc As Long
        Dim srcPx() As Byte, srcSA As SafeArray1D
        
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = srcDIB.GetDIBWidth
        srcHeight = srcDIB.GetDIBHeight
        For y = 0 To srcHeight - 1
            srcDIB.WrapArrayAroundScanline srcPx, srcSA, y
        For x = 0 To srcWidth - 1
            idxSrc = x * 4
            If (numParams = 1) Then
                dstR(idxDst) = srcPx(idxSrc)
            Else
                dstB(idxDst) = srcPx(idxSrc)
                dstG(idxDst) = srcPx(idxSrc + 1)
                dstR(idxDst) = srcPx(idxSrc + 2)
                If (numParams > 3) Then dstA(idxDst) = srcPx(idxSrc + 3)
            End If
            idxDst = idxDst + 1
        Next x
        Next y
        
        'Unwrap all unsafe arrays.  (Note that unwrapping an uninitialized array is fine; we just overwrite the pointer with 0&.)
        srcDIB.UnwrapArrayFromDIB srcPx
        VBHacks.UnwrapArrayFromPtr_Long dstR
        VBHacks.UnwrapArrayFromPtr_Long dstB
        VBHacks.UnwrapArrayFromPtr_Long dstG
        VBHacks.UnwrapArrayFromPtr_Long dstA
        
    Else
        If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Reusing old jp2 image object..."
    End If
    
    'With the OpenJPEG-specific image object prepped, we are finally ready to encode the data in JP2 format
    
    'Prep an encoder.  Note that we have multiple options here - naked j2k codestreams are unsupported in PD
    ' (they have extreme limitations when loading, like not defining component types) so we explicitly write
    ' full JP2 images only.
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing jp2 encoder..."
    Dim hEncoder As Long
    hEncoder = opj_create_compress(OPJ_CODEC_JP2)
    If (hEncoder = 0) Then
        InternalError FUNC_NAME, "opj_create_compress failed"
        GoTo SafeCleanup
    End If
    
    'Before doing anything with the encoder, we must assign the callbacks we'll use to write image data
    ' to memory/file (via the pdStream object we were passed).
    Set m_Stream = dstStream
    
    'Initialize local I/O functions for our constructed decoder.
    ' (Note that this requires a custom-built version of OpenJPEG with manually added support for stdcall callbacks.)
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Initializing callbacks..."
    opj_set_info_handler hEncoder, AddressOf HandlerInfo, 0&
    opj_set_warning_handler hEncoder, AddressOf HandlerWarning, 0&
    opj_set_error_handler hEncoder, AddressOf HandlerError, 0&
    
    'Initialize a blank OpenJPEG memory stream.  Again, this stream object won't actually touch the file -
    ' it'll simply hand chunks of file data to *us* and *we* must write them to our stream.)
    Dim pStream As Long
    pStream = opj_stream_default_create(0&)
    If (pStream = 0&) Then
        InternalError FUNC_NAME, "couldn't start blank jp2 stream"
        GoTo SafeCleanup
    End If
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Blank jp2 stream initialized OK..."
    opj_stream_set_read_function pStream, AddressOf JP2_ReadProcDelegate
    opj_stream_set_write_function pStream, AddressOf JP2_WriteProcDelegate
    opj_stream_set_skip_function pStream, AddressOf JP2_SkipProcDelegate
    opj_stream_set_seek_function pStream, AddressOf JP2_SeekProcDelegate
    
    'With everything initialized, we can now initialize the encoder with our backing image
    ' and any associated encoding parameters
    If (opj_setup_encoder(hEncoder, VarPtr(srcParams), m_OpjExportImg) = 0) Then
        InternalError FUNC_NAME, "failed to set up encoder"
        GoTo SafeCleanup
    End If
    
    'Time to encode the image!
    
    'Attempt to start encoding.  This step will only fail (typically) if you supply bad encoding and/or image parameters
    If (opj_start_compress(hEncoder, m_OpjExportImg, pStream) = 0) Then
        InternalError FUNC_NAME, "opj_start_compress failed"
        GoTo SafeCleanup
    End If
    
    'Perform the rest of the encoding.  This will hand finished bytes over to our delegate I/O functiond
    ' as they're encoded (typically in 1 MB chunks).
    If (opj_encode(hEncoder, pStream) = 0) Then
        InternalError FUNC_NAME, "opj_encode failed"
        GoTo SafeCleanup
    End If
    
    'Explicitly end compression.  (Note that this may require skipping around in the target stream
    ' to write some final markers and lengths - encoding is *not* strictly sequential.)
    If (opj_end_compress(hEncoder, pStream) = 0) Then
        InternalError FUNC_NAME, "opj_end_compress failed"
        GoTo SafeCleanup
    End If
    
    'If we're still here, the stream was written correctly!
    SavePdDIBToJp2Stream = True
    
    'Free everything relevant before exiting.  Note that the backing OpenJPEG-format image is explicitly *not* freed;
    ' the caller must manually free this because we may reuse it between calls (such as when previewing export quality)
    If (pStream <> 0) Then opj_stream_destroy pStream
    If (hEncoder <> 0) Then opj_destroy_codec hEncoder
    
    Set m_Stream = Nothing
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction FUNC_NAME & " successful"
    
    Exit Function

'On failure, attempt as much cleanup as we can, including the cached opj-format image object
SafeCleanup:
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Exiting " & FUNC_NAME & " via SafeCleanup"
    SavePdDIBToJp2Stream = False
    
    If (m_OpjExportImg <> 0) Then
        opj_image_destroy m_OpjExportImg
        m_OpjExportImg = 0
    End If
    If (hEncoder <> 0) Then opj_destroy_codec hEncoder
    If (pStream <> 0) Then opj_stream_destroy pStream
    
    'Free any other caches
    FreeJp2Caches
    Set m_Stream = Nothing
    
End Function

'After saving, *if you don't plan to reuse the source image data*, call this function to free intermediate caches.
' It will reclaim (potentially) very large amounts of memory.
Public Sub FreeJp2Caches()
    If (m_OpjExportImg <> 0) Then
        opj_image_destroy m_OpjExportImg
        m_OpjExportImg = 0
    End If
End Sub

'Figure out how to handle incoming color data.  JPEG-2000 streams are extremely flexible in terms of color components
' (e.g. "undefined" color spaces and infinite component counts are allowed, and each channel is allowed its own
' encoding method and/or grid dimensions via subsampling).  This makes them messy to handle, and a lot of software
' simply doesn't load files encoded with anything but non-subsampled, unsigned 8-bpp RGB.
'
'My goal is to do better than that - not necessarily to cover every possible combination of JP2 file attributes,
' but instead, to make intelligent inferences about unknown data (e.g. three undefined channels are likely RGB,
' four is likely RGBA) and cover as many "real-world" use-cases as I can.
'
'If an obvious correlation with a known color space cannot be made, PD will treat the image data as grayscale and
' load the first channel only.  This typically allows *something* to be recovered from any file.
Private Function DetermineColorHandling(ByVal fileColorSpace As OPJ_COLOR_SPACE, ByVal numComponents As Long, ByRef imgChannels() As opj_image_comp) As OPJ_COLOR_SPACE
    
    'An "unknown" color space notifies the caller that PD is unequipped to handle this image's data.
    ' ("Unknown" is an extremely common state of wild JP2 images, and PD will attempt to reassign that
    '  constant to something useable based on simple heuristics.)
    DetermineColorHandling = OPJ_CLRSPC_UNKNOWN
    
    'Failsafe check for component count (should have been validated by caller)
    If (numComponents <= 0) Then Exit Function
    
    'Some software only checks the size of the first component, and uses that as the size of the image.
    ' PD tries to use the header-defined image size instead (but note that the first channel *can* be
    ' subsampled, unlike other image formats!)
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = m_jp2Image.x1 - m_jp2Image.x0
    targetHeight = m_jp2Image.y1 - m_jp2Image.y0
    
    'If an image is too large for available memory, PD will try to load a downsampled version instead.
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
        ReDim .isChannelSubsampled(0 To numComponents) As Boolean
        ReDim .channelSsWidth(0 To numComponents) As Long
        ReDim .channelSsHeight(0 To numComponents) As Long
        .idxAlphaChannel = -1
    End With
    
    'If there is only one channel in the image, color space doesn't matter - treat it as grayscale.
    ' (Also, to match the reference OpenJPEG implementation, subsampling is ignored and channel
    ' dimensions are forcibly used as final image dimensions, regardless of what size the image header
    ' may have claimed.)
    If (numComponents = 1) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        With m_OpjNotes
            .hasSubsampling = False
            .imgHasAlpha = False
            .idxAlphaChannel = -1
            .isChannelSubsampled(0) = False
            .channelSsWidth(0) = imgChannels(0).w
            .channelSsHeight(0) = imgChannels(0).h
            .finalWidth = .channelSsWidth(0)
            .finalHeight = .channelSsHeight(0)
        End With
        Exit Function
    End If
    
    'If we're still here, this image has multiple channels.  Iterate up to 4 channels and track specific channel data,
    ' including channel dimensions.  (Subsampling in JP2 files means each channel can have independent dimensions,
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
    
    'If components < 3, treat as grayscale
    If (numComponents > 0) And (numComponents < 3) Then
        DetermineColorHandling = OPJ_CLRSPC_GRAY
        
        'Technically we could handle 2-channel data as grayscale+alpha, but the conformance suite never uses this combo
        ' and I don't know if it exists "in the wild".  Since assuming alpha could break otherwise "good" grayscale data,
        ' I've disabled this option pending actual test images.
        m_OpjNotes.imgHasAlpha = False      'Use (m_OpjNotes.idxAlphaChannel >= 0) here to set an actual alpha channel
        If (Not m_OpjNotes.imgHasAlpha) Then m_OpjNotes.numComponents = 1
    
    'CMYK was recently added as a potential JP2 color space, but I have not found any conformance images
    ' using this space so it's currently UNTESTED.  3-4 channels with unknown color-spaces are treated as RGB/A.
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

'Return a human-readable color space name from an OpenJPEG color space constant
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

'Local callbacks for info/warning/error messages
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
' so we must supply our own callbacks and use them for all I/O behavior.
' (As a nice bonus, this also improves performance because we use memory mapped I/O which can
'  greatly improve throughput, especially on modern SSDs.)
'
'Also note: p_user_data is always unused by these functions.
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

'Write [n] bytes to a write-accessible destination stream.  p_user_data is unused.
Private Function JP2_WriteProcDelegate(ByVal p_buffer As Long, ByVal p_nb_bytes As Long, ByVal p_user_data As Long) As Long
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Write requested for " & p_nb_bytes
    
    If (Not m_Stream Is Nothing) Then
        
        'Write [n] bytes to the output stream
        If (m_Stream.WriteBytesFromPointer(p_buffer, p_nb_bytes) <> 0) Then JP2_WriteProcDelegate = p_nb_bytes
        
    End If
    
End Function

'Advance pointer [n] bytes in input file. p_user_data is unused.
Private Function JP2_SkipProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Currency
    
    'PD can't actually use 64-bit values (yet) for file seeks; use only the lower 4 bytes.
    ' (This workaround would not be needed in a 64-bit build.)
    Dim lowerFourSkip As Long
    VBHacks.GetMem4_Ptr VarPtr(p_nb_bytes), VarPtr(lowerFourSkip)
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Skip requested for " & lowerFourSkip & " (" & p_nb_bytes & ")"
    
    If (Not m_Stream Is Nothing) Then
        
        'In *read* mode, simply advance the pointer
        If (Not m_writeMode) Then
            If m_Stream.SetPosition(lowerFourSkip, FILE_CURRENT) Then
                JP2_SkipProcDelegate = p_nb_bytes
            Else
                JP2_SkipProcDelegate = -1
            End If
        
        'In *write* mode, write null values
        Else
            If m_Stream.WritePadding(lowerFourSkip) Then
                JP2_SkipProcDelegate = p_nb_bytes
            Else
                JP2_SkipProcDelegate = -1
            End If
        End If
        
    End If
    
End Function

'Set pointer to [n] bytes from 0 in output file. p_user_data is unused.
Private Function JP2_SeekProcDelegate(ByVal p_nb_bytes As Currency, ByVal p_user_data As Long) As Long

    'PD can't actually use 64-bit values (yet) for file seeks; use only the lower 4 bytes.
    ' (This workaround would not be needed in a 64-bit build.)
    Dim lowerFourSkip As Long
    VBHacks.GetMem4_Ptr VarPtr(p_nb_bytes), VarPtr(lowerFourSkip)
    
    If JP2_DEBUG_VERBOSE Then PDDebug.LogAction "Seek requested for " & lowerFourSkip & ", " & m_Stream.GetStreamSize()
    
    If (Not m_Stream Is Nothing) Then
        If m_Stream.SetPosition(lowerFourSkip, FILE_BEGIN) Then
            JP2_SeekProcDelegate = 1
        Else
            JP2_SeekProcDelegate = 0
        End If
    End If
    
End Function
