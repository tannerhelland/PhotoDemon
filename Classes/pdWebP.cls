VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWebP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'WebP Image Interface
'Copyright 2021-2026 by Tanner Helland
'Created: 22/September/21
'Last updated: 01/October/21
'Last update: add a streaming interface for writing animated WebP files on-the-fly (used by PD's screen recorder)
'
'libWebP is a BSD-licensed WebP image support library.  It is developed and actively maintained by Google.
' Fortunately for PhotoDemon, it comes with a robust C interface and legacy compilation options, enabling
' support all the way back to Windows XP.
'
'PhotoDemon historically used FreeImage to manage WebP files, but using libwebp directly allows for
' significantly better performance and feature support, including import/export support for animated WebP.
'
'Note that all features in this module rely on the libwebp binaries that ship with PhotoDemon (all THREE
' of them).  This class will not work if libwebp or its mux/demux components are missing or broken.
'
'You can freely plug-in your own builds of libwebp and libwebpde/mux, but note that they must be built
' using cdecl exports and x86 only (obviously).  Note also that many exported APIs require us to pass a
' hard-coded ABI identifier which may need to be updated against future library versions (to enable new
' library features, at least).  If you use your own libwebp copies, change the hard-coded ABI constants
' at the top of this class to match.
'
'Please refer to Plugin_WebP for details on initializing the various libraries; that is not handled here.
'
'This wrapper class uses a shorthand implementation of DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To aid debugging, you can activate "verbose" output; this will dump various bits of diagnostic info
' to the debug log.
Private Const WEBP_DEBUG_VERBOSE As Boolean = False

'Some exported APIs require passing ABI version numbers; confusingly, these exports are usually
' suffixed with "Internal"; see the class initialization function for details.
Private Const WEBP_DECODER_ABI_VERSION As Long = &H209& 'From decode.h
Private Const WEBP_DEMUX_ABI_VERSION As Long = &H107&   'From demux.h
Private Const WEBP_ENCODER_ABI_VERSION As Long = &H20F& 'From encode.h
Private Const WEBP_MUX_ABI_VERSION As Long = &H108&     'From mux.h

Private Enum PD_VP8StatusCode
    VP8_STATUS_OK = 0
    VP8_STATUS_OUT_OF_MEMORY
    VP8_STATUS_INVALID_PARAM
    VP8_STATUS_BITSTREAM_ERROR
    VP8_STATUS_UNSUPPORTED_FEATURE
    VP8_STATUS_SUSPENDED
    VP8_STATUS_USER_ABORT
    VP8_STATUS_NOT_ENOUGH_DATA
End Enum

#If False Then
    Private Const VP8_STATUS_OK = 0, VP8_STATUS_OUT_OF_MEMORY = 0, VP8_STATUS_INVALID_PARAM = 0, VP8_STATUS_BITSTREAM_ERROR = 0, VP8_STATUS_UNSUPPORTED_FEATURE = 0, VP8_STATUS_SUSPENDED = 0, VP8_STATUS_USER_ABORT = 0, VP8_STATUS_NOT_ENOUGH_DATA = 0
#End If

'Encoding error conditions.
Private Enum PD_WebPEncodingError
    VP8_ENC_OK = 0
    VP8_ENC_ERROR_OUT_OF_MEMORY             ' memory error allocating objects
    VP8_ENC_ERROR_BITSTREAM_OUT_OF_MEMORY   ' memory error while flushing bits
    VP8_ENC_ERROR_NULL_PARAMETER            ' a pointer parameter is NULL
    VP8_ENC_ERROR_INVALID_CONFIGURATION     ' configuration is invalid
    VP8_ENC_ERROR_BAD_DIMENSION             ' picture has invalid width/height
    VP8_ENC_ERROR_PARTITION0_OVERFLOW       ' partition is bigger than 512k
    VP8_ENC_ERROR_PARTITION_OVERFLOW        ' partition is bigger than 16M
    VP8_ENC_ERROR_BAD_WRITE                 ' error while flushing bytes
    VP8_ENC_ERROR_FILE_TOO_BIG              ' file is bigger than 4G
    VP8_ENC_ERROR_USER_ABORT                ' abort request by user
    VP8_ENC_ERROR_LAST                      ' list terminator. always last.
End Enum

#If False Then
    Private Const VP8_ENC_OK = 0, VP8_ENC_ERROR_OUT_OF_MEMORY = 0, VP8_ENC_ERROR_BITSTREAM_OUT_OF_MEMORY = 0, VP8_ENC_ERROR_NULL_PARAMETER = 0, VP8_ENC_ERROR_INVALID_CONFIGURATION = 0, VP8_ENC_ERROR_BAD_DIMENSION = 0, VP8_ENC_ERROR_PARTITION0_OVERFLOW = 0, VP8_ENC_ERROR_PARTITION_OVERFLOW = 0, VP8_ENC_ERROR_BAD_WRITE = 0, VP8_ENC_ERROR_FILE_TOO_BIG = 0, VP8_ENC_ERROR_USER_ABORT = 0, VP8_ENC_ERROR_LAST = 0
#End If

Private Enum PD_WebPFormatFeature
    WEBP_FF_FORMAT_FLAGS     'bit-wise combination of WebPFeatureFlags corresponding to the 'VP8X' chunk (if present).
    WEBP_FF_CANVAS_WIDTH
    WEBP_FF_CANVAS_HEIGHT
    WEBP_FF_LOOP_COUNT       'only relevant for animated file
    WEBP_FF_BACKGROUND_COLOR 'idem.
    WEBP_FF_FRAME_COUNT      'Number of frames present in the demux object. In case of a partial demux, this is the number
                             ' of frames seen so far, with the last frame possibly being partial.
End Enum

#If False Then
    Private Const WEBP_FF_FORMAT_FLAGS = 0, WEBP_FF_CANVAS_WIDTH = 0, WEBP_FF_CANVAS_HEIGHT = 0, WEBP_FF_LOOP_COUNT = 0, WEBP_FF_BACKGROUND_COLOR = 0, WEBP_FF_FRAME_COUNT = 0
#End If

'VP8X Feature Flags (this chunk is optional in basic single-frame files)
Private Enum PD_WebPFeatureFlags
    ANIMATION_FLAG = &H2&
    XMP_FLAG = &H4&
    EXIF_FLAG = &H8&
    ALPHA_FLAG = &H10&
    ICCP_FLAG = &H20&
    ALL_VALID_FLAGS = &H3E&
End Enum

#If False Then
    Private Const ANIMATION_FLAG = &H2&, XMP_FLAG = &H4&, EXIF_FLAG = &H8&, ALPHA_FLAG = &H10&, ICCP_FLAG = &H20&, ALL_VALID_FLAGS = &H3E&
#End If

Private Enum PD_WEBP_CSP_MODE
    MODE_RGB = 0
    MODE_RGBA = 1
    MODE_BGR = 2
    MODE_BGRA = 3
    MODE_ARGB = 4
    MODE_RGBA_4444 = 5
    MODE_RGB_565 = 6
    'RGB-premultiplied transparent modes (alpha value is preserved)
    ' (names changed to work in VB, which treats enums as case-insensitive)
    MODE_rgbAp = 7
    MODE_bgrAp = 8
    MODE_Argbp = 9
    MODE_rgbAp_4444 = 10
    'YUV modes; note that these are *not* recommended for lossless mode
    MODE_YUV = 11
    MODE_YUVA = 12  '// yuv 4:2:0
    MODE_LAST = 13
End Enum

#If False Then
    Private Const MODE_RGB = 0, MODE_RGBA = 1, MODE_BGR = 2, MODE_BGRA = 3, MODE_ARGB = 4, MODE_RGBA_4444 = 5, MODE_RGB_565 = 6
    Private Const MODE_rgbAp = 7, MODE_bgrAp = 8, MODE_Argbp = 9, MODE_rgbAp_4444 = 10
    Private Const MODE_YUV = 11, MODE_YUVA = 12, MODE_LAST = 13
#End If

'Various structs follow.  These all require use of the advanced decoding interface;
' the "quick" one doesn't work with these
Private Type PD_WebPBitstreamFeatures
    webpWidth As Long        'Width in pixels, as read from the bitstream.
    webpHeight As Long       'Height in pixels, as read from the bitstream.
    webHasAlpha As Long      'True if the bitstream contains an alpha channel.
    webpHasAnimation As Long 'True if the bitstream is an animation.
    webpFormat As Long       '0 = undefined (/mixed), 1 = lossy, 2 = lossless
    webpPadding(5) As Long   'padding for later use
End Type

'Animation options
Private Type PD_WebPAnimDecoderOptions
    'Output colorspace. Only the following modes are supported:
    '  MODE_RGBA, MODE_BGRA, MODE_rgbA and MODE_bgrA.
    color_mode As PD_WEBP_CSP_MODE
    use_threads As Long         'If true, use multi-threaded decoding.
    padding(7) As Long          '// Padding for later use.
End Type

'Global information about the animation
Private Type PD_WebPAnimInfo
    canvas_width As Long
    canvas_height As Long
    loop_count As Long
    bgcolor As Long
    frame_count As Long
    padding(0 To 3) As Long  ' // padding for later use
End Type

'Image characteristics hint for the underlying encoder.
Private Enum PD_WebPImageHint
    WEBP_HINT_DEFAULT = 0    ' default preset.
    WEBP_HINT_PICTURE        ' digital picture, like portrait, inner shot
    WEBP_HINT_PHOTO          ' outdoor photograph, with natural lighting
    WEBP_HINT_GRAPH          ' discrete tone image (graph, map-tile etc).
    WEBP_HINT_LAST
End Enum

'Enumerate some predefined settings for WebPConfig, depending on the type
' of source picture. These presets are used when calling WebPConfigPreset().
Private Enum PD_WebPPreset
    WEBP_PRESET_DEFAULT = 0  ' default preset.
    WEBP_PRESET_PICTURE      ' digital picture, like portrait, inner shot
    WEBP_PRESET_PHOTO        ' outdoor photograph, with natural lighting
    WEBP_PRESET_DRAWING      ' hand or line drawing, with high-contrast details
    WEBP_PRESET_ICON         ' small-sized colorful images
    WEBP_PRESET_TEXT         ' text-like
End Enum

'Compression parameters
Private Type PD_WebPConfig
    lossless As Long          'Lossless encoding (0=lossy(default), 1=lossless).
    quality As Single         'Between 0 and 100. For lossy, 0 gives the smallest
                              ' size and 100 the largest. For lossless, this
                              ' parameter is the amount of effort put into the
                              ' compression: 0 is the fastest but gives larger
                              ' files compared to the slowest, but best, 100.
    compressMethod As Long    'Quality/speed trade-off (0=fast, 6=slower-better)
    image_hint As PD_WebPImageHint 'Hint for image type (lossless only for now).
    target_size As Long       'If non-zero, set the desired target size in bytes.
                              ' Takes precedence over the 'compression' parameter.
    target_PSNR As Single     'If non-zero, specifies the minimal distortion to
                              ' try to achieve. Takes precedence over target_size.
    segments As Long          'Maximum number of segments to use, in [1..4]
    sns_strength As Long      'Spatial Noise Shaping. 0=off, 100=maximum.
    filter_strength As Long   'Range: [0 = off .. 100 = strongest]
    filter_sharpness As Long  'Range: [0 = off .. 7 = least sharp]
    filter_type As Long       'Filtering type: 0 = simple, 1 = strong (only used
                              ' if filter_strength > 0 or autofilter > 0)
    autofilter As Long        'Auto adjust filter's strength [0 = off, 1 = on]
    alpha_compression As Long 'Algorithm for encoding the alpha plane (0 = none,
                              ' 1 = compressed with WebP lossless). Default is 1.
    alpha_filtering As Long   'Predictive filtering method for alpha plane.
                              '  0: none, 1: fast, 2: best. Default if 1.
    alpha_quality As Long     'Between 0 (smallest size) and 100 (lossless).
                              ' Default is 100.
    pass As Long              'Number of entropy-analysis passes (in [1..10]).
    show_compressed As Long   'If true, export the compressed picture back.
                              ' In-loop filtering is not applied.
    preprocessing As Long     'Preprocessing filter:
                              ' 0=none, 1=segment-smooth, 2=pseudo-random dithering
    partitions As Long        'Log2(number of token partitions) in [0..3]. Default
                              ' is set to 0 for easier progressive decoding.
    partition_limit As Long   'Quality degradation allowed to fit the 512k limit
                              ' on prediction modes coding (0: no degradation,
                              ' 100: maximum possible degradation).
    emulate_jpeg_size As Long 'If true, compression parameters will be remapped
                              ' to better match the expected output size from
                              ' JPEG compression. Generally, the output size will
                              ' be similar but the degradation will be lower.
    thread_level As Long      'If non-zero, try and use multi-threaded encoding.
    low_memory As Long        'If set, reduce memory usage (but increase CPU use).
    near_lossless As Long     'Near lossless encoding [0 = max loss .. 100 = off
                              ' (default)].
    exact As Long             'If non-zero, preserve the exact RGB values under
                              ' transparent area. Otherwise, discard this invisible
                              ' RGB information for better compression. The default
                              ' value is 0.
    use_delta_palette As Long 'Reserved for future lossless feature
    use_sharp_yuv As Long     'If needed, use sharp (and slow) RGB->YUV conversion
    qmin As Long              'Minimum permissible quality factor
    qmax As Long              'Maximum permissible quality factor
End Type

'Encoding color spaces
Private Enum PD_WebPEncCSP
    WEBP_YUV420 = 0         '4:2:0
    WEBP_YUV420A = 4        'alpha channel variant
    WEBP_CSP_UV_MASK = 3    'bit-mask to get the UV sampling factors
    WEBP_CSP_ALPHA_BIT = 4  'bit that is set if alpha is present
End Enum

'Main exchange structure (input samples, output bytes, statistics)
'  Once WebPPictureInit() has been called, it's ok to make all the INPUT fields
'  (use_argb, y/u/v, argb, ...) point to user-owned data, even if
'  WebPPictureAlloc() has been called. Depending on the value use_argb,
'  it's guaranteed that either *argb or *y/*u/*v content will be kept untouched.
Private Type PD_WebPPicture
  
    'INPUT
  
    'Main flag for encoder selecting between ARGB or YUV input.
    '  It is recommended to use ARGB input (*argb, argb_stride) for lossless
    '  compression, and YUV input (*y, *u, *v, etc.) for lossy compression
    '  since these are the respective native colorspace for these formats.
    use_argb As Long
    
    'YUV input (mostly used for input to lossy compression)
    colorspace As PD_WebPEncCSP '// colorspace: should be YUV420 for now (=Y'CbCr).
    webp_width As Long
    webp_height As Long         '// dimensions (less or equal to WEBP_MAX_DIMENSION)
    ptrY As Long
    ptrU As Long
    ptrV As Long                '// pointers to luma/chroma planes.
    y_stride As Long
    uv_stride As Long           '// luma/chroma strides.
    ptrA As Long                '// pointer to the alpha plane
    a_stride As Long            '// stride of the alpha plane
    padding1a As Long           '// padding for later use
    padding1b As Long
    
    '// ARGB input (mostly used for input to lossless compression)
    ptrARGB As Long             '// Pointer to argb (32 bit) plane.
    argb_stride As Long         '// This is stride in pixels units, not bytes.
    padding2a As Long           '// padding for later use
    padding2b As Long
    padding2c As Long

    '//   OUTPUT
    
    '// Byte-emission hook, to store compressed bytes as they are ready.
    ptrToWebPWriterFunction As Long '// can be NULL
    custom_ptr As Long              '// can be used by the writer.

    '// map for extra information (only for lossy compression mode)
    extra_info_type As Long     '// 1: intra type, 2: segment, 3: quant
                                '// 4: intra-16 prediction mode,
                                '// 5: chroma prediction mode,
                                '// 6: bit cost, 7: distortion
    extra_info As Long          '// if not NULL, points to an array of size
                                '// ((width + 15) / 16) * ((height + 15) / 16) that
                                '// will be filled with a macroblock map, depending
                                '// on extra_info_type.

    '//   STATS AND REPORTS
    
    '// Pointer to side statistics (updated only if not NULL)
    stats_ptr As Long

    '// Error code for the latest error encountered during encoding
    error_code As PD_WebPEncodingError
    
    '// If not NULL, report progress during encoding.
    progress_hook As Long
    
    user_data As Long       '// this field is free to be set to any value and
                            '// used during callbacks (like progress-report e.g.).

    padding3a As Long       '// padding for later use
    padding3b As Long
    padding3c As Long
    
    '// Unused for now
    padding4 As Long
    padding5 As Long
    padding6a As Long       '// padding for later use
    padding6b As Long
    padding6c As Long
    padding6d As Long
    padding6e As Long
    padding6f As Long
    padding6g As Long
    padding6h As Long
    
    '// PRIVATE FIELDS
    
    memory_yuva As Long     '// row chunk of memory for yuva planes
    memory_argb As Long     '// and for argb too.
    padding7a As Long       '// padding for later use
    padding7b As Long
    
End Type

'WebPMemoryWrite: a special WebPWriterFunction that writes to memory using
' the following WebPMemoryWriter object (to be set as a custom_ptr).
Private Type PD_WebPMemoryWriter
    mem_buffer As Long      '// final buffer (of size 'max_size', larger than 'size').
    mem_size As Long        '// final size
    max_size As Long        '// total capacity
    padding As Long         '// padding for later use
End Type

'Animation parameters.
Private Type PD_WebPMuxAnimParams
    bgColorA As Byte    '// Background color of the canvas stored (in MSB order) as ARGB
    bgColorR As Byte
    bgColorG As Byte
    bgColorB As Byte
    loop_count As Long  '// Number of times to repeat the animation [0 = infinite].
End Type

Private Type PD_WebPAnimEncoderOptions
    anim_params As PD_WebPMuxAnimParams  '// Animation parameters.
    minimize_size As Long     '// If true, minimize the output size (slow). Implicitly
                              '// disables key-frame insertion.
    kmin As Long
    kmax As Long              '// Minimum and maximum distance between consecutive key
                              '// frames in the output. The library may insert some key
                              '// frames as needed to satisfy this criteria.
                              '// Note that these conditions should hold: kmax > kmin
                              '// and kmin >= kmax / 2 + 1. Also, if kmax <= 0, then
                              '// key-frame insertion is disabled; and if kmax == 1,
                              '// then all frames will be key-frames (kmin value does
                              '// not matter for these special cases).
    allow_mixed As Long       '// If true, use mixed compression mode; may choose
                              '// either lossy and lossless for each frame.
    verbose As Long           '// If true, print info and warning messages to stderr.
    padding1a As Long         '// Padding for later use.
    padding1b As Long
    padding1c As Long
    padding1d As Long
End Type

'Note that regular WebP features reside in libwebp; animation features are in libwebpde/mux.

'libwebp uses a Google-specific toolchain for compilation, so we lean on their supplied makefile for building.
' This produces cdecl DLLs which we must wrap using DispCallFunc.
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum PD_WebP_ProcAddress
    WebPAnimDecoderDelete
    WebPAnimDecoderGetInfo
    WebPAnimDecoderGetNext
    WebPAnimDecoderHasMoreFrames
    WebPAnimDecoderNew
    WebPAnimDecoderOptionsInit
    WebPAnimDecoderReset
    WebPAnimEncoderAdd
    WebPAnimEncoderAssemble
    WebPAnimEncoderDelete
    WebPAnimEncoderGetError
    WebPAnimEncoderNew
    WebPAnimEncoderOptionsInit
    WebPBlendAlpha
    WebPCleanupTransparentArea
    WebPConfigInit
    WebPDecodeBGRAInto
    WebPDemux
    WebPDemuxDelete
    WebPDemuxGetI
    WebPEncode
    WebPFree
    WebPGetDecoderVersion
    WebPGetFeatures
    WebPGetInfo
    WebPMemoryWrite
    WebPMemoryWriterClear
    WebPMemoryWriterInit
    WebPPictureAlloc
    WebPPictureARGBToYUVA
    WebPPictureARGBToYUVADithered
    WebPPictureCopy
    WebPPictureCrop
    WebPPictureDistortion
    WebPPictureFree
    WebPPictureHasTransparency
    WebPPictureImportBGR
    WebPPictureImportBGRA
    WebPPictureImportBGRX
    WebPPictureImportRGB
    WebPPictureImportRGBA
    WebPPictureImportRGBX
    WebPPictureInitInternal
    WebPPictureIsView
    WebPPictureRescale
    WebPPictureSharpARGBToYUVA
    WebPPictureSmartARGBToYUVA
    WebPPictureView
    WebPPictureYUVAToARGB
    WebPPlaneDistortion
    WebPValidateConfig
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Features of the current file follow
Private m_fileFeatures As PD_WebPBitstreamFeatures
Private m_imgWidth As Long, m_imgHeight As Long

'Many WebP features require a specialized demuxer object to retrieve.  If this handle is non-zero,
' a demuxer has been allocated.  You *MUST* free this with a call to WebPDemuxDelete().
Private m_hDemuxer As Long

'Animated files will fill this struct.  (Non-animated files WILL NOT.)
Private m_AnimationInfo As PD_WebPAnimInfo

'PhotoDemon provides a "streaming" interface for writing animated WebP files, where the total frame count
' (and contents of individual frames) do not need to be known up-front.  Instead, you "stream" frames to
' the encoder as they arrive.  This requires class-level tracking variables to correctly maintain state.
Private m_StreamingEncoder As Long, m_StreamingSaveConfig As PD_WebPConfig

Friend Function GetLibraryVersion() As Long
    GetLibraryVersion = CallCDeclW(WebPGetDecoderVersion, vbLong)
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = (m_fileFeatures.webHasAlpha <> 0)
End Function

Friend Function IsAnimated() As Boolean
    IsAnimated = (m_fileFeatures.webpHasAnimation <> 0)
End Function

'Load a WebP file from memory.  This is used exclusively in PD because libwebp does *not* provide any
' direct file-read functions.  (All load functions operate on pointers.)
'
'Anyway, we leave it to the caller to figure out how they want to allocate+store the data.  We just need
' the base pointer and the size, which are blindly relayed to libwebp.
'
'The filename is only passed so that we can name the base layer (in a non-animated file)
Friend Function LoadWebP_FromMemory(ByRef srcFile As String, ByVal pData As Long, ByVal sizeOfDataInBytes As Long, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    Const FUNC_NAME = "LoadWebP_FromMemory"
    
    LoadWebP_FromMemory = False
    Dim webpResult As PD_VP8StatusCode
    
    'Start by querying basic features, like "is the file animated"
    webpResult = CallCDeclW(WebPGetFeatures, vbLong, pData, sizeOfDataInBytes, VarPtr(m_fileFeatures), WEBP_DECODER_ABI_VERSION)
    If (webpResult <> VP8_STATUS_OK) Then
        InternalError FUNC_NAME, "WebPGetFeaturesInternal", webpResult
        Exit Function
    End If
    
    If WEBP_DEBUG_VERBOSE Then PDDebug.LogAction "WebP file features: animated: " & (m_fileFeatures.webpHasAnimation <> 0) & ", alpha: " & (m_fileFeatures.webHasAlpha <> 0) & ", lossless: " & (m_fileFeatures.webpFormat = 2)
    m_imgWidth = m_fileFeatures.webpWidth
    m_imgHeight = m_fileFeatures.webpHeight
    If WEBP_DEBUG_VERBOSE Then PDDebug.LogAction "WebP image size: " & m_imgWidth & "x" & m_imgHeight
    
    'Perform basic validation
    If (m_imgWidth < 1) Or (m_imgHeight < 1) Then
        InternalError FUNC_NAME, "bad image width/height"
        Exit Function
    End If
    
    'Unfortunately for us, libwebp has a complex and poorly conceived API design.  You'd think that our
    ' previous call to WebPGetFeatures would tell us all the critical features of the file, but no,
    ' it doesn't.  Instead, we now need to create a demuxer object, which reports its own feature
    ' collection (that overlaps the other feature set in some ways but not all ways, sigh).
    '
    'Note that failure to create a demuxer isn't necessarily a deal-breaker - but it does mean
    ' that we can't query important features like ICC profiles.
    If (Not InitializeDemuxer(pData, sizeOfDataInBytes)) Then InternalError FUNC_NAME, "couldn't allocate demuxer; attempting to load anyway"
    
    'If a demuxer *was* created successfully, use feature flags to look for an ICC profile
    Dim imgHasColorProfile As Boolean
    If (m_hDemuxer <> 0) Then
        
        Dim curImageFeatureFlags As PD_WebPFeatureFlags
        curImageFeatureFlags = CallCDeclW(WebPDemuxGetI, vbLong, m_hDemuxer, WEBP_FF_FORMAT_FLAGS)
        imgHasColorProfile = (curImageFeatureFlags And ICCP_FLAG)
        If WEBP_DEBUG_VERBOSE Then PDDebug.LogAction "WebP features; ICC: " & CBool(curImageFeatureFlags And ICCP_FLAG) & ", EXIF: " & CBool(curImageFeatureFlags And EXIF_FLAG) & ", XMP: " & CBool(curImageFeatureFlags And XMP_FLAG)
        
        'TODO: find WebP images "in the wild" with embedded ICC profiles so I can test retrieval
        
    End If
    
    'If we're still here, the file looks readable.  How we proceed depends on whether the image is
    ' animated or not.  (Animated images produce multiple layers.)
    
    'Attempt animated image import first, so that if something goes wrong we can fall back to single-frame retrieval
    Dim animationLoaded As Boolean: animationLoaded = False
    If (m_fileFeatures.webpHasAnimation <> 0) Then animationLoaded = LoadAnimatedWebP(srcFile, pData, sizeOfDataInBytes, dstImage, dstDIB)
    
    If animationLoaded Then
    
        'Our work here is done!  Note that we don't exit immediately, as the end of the function may
        ' need to clean-up some libwebp objects before exiting
        LoadWebP_FromMemory = True
    
    'If the underlying image is *not* animated (or if animation import failed), load the file as
    ' a single-frame instance.
    Else
    
        'This file is not animated.  Load it normally.
        
        'Create the destination surface.  (We can decode the WebP directly into this buffer, yay!)
        Set dstDIB = New pdDIB
        dstDIB.CreateBlank m_imgWidth, m_imgHeight, 32, 255, 255
        
        'Decode the image directly into the buffer
        Dim decodeOK As Boolean
        With dstDIB
            decodeOK = (CallCDeclW(WebPDecodeBGRAInto, vbLong, pData, sizeOfDataInBytes, .GetDIBPointer, .GetDIBStride * .GetDIBHeight, .GetDIBStride) <> 0)
        End With
        
        'Ensure success (this particular function returns 0 when unsuccessful, per the docs)
        If decodeOK Then
            
            'TODO: get libwebp to paint in premultiplied mode
            dstDIB.SetAlphaPremultiplication True
            
            'Initialize the target layer
            Dim newLayerID As Long, tmpLayer As pdLayer
            If (Not dstImage Is Nothing) Then
                
                newLayerID = dstImage.CreateBlankLayer()
                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                tmpLayer.InitializeNewLayer PDL_Image, Files.FileGetName(srcFile, True), dstDIB, False
                tmpLayer.SetLayerVisibility True
            
                'Finish initializing the target image
                dstImage.Width = m_imgWidth
                dstImage.Height = m_imgHeight
            
                'WebP files do not provide a direct way to set DPI
                dstImage.SetDPI 72, 72
                dstImage.NotifyImageChanged UNDO_Everything
            
            End If
            
            LoadWebP_FromMemory = True
            
        'decoding pixel data failed - not much we can do but abort completely
        Else
            InternalError FUNC_NAME, "bad WebPDecodeBGRAInto"
            Set dstDIB = Nothing
            LoadWebP_FromMemory = False
        End If
    
    '/end single-frame load process
    End If
    
    'Before exiting, free any allocated demuxer(s)
    FreeDemuxer
    
End Function

'Thin wrapper to LoadWebP_FromMemory
Friend Function QuickLoadWebP_FromFile(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    QuickLoadWebP_FromFile = False
    
    If Plugin_WebP.IsWebP(srcFile) Then
        Dim fullFile() As Byte
        If Files.FileLoadAsByteArray(srcFile, fullFile) Then QuickLoadWebP_FromFile = LoadWebP_FromMemory(srcFile, VarPtr(fullFile(0)), UBound(fullFile) + 1, dstImage, dstDIB)
    End If
    
End Function

'Save a pdImage (with animation data) to an animated WebP file
Friend Function SaveAnimatedWebP_ToFile(ByRef srcImage As pdImage, ByRef srcOptions As String, ByRef dstFile As String) As Boolean
    
    Const FUNC_NAME As String = "SaveAnimatedWebP_ToFile"
    SaveAnimatedWebP_ToFile = False
    
    'Basic input validation
    If (srcImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError FUNC_NAME, "invalid function parameters"
        Exit Function
    End If
    
    'Before doing anything else, make sure the source image has multiple layers.  (If it doesn't,
    ' we'll silently forward it to the single-frame saver.)
    If (srcImage.GetNumOfLayers < 2) Then
        PDDebug.LogAction "single frame file found - rerouting to static exporter"
        SaveAnimatedWebP_ToFile = Me.SaveWebP_ToFile(srcImage, srcOptions, dstFile)
        Exit Function
    End If
    
    'Failsafe check only; PD enforces safe overwriting in a parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'Initialize a progress bar
    ProgressBars.SetProgBarMax srcImage.GetNumOfLayers()
    
    'Parse incoming options string
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    cSettings.SetParamString srcOptions
    
    'Retrieve frame delay defaults and overrides from the passed param string
    Dim useFixedFrameDelay As Boolean, frameDelayDefault As Long
    useFixedFrameDelay = cSettings.GetBool("use-fixed-frame-delay", False)
    frameDelayDefault = cSettings.GetLong("frame-delay-default", 100)
    
    'For WebP files, animation encoding is primarily routed through a WebpAnimEncoder object.
    ' This object relies on a set of global animation options; prepare an option struct now.
    Dim enc_options As PD_WebPAnimEncoderOptions
    If (CallCDeclW(WebPAnimEncoderOptionsInit, vbLong, VarPtr(enc_options), WEBP_MUX_ABI_VERSION) = 0) Then
        InternalError FUNC_NAME, "bad options init"
        
        'Docs say: "WebPAnimEncoderOptionsInit() must have succeeded before using the 'enc_options' object.
        Exit Function
    End If
    
    'Populate any settings we have control over (like loop count)
    With enc_options
        .anim_params.bgColorA = 0
        .anim_params.bgColorR = 0
        .anim_params.bgColorG = 0
        .anim_params.bgColorB = 0
        .anim_params.loop_count = cSettings.GetLong("animation-loop-count", 1) '(0 = infinite loops)
        
        'WebP offers some straightforward minimization options.  Note that the value of these functions
        ' is tied to the incoming quality value, because the function will use heuristics to compare
        ' the size of lossless vs lossy encoding of each frame *at the requested quality*.  Lossless will
        ' be used if it produces a smaller file (which is not especially rare if coming from a limited-color
        ' source like a GIF.  Note also that specifying SLOW compression will heavily penalize this function,
        ' because each frame will be tested against a max-compression-lossless encoding... but you can
        ' produce very small files this way (likely beating out both animated GIF and PNG).
        .minimize_size = 1
        .allow_mixed = 1
        
    End With
    
    'Next, we need an encoder object.  (Failure to initialize is a deal-breaker.)
    Dim enc As Long
    enc = CallCDeclW(WebPAnimEncoderNew, vbLong, srcImage.Width, srcImage.Height, VarPtr(enc_options), WEBP_MUX_ABI_VERSION)
    If (enc = 0) Then
        InternalError FUNC_NAME, "bad encoder creation"
        Exit Function
    End If
    
    'Next, we add all frames to the encoder.  This process is fairly complicated as we need to first convert
    ' each individual frame into a WebPPicture (using an associated WebPConfig compression struct).
    
    'Start by initializing a configuration struct using the settings string we were passed.  This struct
    ' will be reused for *all* frames (because they're all encoded with identical settings).
    Dim saveConfig As PD_WebPConfig
    If (Not GetWebPConfig(cSettings, saveConfig)) Then InternalError FUNC_NAME, "config object may be bad!"
    
    '(If you want to manually modify any configuration settings, this is your last chance.)
    
    Dim tmpLayer As pdLayer
    Dim runningTimestamp As Long, curFrameTime As Long
    runningTimestamp = 0
    
    'Note that unlike GIF and PNG, we do not auto-optimize each frame because libwebp will do this for us.
    ' (I have not attempted to compare WebP's built-in optimizations against PD's strategies, but the
    ' animations produced natively by the tool tend to beat GIF and PNG if you allow a slight quality loss,
    ' so whatever the library does to optimize each frame seems to work well enough.)
    Dim idxFrame As Long
    For idxFrame = 0 To srcImage.GetNumOfLayers - 1
        
        'This stage can be time-consuming, so we provide UI feedback
        ProgressBars.SetProgBarVal idxFrame
        Message "Saving animation frame %1 of %2...", idxFrame + 1, srcImage.GetNumOfLayers, "DONOTLOG"
        
        'Convert the source layer to an image-sized, zero-offsets, all-transforms-applied layer,
        ' then cache a local copy of it.  (This MUST be lossless for the source layer!)
        Set tmpLayer = New pdLayer
        tmpLayer.CopyExistingLayer srcImage.GetLayerByIndex(idxFrame)
        tmpLayer.ConvertToNullPaddedLayer srcImage.Width, srcImage.Height, True
        
        'Convert this frame into a WebP picture object
        Dim wpPicture As PD_WebPPicture
        If GetWebPPicture(tmpLayer.GetLayerDIB, wpPicture, True) Then
            
            'Calculate a timestamp for this frame.  WebP doesn't use frame times; instead, it uses
            ' a (0-based, presumably?) additive timestamp and auto-calculates frame times from that.
            If (idxFrame > 0) Then
                
                'Before dealing with pixel data, attempt to retrieve a frame time from the source layer's name.
                ' (If the layer name does not provide a frame time, or if the user has specified a fixed
                ' frame time, this value will be overwritten with their requsted value.)
                curFrameTime = GetFrameTimeFromLayerName(srcImage.GetLayerByIndex(idxFrame - 1).GetLayerName, 0)
                If (useFixedFrameDelay Or (curFrameTime = 0)) Then curFrameTime = frameDelayDefault
                runningTimestamp = runningTimestamp + curFrameTime
                
            Else
                runningTimestamp = 0
            End If
            
            'With a WebPPicture created, we can now dump the frame into the encoder object
            If (CallCDeclW(WebPAnimEncoderAdd, vbLong, enc, VarPtr(wpPicture), runningTimestamp, VarPtr(saveConfig)) = 0) Then
                InternalError FUNC_NAME, "bad encoder add: " & wpPicture.error_code
            End If
            
        Else
            InternalError FUNC_NAME, "bad layer picture # " & idxFrame
        End If
        
        'Free the underlying picture object or it will leak!
        CallCDeclW WebPPictureFree, vbEmpty, VarPtr(wpPicture)
        
    Next idxFrame
    
    'With all frames added, we need to call WebPAnimEncoderAdd one last time with NULL inputs except
    ' for the timestamp.  This (asinine) technique is used to establish frame time for the last frame.
    curFrameTime = GetFrameTimeFromLayerName(srcImage.GetLayerByIndex(srcImage.GetNumOfLayers - 1).GetLayerName, 0)
    If (useFixedFrameDelay Or (curFrameTime = 0)) Then curFrameTime = frameDelayDefault
    runningTimestamp = runningTimestamp + curFrameTime
    If (CallCDeclW(WebPAnimEncoderAdd, vbLong, enc, 0&, runningTimestamp, 0&) = 0) Then InternalError FUNC_NAME, "bad final frame"
    
    'With all frames assembled, we now need to generate the final bitstream.  The results of this operation
    ' get placed into a WebPData struct.
    ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax()
    Message "Finalizing image..."
    
    Dim unionPtrAndLen(0 To 1) As Long
    SaveAnimatedWebP_ToFile = (CallCDeclW(WebPAnimEncoderAssemble, vbLong, enc, VarPtr(unionPtrAndLen(0))) <> 0)
    
    'Regardless of outcome, the encoder is finished; free it immediately
    CallCDeclW WebPAnimEncoderDelete, vbEmpty, enc
    
    If (Not SaveAnimatedWebP_ToFile) Then
        InternalError FUNC_NAME, "bad assembly"
    Else
        
        'Dump the pointer to file!
        Dim cStream As pdStream
        Set cStream = New pdStream
        SaveAnimatedWebP_ToFile = cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadWrite, dstFile, optimizeAccess:=OptimizeSequentialAccess)
        cStream.WriteBytesFromPointer unionPtrAndLen(0), unionPtrAndLen(1)
        cStream.StopStream
        Set cStream = Nothing
        
        'Free the finished buffer
        CallCDeclW WebPFree, vbEmpty, VarPtr(unionPtrAndLen(0))
        
    End If
    
    'Free any remaining UI elements
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
End Function

'This next set of "streaming" functions are used to record an animated WebP to file in stages,
' by "streaming" frames to the class.  You do not need to know the number of frames in advance,
' but you do need to establish the canvas size up-front.
'
'Note that you *must* call the streaming functions in the correct order (Start, Add, Stop).
' Calling the functions out of order could crash (or worse).
Friend Function SaveStreamingWebP_AddFrame(ByRef srcDIB As pdDIB, ByVal frameTimestamp As Long) As Boolean
    
    Const FUNC_NAME As String = "SaveStreamingWebP_AddFrame"
    SaveStreamingWebP_AddFrame = False
    
    'Start by converting the frame into a WebP picture object
    Dim wpPicture As PD_WebPPicture
    If GetWebPPicture(srcDIB, wpPicture, True) Then
        
        'With a WebPPicture created, we can now dump the frame into the encoder object
        If (CallCDeclW(WebPAnimEncoderAdd, vbLong, m_StreamingEncoder, VarPtr(wpPicture), frameTimestamp, VarPtr(m_StreamingSaveConfig)) = 0) Then
            InternalError FUNC_NAME, "bad encoder add: " & wpPicture.error_code
        Else
            SaveStreamingWebP_AddFrame = True
        End If
        
    Else
        InternalError FUNC_NAME, "bad frame"
    End If
    
    'Free the underlying picture object or it will leak!
    CallCDeclW WebPPictureFree, vbEmpty, VarPtr(wpPicture)
    
End Function

Friend Function SaveStreamingWebP_Cancel() As Boolean
    
    'Free any running WebP objects
    CallCDeclW WebPAnimEncoderDelete, vbEmpty, m_StreamingEncoder
    m_StreamingEncoder = 0
    
End Function

Friend Function SaveStreamingWebP_Start(ByVal pxWidth As Long, ByVal pxHeight As Long, Optional ByVal srcLoopCount As Long = 0, Optional ByVal encodingQuality As Single = 75!) As Boolean
    
    Const FUNC_NAME As String = "SaveStreamingWebP_Start"
    SaveStreamingWebP_Start = False
    
    'Basic input validation
    If (pxWidth <= 0) Or (pxHeight <= 0) Then
        InternalError FUNC_NAME, "invalid parameters"
        Exit Function
    End If
    
    'Note that we don't need to touch the destination file (or even ask for it) yet; libwebp caches all
    ' animation work in-memory, and simply supplies a pointer to the assembled image at the end.
    
    'For WebP files, animation encoding is primarily routed through a WebpAnimEncoder object.
    ' This object relies on a set of global animation options; prepare an option struct now.
    Dim enc_options As PD_WebPAnimEncoderOptions
    If (CallCDeclW(WebPAnimEncoderOptionsInit, vbLong, VarPtr(enc_options), WEBP_MUX_ABI_VERSION) = 0) Then
        InternalError FUNC_NAME, "bad options init"
        
        'Docs say: "WebPAnimEncoderOptionsInit() must have succeeded before using the 'enc_options' object.
        Exit Function
    End If
    
    'Populate any settings we have control over (like loop count)
    With enc_options
        .anim_params.bgColorA = 255
        .anim_params.bgColorR = 255
        .anim_params.bgColorG = 255
        .anim_params.bgColorB = 255
        .anim_params.loop_count = srcLoopCount
        
        'WebP offers some straightforward minimization options.  Note that the value of these functions
        ' is tied to the incoming quality value, because the function will use heuristics to compare
        ' the size of lossless vs lossy encoding of each frame *at the requested quality*.  Lossless will
        ' be used if it produces a smaller file (which is not especially rare if coming from a limited-color
        ' source like a GIF.  Note also that specifying SLOW compression will heavily penalize this function,
        ' because each frame will be tested against a max-compression-lossless encoding... but you can
        ' produce very small files this way (likely beating out both animated GIF and PNG).
        .minimize_size = 1
        .allow_mixed = 1
        
    End With
    
    'Next, we need an encoder object.  (Failure to initialize is a deal-breaker.)
    ' This encoder must be cached, as it will be referenced on all frame "pushes".
    m_StreamingEncoder = CallCDeclW(WebPAnimEncoderNew, vbLong, pxWidth, pxHeight, VarPtr(enc_options), WEBP_MUX_ABI_VERSION)
    If (m_StreamingEncoder = 0) Then
        InternalError FUNC_NAME, "bad encoder creation"
        Exit Function
    End If
    
    'Because we are using identical encoding options for each frame, it's easier to construct (and store)
    ' a single config object up-front.  Because PD's config constructor requires a PD-format param string,
    ' we need to serialize any settings to such an object before handing it off to the constructor.
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    cSettings.AddParam "webp-quality", encodingQuality
    
    'Screen recordings are likely to be discrete-tone images, so supply some hints to improve encoding
    cSettings.AddParam "webp-hint", "chart"
    
    'Build the config object
    If (Not GetWebPConfig(cSettings, m_StreamingSaveConfig)) Then InternalError FUNC_NAME, "config object may be bad!"
    
    'Initialization is complete.
    SaveStreamingWebP_Start = True
    
End Function

Friend Function SaveStreamingWebP_Stop(ByVal finalTimestamp As Long, ByRef dstFile As String) As Boolean
    
    Const FUNC_NAME As String = "SaveStreamingWebP_Stop"
    SaveStreamingWebP_Stop = False
    
    'With all frames added, we need to call WebPAnimEncoderAdd one last time with NULL inputs except
    ' for the timestamp.  This (asinine) technique is used to establish frame time for the last frame.
    If (CallCDeclW(WebPAnimEncoderAdd, vbLong, m_StreamingEncoder, 0&, finalTimestamp, 0&) = 0) Then InternalError FUNC_NAME, "bad final frame"
    
    'With all frames assembled, we now need to generate the final bitstream.  The results of this operation
    ' get placed into a WebPData struct.
    
    Dim unionPtrAndLen(0 To 1) As Long
    SaveStreamingWebP_Stop = (CallCDeclW(WebPAnimEncoderAssemble, vbLong, m_StreamingEncoder, VarPtr(unionPtrAndLen(0))) <> 0)
    
    'Regardless of outcome, the encoder is finished; free it immediately
    CallCDeclW WebPAnimEncoderDelete, vbEmpty, m_StreamingEncoder
    
    If (Not SaveStreamingWebP_Stop) Then
        InternalError FUNC_NAME, "bad assembly"
    Else
        
        'Dump the pointer to file!
        Files.FileDeleteIfExists dstFile
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        SaveStreamingWebP_Stop = cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadWrite, dstFile, optimizeAccess:=OptimizeSequentialAccess)
        cStream.WriteBytesFromPointer unionPtrAndLen(0), unionPtrAndLen(1)
        cStream.StopStream
        Set cStream = Nothing
        
        'Free the finished buffer
        CallCDeclW WebPFree, vbEmpty, VarPtr(unionPtrAndLen(0))
        
    End If
    
End Function

'Preview the effect of WebP compression settings (in srcOptions) on an arbitrary pdDIB
Friend Function SaveWebP_PreviewOnly(ByRef srcDIB As pdDIB, ByRef srcOptions As String, ByRef dstDIB As pdDIB) As Boolean
    
    SaveWebP_PreviewOnly = False
    
    'The vast majority of this function is handled by a separate in-memory encoder.
    Dim memWriter As PD_WebPMemoryWriter
    SaveWebP_PreviewOnly = SaveWebP_ToMemory(srcDIB, srcOptions, memWriter)
    
    'If saving was successful, we simply need to decode the in-memory WebP back to a pdDIB object.
    If SaveWebP_PreviewOnly Then
        SaveWebP_PreviewOnly = Me.LoadWebP_FromMemory(vbNullString, memWriter.mem_buffer, memWriter.mem_size, Nothing, dstDIB)
    End If
    
    'Regardless of outcome, free the WebP memory struct
    If (memWriter.mem_buffer <> 0) Then CallCDeclW WebPMemoryWriterClear, vbEmpty, VarPtr(memWriter)
    
End Function

'Save an arbitrary pdImage to a single-frame WebP file
Friend Function SaveWebP_ToFile(ByRef srcImage As pdImage, ByRef srcOptions As String, ByRef dstFile As String) As Boolean
    
    SaveWebP_ToFile = False
    
    'Retrieve a composite copy of the image (layers are not supported by WebP)
    Dim tmpDIB As pdDIB
    srcImage.GetCompositedImage tmpDIB, False
    
    'The vast majority of this function is handled by a separate in-memory encoder.
    Dim memWriter As PD_WebPMemoryWriter
    SaveWebP_ToFile = SaveWebP_ToMemory(tmpDIB, srcOptions, memWriter)
    Set tmpDIB = Nothing
    
    'If saving was successful, we simply need to dump the encoded bits out to file
    If SaveWebP_ToFile Then
        
        'Dump the pointer to file!
        Dim cStream As pdStream
        Set cStream = New pdStream
        SaveWebP_ToFile = cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadWrite, dstFile, optimizeAccess:=OptimizeSequentialAccess)
        cStream.WriteBytesFromPointer memWriter.mem_buffer, memWriter.mem_size
        cStream.StopStream
        
    End If
    
    'Regardless of outcome, free the WebP memory struct
    If (memWriter.mem_buffer <> 0) Then CallCDeclW WebPMemoryWriterClear, vbEmpty, VarPtr(memWriter)
    
End Function

'Save an arbitrary DIB to a WebP memory struct.  Returns a WebP memory struct that the *CALLER* must dispose of.
' (This function obviously CANNOT dispose of the object without making this whole function pointless.)
' Disposal uses code like: CallCDeclW WebPMemoryWriterClear, vbEmpty, VarPtr(memWriter)
Private Function SaveWebP_ToMemory(ByRef srcDIB As pdDIB, ByRef srcOptions As String, ByRef dstMemWriter As PD_WebPMemoryWriter) As Boolean

    Const FUNC_NAME As String = "SaveWebP_ToMemory"
    SaveWebP_ToMemory = False
    
    'Prepare a compression options parser
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    cSettings.SetParamString srcOptions
    
    'Retrieve save quality (as it informs other settings)
    Dim saveQuality As Single
    saveQuality = cSettings.GetSingle("webp-quality", 100!, True)
    If (saveQuality < 0!) Then saveQuality = 0
    If (saveQuality > 100!) Then saveQuality = 100!
    
    'We now need a WebP config object.  Initializing this is long and complicated, so we
    ' hand it off to a dedicated initialization function.
    Dim saveConfig As PD_WebPConfig
    If (Not GetWebPConfig(cSettings, saveConfig)) Then
        InternalError FUNC_NAME, "config object may be bad!"
        GoTo SafeDeallocate
    End If
    
    'Config struct looks OK
    
    'PD sometimes needs to force use of data in RGBA format (animation files require this, for example).
    ' RGBA images are processed a little differently, because we supply direct pointers to the underlying
    ' PD object (instead of importing them via a libwebp function that auto-converts to YUV).
    Dim forceRGBA As Boolean
    forceRGBA = cSettings.GetBool("force-rgba", False, True)
    
    'WebP docs also suggest using RGBA for all lossless encoding (which makes sense if you're supplying
    ' the original image bytes in RGBA format!)
    forceRGBA = forceRGBA Or (saveQuality >= 100!)
    
    'Next we need to create a WebPPicture struct, which will hold all necessary information related to
    ' the underlying image we're trying to encode.  An external function handles this for us.
    Dim wpPicture As PD_WebPPicture
    If (Not GetWebPPicture(srcDIB, wpPicture, forceRGBA)) Then
        InternalError FUNC_NAME, "bad picture init"
        GoTo SafeDeallocate
    End If
    
    'With a picture object initialized, we now need a "memory writer" struct (which defines where the
    ' encoded WebP stream gets stored in-memory).
    
    'libwebp can use a custom callback for writing WebP data to file.  This is complicated to make work
    ' in VB6 because there's no macro for the calling-convention for this callback (and implementing cdecl
    ' callbacks in VB6 - while not impossible (see https://github.com/thetrik/VBCDeclFix) - is problematic)
    ' but fortunately, the authors of libwebp provide a workaround by plugging-in a library-managed struct
    ' that handles the callback for you.  The only downside of this approach is that you lose progress
    ' callbacks (not a huge deal) and the destination memory is obviously libwebp-managed, so you have to
    ' copy it into your own target after compression completes (also not a huge deal).
    
    'Anyway, with that limitation in mind, let's create the destination libwebp-managed memory writer.
    CallCDeclW WebPMemoryWriterInit, vbEmpty, VarPtr(dstMemWriter)
    
    'We now need to place the writer information inside the WebPPicture object (I know, it's baffling).
    wpPicture.ptrToWebPWriterFunction = m_ProcAddresses(WebPMemoryWrite)
    wpPicture.custom_ptr = VarPtr(dstMemWriter)
    
    'With the data safely imported, we can proceed with encoding
    If (CallCDeclW(WebPEncode, vbLong, VarPtr(saveConfig), VarPtr(wpPicture)) = 0) Then
        InternalError FUNC_NAME, "bad encode: " & wpPicture.error_code
        GoTo SafeDeallocate
    Else
        
        'Encoding was successful!  memWriter.mem_buffer now points at the compressed WebP bytes.
        ' It is up to the caller to handle things from here.
        SaveWebP_ToMemory = True
        
    End If
    
SafeDeallocate:
    CallCDeclW WebPPictureFree, vbEmpty, VarPtr(wpPicture)
    
End Function

'Given a source image, populate a WebPPicture struct.  Note that a memory writer is *NOT* initialized
' by this function - you must manually populate the WebPPicture object with that information as needed.
Private Function GetWebPPicture(ByRef srcDIB As pdDIB, ByRef dstPicture As PD_WebPPicture, Optional ByVal useRGBAMode As Boolean = True) As Boolean
    
    Const FUNC_NAME As String = "GetWebPPicture"
    GetWebPPicture = False
    
    If (CallCDeclW(WebPPictureInitInternal, vbLong, VarPtr(dstPicture), WEBP_ENCODER_ABI_VERSION) = 0) Then
        InternalError FUNC_NAME, "bad WebPPictureInitInternal"
        Exit Function
    End If
    
    'Populate some elements before initializing
    With dstPicture
        
        'RGBA is recommended for lossless compression and animation frames (which do not work in YUV);
        ' anything else (including near-lossless single frames) should use YUV for improved compression.
        If useRGBAMode Then .use_argb = 1 Else .use_argb = 0
        .webp_width = srcDIB.GetDIBWidth
        .webp_height = srcDIB.GetDIBHeight
        
    End With
    
    'After initializing the WebP object, we now need to pass it our RGB data.  Dedicated import functions
    ' exist for this (and they handle swizzling, which is especially useful, since there's no fast way to
    ' swizzle in VB6).
    Dim fReturn As Long
    If (srcDIB.GetDIBColorDepth = 32) Then
        fReturn = CallCDeclW(WebPPictureImportBGRA, vbLong, VarPtr(dstPicture), srcDIB.GetDIBPointer, srcDIB.GetDIBStride)
    Else
        fReturn = CallCDeclW(WebPPictureImportBGR, vbLong, VarPtr(dstPicture), srcDIB.GetDIBPointer, srcDIB.GetDIBStride)
    End If
    
    If (fReturn = 0) Then
        InternalError FUNC_NAME, "bad import"
        Exit Function
    Else
        GetWebPPicture = True
    End If
    
    'TODO: test if alpha channel is ignored automatically or not; if it isn't, we can possibly
    ' test and flag automatically using WebP APIs:
    '// Scan the picture 'picture' for the presence of non fully opaque alpha values.
    '// Returns true in such case. Otherwise returns false (indicating that the
    '// alpha plane can be ignored altogether e.g.).
    'WEBP_EXTERN int WebPPictureHasTransparency(const WebPPicture* picture);
    '
    '// Remove the transparency information (if present) by blending the color with
    '// the background color 'background_rgb' (specified as 24bit RGB triplet).
    '// After this call, all alpha values are reset to 0xff.
    'WEBP_EXTERN void WebPBlendAlpha(WebPPicture* pic, uint32_t background_rgb);
    
End Function

'Given a PD-specific serialization object, return a fully initialized (and validated) WebPConfig object
Private Function GetWebPConfig(ByRef cSettings As pdSerialize, ByRef dstConfig As PD_WebPConfig) As Boolean
    
    Const FUNC_NAME As String = "GetWebPConfig"
    GetWebPConfig = False
    
    Dim saveQuality As Single
    saveQuality = cSettings.GetSingle("webp-quality", 100!, True)
    If (saveQuality < 0!) Then saveQuality = 0!
    If (saveQuality > 100!) Then saveQuality = 100!
    
    Dim saveSpeed As PD_PerformanceSetting
    Select Case cSettings.GetString("webp-compression", "default", True)
        Case "fast"
            saveSpeed = PD_PERF_FASTEST
        Case "slow"
            saveSpeed = PD_PERF_BESTQUALITY
        Case Else
            saveSpeed = PD_PERF_BALANCED
    End Select
    
    'Presets and hints are supported, but it's unclear how useful these are in real-world usage.
    ' (Note that the options string can contain both a "preset" and "hint" value.  The "preset" is used
    ' when INITIALIZING a config object, while the "hint" is actually stored WITHIN the config struct.
    ' Obnoxiously, these are DIFFERENT ENUMS with DIFFERENT values smdh.)
    '
    'Anyway, all this is to say that we will use the preset up-front (during initialization), but the
    ' hint won't be used until deeper into this function.
    Dim imgPreset As PD_WebPPreset, imgHint As PD_WebPImageHint
    imgPreset = WEBP_PRESET_DEFAULT
    imgHint = WEBP_HINT_DEFAULT
    
    Select Case cSettings.GetString("webp-hint", vbNullString, True)
        Case "photo-indoor"
            imgPreset = WEBP_PRESET_PICTURE
            imgHint = WEBP_HINT_PICTURE
        Case "photo-outdoor"
            imgPreset = WEBP_PRESET_PHOTO
            imgHint = WEBP_HINT_PHOTO
        Case "art"
            imgPreset = WEBP_PRESET_DRAWING
            imgHint = WEBP_HINT_DEFAULT
        Case "chart"
            imgPreset = WEBP_PRESET_DRAWING
            imgHint = WEBP_HINT_GRAPH
        Case "icon"
            imgPreset = WEBP_PRESET_ICON
            imgHint = WEBP_HINT_GRAPH
        Case "text"
            imgPreset = WEBP_PRESET_TEXT
            imgHint = WEBP_HINT_DEFAULT
    End Select
    
    'Initialize a blank options struct, and note that we need to pass the expected save quality as part
    ' of initialization.
    If (CallCDeclW(WebPConfigInit, vbLong, VarPtr(dstConfig), imgPreset, CSng(saveQuality), WEBP_ENCODER_ABI_VERSION) = 0) Then InternalError FUNC_NAME, "bad WebPConfigInit"
    
    'Modify any desired settings in the config struct.  Note that these mappings are adopted from
    ' the official Google-sponsored WebP plugin for Photoshop (https://github.com/webmproject/WebPShop)
    
    '98+ map to lossless mode (or its close relative, "near-lossless" mode)
    Const NEAR_LOSSLESS_STARTS_AT As Single = 98!
    If (saveQuality >= NEAR_LOSSLESS_STARTS_AT) Then
        dstConfig.lossless = 1
        If (saveQuality = NEAR_LOSSLESS_STARTS_AT) Then
            dstConfig.near_lossless = 60
        ElseIf (saveQuality = NEAR_LOSSLESS_STARTS_AT + 1) Then
            dstConfig.near_lossless = 80
        Else
            dstConfig.near_lossless = 100
        End If
        If (saveSpeed = PD_PERF_FASTEST) Then
            dstConfig.quality = 0
        ElseIf (saveSpeed = PD_PERF_BALANCED) Then
            dstConfig.quality = 75!
        Else
            dstConfig.quality = 100!
        End If
    
    'lossy mode
    Else
        
        dstConfig.lossless = 0
        
        'map input range [0, 97] to output range [0, 100]
        dstConfig.quality = saveQuality * (100! / (NEAR_LOSSLESS_STARTS_AT - 1!))
        dstConfig.use_sharp_yuv = IIf(saveSpeed = PD_PERF_BESTQUALITY, 1, 0)
        
    End If
    
    'We also need to manually set the "speed" of the encoding; slower speeds can compress better,
    ' but when they say "slow", they mean it!  (Seriously - slow means SLOW in this context.)
    Select Case saveSpeed
        Case PD_PERF_FASTEST
            dstConfig.compressMethod = 0
        Case PD_PERF_BESTQUALITY
            dstConfig.compressMethod = 6
        Case Else
            dstConfig.compressMethod = 4
    End Select
    
    'Per the WebP docs, "Low alpha qualities are terrible on gradients: limit to acceptable range."
    ' (From WebPShopEncodeUtils.cpp in WebPShop)
    dstConfig.alpha_quality = 50! + (saveQuality / 2!)
    If (dstConfig.alpha_quality > 100!) Then dstConfig.alpha_quality = 100!
    
    'We calculated hint earlier; note that it generally only matters for lossless mode, but supplying
    ' it regardless causes no harm
    dstConfig.image_hint = imgHint
    
    'Multi-threading is (generally) fine, but we avoid it on single-core systems
    If (OS.LogicalCoreCount > 1) Then dstConfig.thread_level = 1 Else dstConfig.thread_level = 0
    
    'Validate the config struct before proceeding
    If (CallCDeclW(WebPValidateConfig, vbLong, VarPtr(dstConfig)) <> 0) Then
        GetWebPConfig = True
    Else
        InternalError FUNC_NAME, "bad config struct"
        GetWebPConfig = False
    End If
    
End Function

'Parse and import all frames from an animated WebP image.
' DO NOT CALL without verifying that the passed image is actually an *ANIMATED* WebP image.
Private Function LoadAnimatedWebP(ByRef srcFile As String, ByVal pData As Long, ByVal sizeOfDataInBytes As Long, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB) As Boolean
    
    LoadAnimatedWebP = False
    
    'Due to the design of libwebp (and animation functions being shunted into an entirely different
    ' library, libwebpdemuxer), animated images require a unique import process.
    Const FUNC_NAME As String = "LoadAnimatedWebP"
    
    'If a demuxer hasn't been allocated, allocate one now.  (Note that there is no harm in calling
    ' this if a demuxer already exists - we'll simply reuse the current class-level one.)
    If (Not InitializeDemuxer(pData, sizeOfDataInBytes)) Then
        InternalError FUNC_NAME, "couldn't allocate demuxer; abandoning animation import"
        LoadAnimatedWebP = False
        Exit Function
    End If
    
    'We are now guaranteed a valid demuxer in m_hDemuxer.
    
    'Retrieve the number of frames and perform basic sanity checks.
    Dim numFrames As Long
    numFrames = CallCDeclW(WebPDemuxGetI, vbLong, m_hDemuxer, WEBP_FF_FRAME_COUNT)
    If (numFrames > 1) Then
    
        'Next, we need to initialize an animation decoder options struct.  The library suggests
        ' allowing the library to initialize the struct for us, which... whatever, I guess?
        Dim animDecodeOptions As PD_WebPAnimDecoderOptions
        If (CallCDeclW(WebPAnimDecoderOptionsInit, vbLong, VarPtr(animDecodeOptions), WEBP_DEMUX_ABI_VERSION) = 0) Then
            InternalError FUNC_NAME, "WebPAnimDecoderOptionsInit"
            GoTo AnimationBroken
        End If
        
        'Set the animation decoder options to decode directly to premultiplied BGRA...
        animDecodeOptions.color_mode = MODE_bgrAp
        
        '...and explicitly allow multi-thread decoding on relevant systems
        If (OS.LogicalCoreCount > 1) Then animDecodeOptions.use_threads = 1 Else animDecodeOptions.use_threads = 0
        
        '...and finally, use the options struct to initialize an actual decoder object
        Dim unionPtrAndLen(0 To 1) As Long
        unionPtrAndLen(0) = pData
        unionPtrAndLen(1) = sizeOfDataInBytes
        
        Dim animDecoder As Long
        animDecoder = CallCDeclW(WebPAnimDecoderNew, vbLong, VarPtr(unionPtrAndLen(0)), VarPtr(animDecodeOptions), WEBP_DEMUX_ABI_VERSION)
        
        'Further animation decoding requires a valid decoder; without one, we have to abandon ship
        If (animDecoder = 0) Then
            InternalError FUNC_NAME, "WebPAnimDecoderNew"
            GoTo AnimationBroken
        End If
        
        'Retrieve global animation info, like final canvas size and loop count
        If (CallCDeclW(WebPAnimDecoderGetInfo, vbLong, animDecoder, VarPtr(m_AnimationInfo)) = 0) Then
            InternalError FUNC_NAME, "WebPAnimDecoderGetInfo"
            GoTo AnimationBroken
        End If
        
        If WEBP_DEBUG_VERBOSE Then PDDebug.LogAction "animation frame count: " & m_AnimationInfo.frame_count & ", loop count: " & m_AnimationInfo.loop_count & ", bgcolor: " & Hex$(m_AnimationInfo.bgcolor)
        
        'The animation canvas size should match the original size retrieve in our parent function,
        ' but if (for some reason?) they don't, we'll defer to the animation canvas settings.
        dstImage.Width = m_AnimationInfo.canvas_width
        dstImage.Height = m_AnimationInfo.canvas_height
        
        'WebP files do not provide a direct way to store DPI (although you could hide DPI settings in
        ' EXIF or XMP metadata, if you really wanted to), but PD expects DPI values on all incoming
        ' pdImage objects - so set arbitrary defaults now.
        dstImage.SetDPI 72, 72
        
        'We also need to flag the underlying format in advance, since it changes the way layer
        ' names are assigned (animation layers are called "frames" instead of "pages")
        dstImage.SetOriginalFileFormat PDIF_WEBP
        
        'We also need to store the animation loop count inside the parent object
        dstImage.ImgStorage.AddEntry "animation-loop-count", Trim$(Str$(m_AnimationInfo.loop_count))
        
        'With all global options stored, we can now start iterating frames.  (Individual frame times
        ' will be retrieved as part of this process.)
        Dim idxFrame As Long, numFramesOK As Long
        For idxFrame = 0 To m_AnimationInfo.frame_count - 1
            
            'Yes, this text uses "page" instead of "frame" - this is purely to reduce localization burdens
            Message "Loading page %1 of %2...", CStr(idxFrame + 1), m_AnimationInfo.frame_count, "DONOTLOG"
            
            'Failsafe check to ensure the decoder has more frames for us
            If (CallCDeclW(WebPAnimDecoderHasMoreFrames, vbLong, animDecoder) <> 0) Then
                
                'Get the next frame.  Two values are returned: a pointer to a constructed 32-bpp buffer
                ' of size canvaswidth * canvasheight (managed by libwebpdemux), and a timestamp for this
                ' frame (in ms, like APNG, not cs like GIF)
                Dim ptrPixels As Long, frameTimestamp As Long, lastFrameTimestamp As Long, frameTimeInMS As Long
                ptrPixels = 0: frameTimestamp = lastFrameTimestamp
                
                If (CallCDeclW(WebPAnimDecoderGetNext, vbLong, animDecoder, VarPtr(ptrPixels), VarPtr(frameTimestamp)) <> 0) Then
                    
                    'Success!  Create a new layer in the destination image, then copy the pixel data and
                    ' timestamp into it.
                    Dim newLayerID As Long, newLayerName As String, tmpLayer As pdLayer, tmpDIB As pdDIB
                    newLayerID = dstImage.CreateBlankLayer()
                    Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                    newLayerName = Layers.GenerateInitialLayerName(vbNullString, vbNullString, True, dstImage, dstDIB, idxFrame)
                    tmpLayer.InitializeNewLayer PDL_Image, newLayerName, Nothing, True
                    tmpLayer.SetLayerVisibility (idxFrame = 0)
                    
                    Set tmpDIB = New pdDIB
                    tmpDIB.CreateBlank m_AnimationInfo.canvas_width, m_AnimationInfo.canvas_height, 32, 0, 0
                    CopyMemoryStrict tmpDIB.GetDIBPointer, ptrPixels, tmpDIB.GetDIBStride * tmpDIB.GetDIBHeight
                    
                    'We explicitly request premultiplied data, so we don't need to manually premultiply now
                    tmpDIB.SetInitialAlphaPremultiplicationState True
                    
                    'Assign the finished layer image
                    tmpLayer.SetLayerDIB tmpDIB
                    
                    'As part of storing frametime, update the layer's name with ([time] ms) at the end
                    frameTimeInMS = frameTimestamp - lastFrameTimestamp
                    tmpLayer.SetLayerFrameTimeInMS frameTimeInMS
                    tmpLayer.SetLayerName tmpLayer.GetLayerName & " (" & CStr(frameTimeInMS) & "ms)"
                    lastFrameTimestamp = frameTimestamp
                    tmpLayer.NotifyOfDestructiveChanges
                    
                    'Track how many frames we've successfully loaded
                    numFramesOK = numFramesOK + 1
                
                'Couldn't get next frame
                Else
                    InternalError FUNC_NAME, "WebPAnimDecoderGetNext: unexpected fail"
                End If
            
            '/decoder doesn't have more frames, despite all frames not being read
            Else
                InternalError FUNC_NAME, "WebPAnimDecoderHasMoreFrames: unexpected fail"
            End If
                
        Next idxFrame
        
        'Reset the underlying iterator before continuing
        CallCDeclW WebPAnimDecoderReset, vbLong, animDecoder
        
        'Report success if at least one frame was retrieved correctly
        LoadAnimatedWebP = (numFramesOK > 0)
        If LoadAnimatedWebP Then dstImage.NotifyImageChanged UNDO_Everything
        
    '/animation frames <= 1
    Else
        InternalError FUNC_NAME, "only one animation frame found; reverting to static load process"
    End If
    
'Animated files can break in a lot (seriously, a LOT) of ways.  Because VB doesn't have normal try/catch syntax,
' we're stuck with GOTOs.  Broken animation files will divert here, which allows us to continue loading the
' file as a non-animated image while also freeing any animation objects we created along the way.
AnimationBroken:

    'Free a decoder object, if any
    If (animDecoder <> 0) Then
        CallCDeclW WebPAnimDecoderDelete, vbEmpty, animDecoder
        animDecoder = 0
    End If
    
    'Free the underlying demuxer (it isn't required by the single-frame loader)
    FreeDemuxer
    
End Function

Private Sub FreeDemuxer()
    If (m_hDemuxer <> 0) Then
        CallCDeclW WebPDemuxDelete, vbEmpty, m_hDemuxer
        m_hDemuxer = 0
    End If
End Sub

Private Function InitializeDemuxer(ByVal pWebPData As Long, ByVal pWebPDataLenB As Long) As Boolean
    
    'If we already have a demuxer, reuse it
    If (m_hDemuxer <> 0) Then
        InitializeDemuxer = True
    Else
        
        'The demux API takes a *WebPData as input; this is a simple union of data pointer and size
        Dim webpDataStruct(0 To 1) As Long
        webpDataStruct(0) = pWebPData
        webpDataStruct(1) = pWebPDataLenB
        
        m_hDemuxer = CallCDeclW(WebPDemux, vbLong, VarPtr(webpDataStruct(0)), 0&, 0&, WEBP_DEMUX_ABI_VERSION)
        InitializeDemuxer = (m_hDemuxer <> 0)
        If WEBP_DEBUG_VERBOSE Then PDDebug.LogAction "Allocated demux object: " & m_hDemuxer
        
    End If
    
End Function

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As PD_WebP_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

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

Private Sub InternalError(ByVal funcName As String, Optional ByRef errString As String = vbNullString, Optional ByVal libReturn As PD_VP8StatusCode = VP8_STATUS_OK)
    funcName = "pdWebP." & funcName & "() "
    If (libReturn <> VP8_STATUS_OK) Then
        PDDebug.LogAction funcName & "returned error #" & libReturn & "(" & errString & ")", PDM_External_Lib
    Else
        PDDebug.LogAction funcName & "error: " & errString, PDM_External_Lib
    End If
End Sub

Private Sub Class_Initialize()

    'Pre-load all relevant proc addresses
    ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
    m_ProcAddresses(WebPAnimDecoderDelete) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderDelete")
    m_ProcAddresses(WebPAnimDecoderGetInfo) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderGetInfo")
    m_ProcAddresses(WebPAnimDecoderGetNext) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderGetNext")
    m_ProcAddresses(WebPAnimDecoderHasMoreFrames) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderHasMoreFrames")
    m_ProcAddresses(WebPAnimDecoderNew) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderNewInternal")
    m_ProcAddresses(WebPAnimDecoderOptionsInit) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderOptionsInitInternal")
    m_ProcAddresses(WebPAnimDecoderReset) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPAnimDecoderReset")
    m_ProcAddresses(WebPAnimEncoderAdd) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderAdd")
    m_ProcAddresses(WebPAnimEncoderAssemble) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderAssemble")
    m_ProcAddresses(WebPAnimEncoderDelete) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderDelete")
    m_ProcAddresses(WebPAnimEncoderGetError) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderGetError")
    m_ProcAddresses(WebPAnimEncoderNew) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderNewInternal")
    m_ProcAddresses(WebPAnimEncoderOptionsInit) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPMux(), "WebPAnimEncoderOptionsInitInternal")
    m_ProcAddresses(WebPBlendAlpha) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPBlendAlpha")
    m_ProcAddresses(WebPCleanupTransparentArea) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPCleanupTransparentArea")
    m_ProcAddresses(WebPConfigInit) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPConfigInitInternal")
    m_ProcAddresses(WebPDecodeBGRAInto) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPDecodeBGRAInto")
    m_ProcAddresses(WebPDemux) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPDemuxInternal")
    m_ProcAddresses(WebPDemuxDelete) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPDemuxDelete")
    m_ProcAddresses(WebPDemuxGetI) = GetProcAddress(Plugin_WebP.GetHandle_LibWebPDemux(), "WebPDemuxGetI")
    m_ProcAddresses(WebPEncode) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPEncode")
    m_ProcAddresses(WebPFree) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPFree")
    m_ProcAddresses(WebPGetDecoderVersion) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPGetDecoderVersion")
    m_ProcAddresses(WebPGetFeatures) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPGetFeaturesInternal")
    m_ProcAddresses(WebPGetInfo) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPGetInfo")
    m_ProcAddresses(WebPMemoryWrite) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPMemoryWrite")
    m_ProcAddresses(WebPMemoryWriterClear) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPMemoryWriterClear")
    m_ProcAddresses(WebPMemoryWriterInit) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPMemoryWriterInit")
    m_ProcAddresses(WebPPictureAlloc) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureAlloc")
    m_ProcAddresses(WebPPictureARGBToYUVA) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureARGBToYUVA")
    m_ProcAddresses(WebPPictureARGBToYUVADithered) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureARGBToYUVADithered")
    m_ProcAddresses(WebPPictureCopy) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureCopy")
    m_ProcAddresses(WebPPictureCrop) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureCrop")
    m_ProcAddresses(WebPPictureDistortion) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureDistortion")
    m_ProcAddresses(WebPPictureFree) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureFree")
    m_ProcAddresses(WebPPictureHasTransparency) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureHasTransparency")
    m_ProcAddresses(WebPPictureImportBGR) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportBGR")
    m_ProcAddresses(WebPPictureImportBGRA) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportBGRA")
    m_ProcAddresses(WebPPictureImportBGRX) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportBGRX")
    m_ProcAddresses(WebPPictureImportRGB) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportRGB")
    m_ProcAddresses(WebPPictureImportRGBA) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportRGBA")
    m_ProcAddresses(WebPPictureImportRGBX) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureImportRGBX")
    m_ProcAddresses(WebPPictureInitInternal) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureInitInternal")
    m_ProcAddresses(WebPPictureIsView) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureIsView")
    m_ProcAddresses(WebPPictureRescale) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureRescale")
    m_ProcAddresses(WebPPictureSharpARGBToYUVA) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureSharpARGBToYUVA")
    m_ProcAddresses(WebPPictureSmartARGBToYUVA) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureSmartARGBToYUVA")
    m_ProcAddresses(WebPPictureView) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureView")
    m_ProcAddresses(WebPPictureYUVAToARGB) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPictureYUVAToARGB")
    m_ProcAddresses(WebPPlaneDistortion) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPPlaneDistortion")
    m_ProcAddresses(WebPValidateConfig) = GetProcAddress(Plugin_WebP.GetHandle_LibWebP(), "WebPValidateConfig")
    
    'Initialize all module-level arrays
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
    
End Sub

Private Sub Class_Terminate()
    FreeDemuxer
End Sub
