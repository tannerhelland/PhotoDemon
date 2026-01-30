Attribute VB_Name = "Public_EnumsAndTypes"
Option Explicit

'Standard WAPI types
Public Type PointAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectF_RB
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Type winRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Currently supported tools; these numbers correspond to the index of the tool's command button on the main form.
' In theory, adding new tools should be as easy as changing these numbers.  All tool-related code is tied into these
' constants, so any changes here should automatically propagate throughout the software.  (In practice, be sure to
' double-check everything!)
Public Enum PDTools
    TOOL_UNDEFINED = -1
    
    NAV_DRAG = 0
    NAV_ZOOM = 1
    NAV_MOVE = 2
    COLOR_PICKER = 3
    ND_MEASURE = 4
    ND_CROP = 5
    
    SELECT_RECT = 6
    SELECT_CIRC = 7
    SELECT_POLYGON = 8
    SELECT_LASSO = 9
    SELECT_WAND = 10
    
    TEXT_BASIC = 11
    TEXT_ADVANCED = 12
    
    PAINT_PENCIL = 13
    PAINT_SOFTBRUSH = 14
    PAINT_ERASER = 15
    PAINT_CLONE = 16
    PAINT_FILL = 17
    PAINT_GRADIENT = 18
End Enum

#If False Then
    Private Const TOOL_UNDEFINED = -1, NAV_DRAG = 0, NAV_ZOOM = 1, NAV_MOVE = 2, COLOR_PICKER = 3, ND_MEASURE = 4, ND_CROP = 5
    Private Const SELECT_RECT = 6, SELECT_CIRC = 7, SELECT_POLYGON = 8, SELECT_LASSO = 9, SELECT_WAND = 10
    Private Const TEXT_BASIC = 11, TEXT_ADVANCED = 12
    Private Const PAINT_PENCIL = 13, PAINT_SOFTBRUSH = 14, PAINT_ERASER = 15, PAINT_CLONE = 16, PAINT_FILL = 17, PAINT_GRADIENT = 18
#End If

'Currently supported file tools; these numbers correspond to the index of the tool's command button on the main form.
' In theory, adding new tools should be as easy as changing these numbers.  All file-tool-related code is tied into
' these constants, so any changes here should automatically propagate throughout the software.  (In practice, be sure
' to double-check everything!)
Public Enum PDFileTools
    FILE_NEW = 0
    FILE_OPEN = 1
    FILE_CLOSE = 2
    FILE_SAVE = 3
    FILE_SAVEAS_LAYERS = 4
    FILE_SAVEAS_FLAT = 5
    FILE_UNDO = 6
    FILE_REDO = 7
End Enum

#If False Then
    Private Const FILE_NEW = 0, FILE_OPEN = 1, FILE_CLOSE = 2, FILE_SAVE = 3, FILE_SAVEAS_LAYERS = 4, FILE_SAVEAS_FLAT = 5, FILE_UNDO = 6, FILE_REDO = 7
#End If

'How should a selection be rendered?
Public Enum PD_SelectionRender
    PDSR_Highlight = 0
    PDSR_Lightbox = 1
    PDSR_Ants = 2
    PDSR_Outline = 3
End Enum

#If False Then
    Private Const PDSR_Highlight = 0, PDSR_Lightbox = 1, PDSR_Ants = 2, PDSR_Outline = 3
#End If

'PhotoDemon's language files provide a small amount of metadata to help the program know how to use them.  This type
' was previously declared inside the pdTranslate class, but with the addition of a Language Editor, I have moved it
' here, so the entire project can access the type.
Public Type PDLanguageFile
    Author As String
    FileName As String
    langID As String
    LangName As String
    LangType As String
    langVersion As String
    LangStatus As String
    InternalDisplayName As String
    UpdateChecksum As Long
    IsOfficial As Boolean
End Type

'Replacement mouse button type.  VB doesn't report X-button clicks in their native button type, but PD does.  Whether
' this is useful is anybody's guess, but it doesn't hurt to have... right?  Also, note that the left/middle/right button
' values are identical to VB, so existing code won't break if using this enum against VB's standard mouse constants.
Public Enum PDMouseButtonConstants
    pdLeftButton = 1
    pdRightButton = 2
    pdMiddleButton = 4
    pdXButtonOne = 8
    pdXButtonTwo = 16
End Enum

#If False Then
    Private Const pdLeftButton = 1, pdRightButton = 2, pdMiddleButton = 4, pdXButtonOne = 8, pdXButtonTwo = 16
#End If

'Supported save events.  To try and handle workflow issues gracefully, PhotoDemon will track image save state for a few
' different save events.  See the pdImage function setSaveState for details.
Public Enum PD_SAVE_EVENT
    pdSE_AnySave = 0        'Any type of save event; used to set the enabled state of the main toolbar's Save button
    pdSE_SavePDI = 1        'Image has been saved to PDI format in its current state
    pdSE_SaveFlat = 2       'Image has been saved to some flattened format (JPEG, PNG, etc) in its current state
End Enum

#If False Then
    Private Const pdSE_AnySave = 0, pdSE_SavePDI = 1, pdSE_SaveFlat = 2
#End If

'Edge-handling methods for distort-style filters
Public Enum PD_EdgeOperator
    pdeo_Clamp = 0
    pdeo_Reflect = 1
    pdeo_Wrap = 2
    pdeo_Erase = 3
    pdeo_Original = 4
End Enum

#If False Then
    Private Const pdeo_Clamp = 0, pdeo_Reflect = 1, pdeo_Wrap = 2, pdeo_Erase = 3, pdeo_Original = 4
#End If

'Orientation (used in a whole bunch of different situations)
Public Enum PD_Orientation
    pdo_Horizontal = 0
    pdo_Vertical = 1
    pdo_Both = 2
End Enum

#If False Then
    Private Const pdo_Horizontal = 0, pdo_Vertical = 1, pdo_Both = 2
#End If

'Some PhotoDemon actions can operate on the whole image, or on just a specific layer (e.g. resize).  When initiating
' one of these dual-action operations, the constants below can be used to specify the mode.
Public Enum PD_ActionTarget
    pdat_Image = 0
    pdat_SingleLayer = 1
End Enum

#If False Then
    Private Const pdat_Image = 0, pdat_SingleLayer = 1
#End If

'When an action triggers the creation of Undo/Redo data, it must specify what kind of Undo/Redo data it wants created.
' This type is used by PD to determine the most efficient way to store/restore previous actions.
Public Enum PD_UndoType
    UNDO_Nothing = -1
    UNDO_Everything = 0
    UNDO_Image = 1
    UNDO_Image_VectorSafe = 2
    UNDO_ImageHeader = 3
    UNDO_Layer = 4
    UNDO_Layer_VectorSafe = 5
    UNDO_LayerHeader = 6
    UNDO_Selection = 7
End Enum

#If False Then
    Private Const UNDO_Nothing = -1, UNDO_Everything = 0, UNDO_Image = 1, UNDO_Image_VectorSafe = 2, UNDO_ImageHeader = 3
    Private Const UNDO_Layer = 4, UNDO_Layer_VectorSafe = 5, UNDO_LayerHeader = 5, UNDO_Selection = 7
#End If

'Enums for App Command messages, which are (optionally) returned by the pdInput class
Public Enum AppCommandConstants
    AC_BROWSER_BACKWARD = 1
    AC_BROWSER_FORWARD = 2
    AC_BROWSER_REFRESH = 3
    AC_BROWSER_STOP = 4
    AC_BROWSER_SEARCH = 5
    AC_BROWSER_FAVORITES = 6
    AC_BROWSER_HOME = 7
    AC_VOLUME_MUTE = 8
    AC_VOLUME_DOWN = 9
    AC_VOLUME_UP = 10
    AC_MEDIA_NEXTTRACK = 11
    AC_MEDIA_PREVIOUSTRACK = 12
    AC_MEDIA_STOP = 13
    AC_MEDIA_PLAY_PAUSE = 14
    AC_LAUNCH_MAIL = 15
    AC_LAUNCH_MEDIA_SELECT = 16
    AC_LAUNCH_APP1 = 17
    AC_LAUNCH_APP2 = 18
    AC_BASS_DOWN = 19
    AC_BASS_BOOST = 20
    AC_BASS_UP = 21
    AC_TREBLE_DOWN = 22
    AC_TREBLE_UP = 23
    AC_MICROPHONE_VOLUME_MUTE = 24
    AC_MICROPHONE_VOLUME_DOWN = 25
    AC_MICROPHONE_VOLUME_UP = 26
    AC_HELP = 27
    AC_FIND = 28
    AC_NEW = 29
    AC_OPEN = 30
    AC_CLOSE = 31
    AC_SAVE = 32
    AC_PRINT = 33
    AC_UNDO = 34
    AC_REDO = 35
    AC_COPY = 36
    AC_CUT = 37
    AC_PASTE = 38
    AC_REPLY_TO_MAIL = 39
    AC_FORWARD_MAIL = 40
    AC_SEND_MAIL = 41
    AC_SPELL_CHECK = 42
    AC_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43
    AC_MIC_ON_OFF_TOGGLE = 44
    AC_CORRECTION_LIST = 45
End Enum

#If False Then
    Private Const AC_BROWSER_BACKWARD = 1, AC_BROWSER_FORWARD = 2, AC_BROWSER_REFRESH = 3, AC_BROWSER_STOP = 4, AC_BROWSER_SEARCH = 5, AC_BROWSER_FAVORITES = 6, AC_BROWSER_HOME = 7, AC_VOLUME_MUTE = 8, AC_VOLUME_DOWN = 9, AC_VOLUME_UP = 10, AC_MEDIA_NEXTTRACK = 11, AC_MEDIA_PREVIOUSTRACK = 12, AC_MEDIA_STOP = 13, _
    AC_MEDIA_PLAY_PAUSE = 14, AC_LAUNCH_MAIL = 15, AC_LAUNCH_MEDIA_SELECT = 16, AC_LAUNCH_APP1 = 17, AC_LAUNCH_APP2 = 18, AC_BASS_DOWN = 19, AC_BASS_BOOST = 20, AC_BASS_UP = 21, AC_TREBLE_DOWN = 22, AC_TREBLE_UP = 23, AC_MICROPHONE_VOLUME_MUTE = 24, AC_MICROPHONE_VOLUME_DOWN = 25, AC_MICROPHONE_VOLUME_UP = 26, _
    AC_HELP = 27, AC_FIND = 28, AC_NEW = 29, AC_OPEN = 30, AC_CLOSE = 31, AC_SAVE = 32, AC_PRINT = 33, AC_UNDO = 34, AC_REDO = 35, AC_COPY = 36, AC_CUT = 37, AC_PASTE = 38, AC_REPLY_TO_MAIL = 39, AC_FORWARD_MAIL = 40, AC_SEND_MAIL = 41, AC_SPELL_CHECK = 42, AC_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43, _
    AC_MIC_ON_OFF_TOGGLE = 44, AC_CORRECTION_LIST = 45
#End If

'Supported edge-detection algorithms
Public Enum PD_EdgeDetector
    pded_Contour = 0
    pded_Hilite = 1
    pded_Laplacian = 2
    pded_PhotoDemon = 3
    pded_Prewitt = 4
    pded_Roberts = 5
    pded_Sobel = 6
End Enum

#If False Then
    Private Const pded_Contour = 0, pded_Hilite = 1, pded_Laplacian = 2, pded_PhotoDemon = 3, pded_Prewitt = 4, pded_Roberts = 5, pded_Sobel = 6
#End If

Public Enum PD_EdgeDirection
    pded_All = 0
    pded_Horizontal = 1
    pded_Vertical = 2
End Enum

#If False Then
    Private Const pded_All = 0, pded_Horizontal = 1, pded_Vertical = 2
#End If

'PhotoDemon performance settings are generally provided in three groups: Max Quality, Balanced, and Max Performance
Public Enum PD_PerformanceSetting
    PD_PERF_BESTQUALITY = 0
    PD_PERF_BALANCED = 1
    PD_PERF_FASTEST = 2
End Enum

#If False Then
    Private Const PD_PERF_BESTQUALITY = 0, PD_PERF_BALANCED = 1, PD_PERF_FASTEST = 2
#End If

'PhotoDemon supports multiple image encoders and decoders.  Many of these are internal homebrew parsers,
' but sometimes we lean on 3rd-party libraries as relevant.  (Anything that ends in "_Parser" is homebrew.)
Public Enum PD_ImageDecoder
    id_Failure = -1
    id_GDIPlus = 0
    id_FreeImage
    id_CBZParser
    id_HDRParser
    id_HGTParser
    id_ICOParser
    id_MBMParser
    id_ORAParser
    id_PCXParser
    id_PDIParser
    id_PNGParser
    id_PSDParser
    id_PSPParser
    id_QOIParser
    id_WBMPParser
    id_XBMParser
    id_XCFParser
    id_WIC
    id_libavif
    id_CharLS
    id_libwebp
    id_resvg
    id_libjxl
    id_pdfium
    id_libheif
    id_DirectXTex
    id_OpenJPEG
End Enum

#If False Then
    Private Const id_Failure = -1, id_GDIPlus = 0, id_FreeImage = 0
    Private Const id_CBZParser = 0, id_HDRParser = 0, id_HGTParser = 0, id_ICOParser = 0, id_MBMParser = 0, id_ORAParser = 0, id_PDIParser = 0
    Private Const id_PNGParser = 0, id_PSDParser = 0, id_PSPParser = 0, id_QOIParser = 0, id_WBMPParser = 0, id_XBMParser = 0
    Private Const id_XCFParser = 0, id_WIC = 0, id_libavif = 0, id_CharLS = 0, id_libwebp = 0, id_resvg = 0, id_libjxl = 0
    Private Const id_pdfium = 0, id_libheif = 0, id_PCXParser = 0, id_DirectXTex = 0, id_OpenJPEG = 0
#End If

'Some UI DIBs are generated at run-time.  These DIBs can be requested by using the getRuntimeUIDIB() function.
Public Enum PD_RuntimeIcon
    pdri_ChannelRed = 0
    pdri_ChannelGreen = 1
    pdri_ChannelBlue = 2
    pdri_ChannelRGB = 3
    pdri_ArrowUp = 4
    pdri_ArrowUpR = 5
    pdri_ArrowRight = 6
    pdri_ArrowDownR = 7
    pdri_ArrowDown = 8
    pdri_ArrowDownL = 9
    pdri_ArrowLeft = 10
    pdri_ArrowUpL = 11
    pdri_Play = 12
    pdri_Pause = 13
End Enum

#If False Then
    Private Const pdri_ChannelRed = 0, pdri_ChannelGreen = 1, pdri_ChannelBlue = 2, pdri_ChannelRGB = 3, pdri_ArrowUp = 4, pdri_ArrowUpR = 5, pdri_ArrowRight = 6, pdri_ArrowDownR = 7, pdri_ArrowDown = 8, pdri_ArrowDownL = 9
    Private Const pdri_ArrowLeft = 10, pdri_ArrowUpL = 11, pdri_Play = 12, pdri_Pause = 13
#End If

'Metadata formats.  These are important when writing metadata to a file that is being saved to a different format
' from its original state (e.g. JPEG to PNG, which requires complicated metadata conversions).
Public Enum PD_METADATA_FORMAT
    PDMF_NONE = 0
    PDMF_EXIF = 1
    PDMF_IPTC = 2
    PDMF_XMP = 3
End Enum

#If False Then
    Private Const PDMF_NONE = 0, PDMF_EXIF = 1, PDMF_IPTC = 2, PDMF_XMP = 3
#End If

'Some options in PD support automatic enablement, contingent on various (hopefully) intelligent algorithms.
' Use this enum instead of raw Booleans if an algorithm is capable of self-setting certain settings.
' (Say that 10x fast :p)  Similarly, if an option has never been set, we can safely detect that case, too.
Public Enum PD_BOOL
    PD_BOOL_UNKNOWN = -1
    PD_BOOL_FALSE = 0
    PD_BOOL_TRUE = 1
    PD_BOOL_AUTO = 2
End Enum

#If False Then
    Private Const PD_BOOL_UNKNOWN = -1, PD_BOOL_FALSE = 0, PD_BOOL_TRUE = 1, PD_BOOL_AUTO = 2
#End If

'Tone-mapping is required for high bit-depth images.  PhotoDemon supports a variety of tone-map operations.
Public Enum PD_TONE_MAP
    PDTM_LINEAR = 0
    PDTM_FILMIC = 1
    PDTM_DRAGO = 2
    PDTM_REINHARD = 3
End Enum

#If False Then
    Private Const PDTM_LINEAR = 0, PDTM_FILMIC = 1, PDTM_DRAGO = 2, PDTM_REINHARD = 3
#End If

'Some operations need to return more detailed state than just FALSE/TRUE.  (For example, loading images via FreeImage.)
Public Enum PD_OPERATION_OUTCOME
    PD_SUCCESS = -1
    PD_FAILURE_GENERIC = 0
    PD_FAILURE_USER_CANCELED = 1
End Enum

#If False Then
    Private Const PD_SUCCESS = -1, PD_FAILURE_GENERIC = 0, PD_FAILURE_USER_CANCELED = 1
#End If

'How frequently to check for new versions of the app
Public Enum PD_UPDATE_FREQUENCY
    PDUF_EACH_SESSION = 0
    PDUF_WEEKLY = 1
    PDUF_MONTHLY = 2
    PDUF_NEVER = 3
End Enum

#If False Then
    Private Const PDUF_EACH_SESSION = 0, PDUF_WEEKLY = 1, PDUF_MONTHLY = 2, PDUF_NEVER = 3
#End If

'...and what kind of updates we are allowed to download automatically
Public Enum PD_UpdateTrack
    [ut_None] = -1
    ut_Stable = 0
    ut_Beta = 1
    ut_Developer = 2
End Enum

#If False Then
    Private Const ut_None = -1, ut_Stable = 0, ut_Beta = 1, ut_Developer = 2
#End If

'pdCompositor makes heavy use of level-of-detail (LOD) caches stored inside individual pdLayer objects.
' Callers need to identify compositor requests with one of these IDs, which tells the compositor which
' cache to preferentially use.  Correct LOD tags greatly improve performance, particularly on the
' primary canvas.
Public Enum PD_CompositorLOD
    CLC_Generic = 0
    CLC_Viewport = 1
    CLC_Thumbnail = 2
    CLC_Painting = 3
    CLC_ColorSample = 4
End Enum

#If False Then
    Private Const CLC_Generic = 0, CLC_Viewport = 1, CLC_Thumbnail = 2, CLC_Painting = 3, CLC_ColorSample = 4
#End If

Public Const NUM_OF_LOD_CACHES As Long = 5

'PD's gotten much better about abstracting and encapsulating clipboard-specific functionality.
' Unfortunately, some formats (most notably CF_BITMAP) require special heuristics from PD's
' image load function, because the alpha data CF_BITMAP provides is unlikely to be valid, but we
' can't know for sure without examining it.  As such, some clipboard-specific data can be
' retrieved via this struct.
Public Type PD_Clipboard_Info
    pdci_CurrentFormat As PredefinedClipboardFormatConstants
    pdci_OriginalFormat As PredefinedClipboardFormatConstants
    pdci_DIBv5AlphaMask As Long
End Type

'When iterating through pixels via pdPixelIterator, PD now supports a variety of region shapes.
Public Enum PD_PixelRegionShape
    PDPRS_Rectangle = 0
    PDPRS_Circle = 1
    PDPRS_Diamond = 2
End Enum

#If False Then
    Private Const PDPRS_Rectangle = 0, PDPRS_Circle = 1, PDPRS_Diamond = 2
#End If

'pdPixelIterator also supports multiple modes of operation, which determine what kind of histogram it generates.
Public Enum PD_PixelIteratorMode
    PDPIM_RGBA = 0
    PDPIM_Luminance = 1
    PDPIM_ByteArray = 2
End Enum

#If False Then
    Private Const PDPIM_RGBA = 0, PDPIM_Luminance = 1, PDPIM_ByteArray = 2
#End If

Public Enum PD_LuminanceMode
    PDLM_Value = 0
    PDLM_Lightness = 1
End Enum

#If False Then
    Private Const PDLM_Value = 0, PDLM_Lightness = 1
#End If

'List boxes support several different per-item height modes
Public Enum PD_ListboxHeight
    PDLH_Fixed = 0
    PDLH_Separators = 1
    PDLH_Custom = 2
End Enum

#If False Then
    Private Const PDLH_Fixed = 0, PDLH_Separators = 1, PDLH_Custom = 2
#End If

'PD's central list support class can also adjust its behavior automatically, depending on whether its being used by
' an underlying list box or a combo box.  (This primarily affects how the support class interprets things like
' mouse and key events; e.g. MouseWheel has a different meaning for a scrollable list vs a closed dropdown.)
Public Enum PD_ListSupportMode
    PDLM_ListBox = 0
    PDLM_DropDown = 1
    PDLM_LB_Inside_DD = 2
End Enum

#If False Then
    Private Const PDLM_ListBox = 0, PDLM_DropDown = 1, PDLM_LB_Inside_DD = 2
#End If

'Color definition.  If one of the non-BaseColor values is missing in the theme, it will be replaced by the
' BaseColor value.  (As such, the BaseColor value will always be present in a color definition.)
Public Type PDThemeColor
    baseColor As Long
    disabledColor As Long
    ActiveColor As Long
    hoverColor As Long
    ActiveHoverColor As Long
End Type

Public Type PDCachedColor
    OrigObjectName As String
    OrigColorName As String
    OrigColorValues As PDThemeColor
End Type

'Supported file formats.  Note that the import/export/feature availability of these formats is complex,
' and not always symmetrical (i.e. just because we read a given format doesn't mean we also write it).
' For details on a specific format's import/export available, refer to the ImageFormats module.
'
'For historical reasons, this list of formats was originally based off a matching list of FIF_ constants
' used by the 3rd-party FreeImage module.  FreeImage is no longer used for most formats in PD, but identifiers
' have been kept to avoid breaking old code paths.
Public Enum PD_IMAGE_FORMAT
    PDIF_UNKNOWN = -1
    PDIF_BMP = 0
    PDIF_ICO = 1    'internal parser used
    PDIF_JPEG = 2
    PDIF_JNG = 3
    PDIF_KOALA = 4
    PDIF_LBM = 5
    PDIF_IFF = PDIF_LBM
    PDIF_MNG = 6    'not currently supported
    PDIF_PBM = 7
    PDIF_PBMRAW = 8
    PDIF_PCD = 9
    PDIF_PCX = 10   'internal parser used
    PDIF_PGM = 11
    PDIF_PGMRAW = 12
    PDIF_PNG = 13   'internal parser used
    PDIF_PPM = 14
    PDIF_PPMRAW = 15
    PDIF_RAS = 16
    PDIF_TARGA = 17
    PDIF_TIFF = 18
    PDIF_WBMP = 19  'internal parser used
    PDIF_PSD = 20   'internal parser used
    PDIF_CUT = 21
    PDIF_XBM = 22   'internal parser used
    PDIF_XPM = 23
    PDIF_DDS = 24
    PDIF_GIF = 25   'internal parser used
    PDIF_HDR = 26
    PDIF_FAXG3 = 27
    PDIF_SGI = 28
    PDIF_EXR = 29
    PDIF_J2K = 30
    PDIF_JP2 = 31
    PDIF_PFM = 32
    PDIF_PICT = 33
    PDIF_RAW = 34
    PDIF_WEBP = 35  'FreeImage is *not* used to load/save WebP; we use libwebp
    PDIF_JXR = 36
   
    'PhotoDemon's internal PDI format identifier(s).
    PDIF_PDI = 100
    PDIF_RAWBUFFER = 101
    PDIF_TMPFILE = 102
    
    'Other image formats supported by PhotoDemon (these are unsupported by FreeImage and do not map to their FIF constants)
    PDIF_WMF = 110
    PDIF_EMF = 111
    PDIF_PNM = 112      'Catch-all for various portable pixmap filetypes
    PDIF_SVG = 113      'Support is currently experimental *only*!  Recommend disabling in production builds.
    PDIF_ORA = 114      'OpenRaster support was added in 8.0 nightly builds
    PDIF_HEIF = 115     'Loaded via libheif
    PDIF_MBM = 116      'Symbian MBM support (including .aif and similar variants) was added in 9.0
    PDIF_PSP = 117      'JASC/Corel PaintShop Pro support was added in 9.0
    PDIF_CBZ = 118      'Comic book archives (zip format only) were added in 9.0
    PDIF_AVIF = 119     'Modern high-efficiency images (AV1 frames) were added in 9.0
    PDIF_JLS = 120      'JPEG-LS was added in 9.0
    PDIF_QOI = 121      'QOI was added in 9.0
    PDIF_XCF = 122      'GIMP XCF was added in 9.0
    PDIF_HGT = 123      'Shuttle Radar Topography Mission (SRTM) import was added in 10.0 nightly builds
    PDIF_JXL = 124      'JPEG XL was added in 10.0 nightly builds
    PDIF_PDF = 125      'PDF was added in 10.0 nightly builds
    
    'AFTER ADDING A NEW FORMAT, ALWAYS UPDATE THE MAX ID VALUE.
    ' PhotoDemon uses this value to iterate known file formats against known extensions,
    ' and identify mismatched values.
    PDIF_MAX_ID = 125
    
End Enum

#If False Then
    Const PDIF_UNKNOWN = -1, PDIF_BMP = 0, PDIF_ICO = 1, PDIF_JPEG = 2, PDIF_JNG = 3, PDIF_KOALA = 4, PDIF_LBM = 5
    Const PDIF_IFF = PDIF_LBM, PDIF_MNG = 6, PDIF_PBM = 7, PDIF_PBMRAW = 8, PDIF_PCD = 9, PDIF_PCX = 10, PDIF_PGM = 11
    Const PDIF_PGMRAW = 12, PDIF_PNG = 13, PDIF_PPM = 14, PDIF_PPMRAW = 15, PDIF_RAS = 16, PDIF_TARGA = 17, PDIF_TIFF = 18
    Const PDIF_WBMP = 19, PDIF_PSD = 20, PDIF_CUT = 21, PDIF_XBM = 22, PDIF_XPM = 23, PDIF_DDS = 24, PDIF_GIF = 25
    Const PDIF_HDR = 26, PDIF_FAXG3 = 27, PDIF_SGI = 28, PDIF_EXR = 29, PDIF_J2K = 30, PDIF_JP2 = 31, PDIF_PFM = 32
    Const PDIF_PICT = 33, PDIF_RAW = 34, PDIF_WEBP = 35, PDIF_JXR = 36
    Const PDIF_PDI = 100, PDIF_RAWBUFFER = 101, PDIF_TMPFILE = 102
    Const PDIF_WMF = 110, PDIF_EMF = 111, PDIF_PNM = 112, PDIF_ORA = 114, PDIF_HEIF = 115, PDIF_MBM = 116, PDIF_PSP = 117
    Const PDIF_CBZ = 118, PDIF_AVIF = 119, PDIF_JLS = 120, PDIF_QOI = 121, PDIF_XCF = 122, PDIF_HGT = 123, PDIF_JXL = 124
    Const PDIF_PDF = 125, PDIF_MAX_ID = 125
#End If

'MSDN page: https://msdn.microsoft.com/en-us/library/windows/desktop/ms645603(v=vs.85).aspx
Public Type MOUSEMOVEPOINT
    x As Long
    y As Long
    ptTime As Long
    dwExtraInfo As Long
End Type

'PD color quantization methods.
Public Enum PD_COLOR_QUANTIZE
    PDCQ_MedianCut = 0
    PDCQ_Neuquant = 1
End Enum

#If False Then
    Private Const PDCQ_MedianCut = 0, PDCQ_Neuquant = 1
#End If

'Dithering methods.  All of these are implemented internally.
Public Enum PD_DITHER_METHOD
    PDDM_None = 0
    PDDM_Ordered_Bayer4x4 = 1
    PDDM_Ordered_Bayer8x8 = 2
    PDDM_SingleNeighbor = 3
    PDDM_FloydSteinberg = 4
    PDDM_JarvisJudiceNinke = 5
    PDDM_Stucki = 6
    PDDM_Burkes = 7
    PDDM_Sierra3 = 8
    PDDM_SierraTwoRow = 9
    PDDM_SierraLite = 10
    PDDM_Atkinson = 11
End Enum

#If False Then
    Private Const PDDM_None = 0, PDDM_Ordered_Bayer4x4 = 1, PDDM_Ordered_Bayer8x8 = 2, PDDM_SingleNeighbor = 3, PDDM_FloydSteinberg = 4, PDDM_JarvisJudiceNinke = 5, PDDM_Stucki = 6, PDDM_Burkes = 7, PDDM_Sierra3 = 8, PDDM_SierraTwoRow = 9, PDDM_SierraLite = 10, PDDM_Atkinson = 11
#End If

'Points of interest.  These are typically used to associate a current mouse position with a relevant point in the active object
' (e.g. a selection or layer).  Note that the hard-coded constants are negative, by design; complex shapes (like arbitrary
' polygons) may use >= 0 values to identify actual indices in their point array; depending on context, these can also be valid
' interaction point for the user.
Public Enum PD_PointOfInterest

    'An undefined POI just means "the mouse isn't over this object at all"
    poi_Undefined = -1
    
    'The mouse is somewhere in the interior of this object, but not on a corner or edge
    poi_Interior = -2
    
    'Corner POI constants.
    ' Depending on context, this may mean the corner of a complex shape's bounding box (vs the shape itself).
    poi_CornerNW = -3
    poi_CornerNE = -4
    poi_CornerSE = -5
    poi_CornerSW = -6
    
    'Edge POI constants.
    ' Depending on context, this may mean the edge of a complex shape's bounding box (vs the shape itself).
    poi_EdgeN = -7
    poi_EdgeE = -8
    poi_EdgeS = -9
    poi_EdgeW = -10
    
    'Special POI flag that means, "reuse the last POI, whatever it was".  We use this in the main viewport compositor
    ' when a marching ant selection prompts a redraw, and we don't want to lose our last POI.
    poi_ReuseLast = -11
    
End Enum

#If False Then
    Private Const poi_Undefined = -1, poi_Interior = -2, poi_CornerNW = -3, poi_CornerNE = -4, poi_CornerSE = -5, poi_CornerSW = -6, poi_EdgeN = -7, poi_EdgeE = -8, poi_EdgeS = -9, poi_EdgeW = -10
#End If

'When resizing images, PD offers several "fit" strategies.  These allow you to resize an image into a container
' with a different aspect ratio without distorting the image (if desired).
Public Enum PD_ResizeFit
    ResizeFitStretch = 0
    ResizeFitInclusive = 1
    ResizeFitExclusive = 2
End Enum

#If False Then
    Private Const ResizeFitStretch = 0, ResizeFitInclusive = 1, ResizeFitExclusive = 2
#End If

'Internal struct for tracking processor calls.  These are constructed from data passed to the Processor module.
' (NOTE: this struct was finalized in 2013; previous struct versions are no longer supported.)
Public Type PD_ProcessCall
    pcID As String
    pcParameters As String
    pcUndoType As PD_UndoType
    pcTool As PDTools
    pcRaiseDialog As Boolean
    pcRecorded As Boolean
End Type

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)
Public Type PD_UndoEntry
    srcProcCall As PD_ProcessCall   'The processor call that initiated this Undo action
    undoLayerID As Long             'If the undoType is UNDO_LAYER, UNDO_LAYER_VECTORSAFE, or UNDO_LAYERHEADER, this value will note the ID (NOT THE INDEX) of the affected layer
    undoFileSize As Long            'Size of the generated Undo data
    thumbnailLarge As pdDIB         'A large thumbnail associated with the current action.
End Type

'As of 7.0, PD automatically handles navigation keypresses for a variety of controls.
' Want more keys handled?  Add them to this enum.
Public Enum PD_NavigationKey
    pdnk_Enter = vbKeyReturn
    pdnk_Escape = vbKeyEscape
    pdnk_Space = vbKeySpace
    pdnk_Tab = vbKeyTab
End Enum

#If False Then
    Private Const pdnk_Enter = vbKeyReturn, pdnk_Escape = vbKeyEscape, pdnk_Space = vbKeySpace, pdnk_Tab = vbKeyTab
#End If

'Prior to 7.0, only selections offered detailed control over smoothing.  However, new tools (like flood fills)
' also need to describe smoothing features, so this enum is now used in multiple places.
Public Enum PD_EdgeSmoothing
    es_None = 0
    es_Antialiased = 1
    es_FullyFeathered = 2
End Enum

#If False Then
    Private Const es_None = 0, es_Antialiased = 1, es_FullyFeathered = 2
#End If

'This painting struct stores the data passed between BeginPaint and EndPaint
Public Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
End Type

'PNG handling requires special types.
Public Enum PD_PNGColorType
    png_Greyscale = 0
    png_Truecolor = 2
    png_Indexed = 3
    png_GreyscaleAlpha = 4
    png_TruecolorAlpha = 6
    png_AutoColorType = 8   'USED ONLY AT EXPORT TIME, to indicate to the encoder that it is allowed
                            ' to choose the best color type for the source image.
End Enum

#If False Then
    Private Const png_Greyscale = 0, png_Truecolor = 2, png_Indexed = 3, png_GreyscaleAlpha = 4, png_TruecolorAlpha = 6, png_AutoColorType = 8
#End If

Public Type PD_PNGHeader
    Width As Long
    Height As Long
    ColorType As PD_PNGColorType
    Interlaced As Boolean
    BitDepth As Byte
    BitsPerPixel As Byte
End Type

Public Enum PD_Dimension
    pdd_Left
    pdd_Top
    pdd_Width
    pdd_Height
    pdd_AspectRatioW
    pdd_AspectRatioH
    pdd_SwapAspectRatio
    pdd_AspectBoth
End Enum

#If False Then
    Private Const pdd_Left = 0, pdd_Top = 0, pdd_Width = 0, pdd_Height = 0, pdd_AspectRatioW = 0, pdd_AspectRatioH = 0, pdd_SwapAspectRatio = 0, pdd_AspectBoth = 0
#End If
