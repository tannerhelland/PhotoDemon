Attribute VB_Name = "Public_Constants"
Option Explicit

'Enable this constant if you want PhotoDemon to report time-to-completion for filters and effects
Public Const DISPLAY_TIMINGS As Boolean = True

'Enable this constant if you want PhotoDemon to use experimental methods (when available).  This is helpful
' during debugging, but SHOULD NEVER BE ENABLED IN PRODUCTION BUILDS!
Public Const PD_EXPERIMENTAL_MODE As Boolean = False

'Identifier for pdImage data saved to file.  (ASCII characters "PDID", as hex, listed here in little-endian notation.)
Public Const PD_IMAGE_IDENTIFIER As Long = &H44494450

'Identifier for pdLayer data saved to file.  (ASCII characters "PDIL", as hex, listed here in little-endian notation.)
Public Const PD_LAYER_IDENTIFIER As Long = &H4C494450

'Magic number for errors that arise during pdPackage interactions
Public Const PDP_GENERIC_ERROR As Long = 9001

'Expected version numbers of plugins.  These are updated at each new PhotoDemon release (if a new version of
' the plugin is available, obviously).
Public Const EXPECTED_FREEIMAGE_VERSION As String = "3.16.1"
Public Const EXPECTED_ZLIB_VERSION As String = "1.2.8"
Public Const EXPECTED_EZTWAIN_VERSION As String = "1.18.0"
Public Const EXPECTED_PNGNQ_VERSION As String = "2.0.1"
Public Const EXPECTED_EXIFTOOL_VERSION As String = "9.62"

'Some constants used for general program changes (better to leave them as constants here, then to
' have to manually change them when I think up better or more appropriate ones)
Public Const PROGRAMNAME As String = "PhotoDemon"
Public Const FILTER_EXT As String * 3 = "pde"
Public Const MACRO_EXT As String * 3 = "pdm"
Public Const SELECTION_EXT As String * 3 = "pds"

'Constants used for passing image resize options.
' Note that options 3-6 require use of the FreeImage library
Public Const RESIZE_NORMAL As Long = 0
Public Const RESIZE_HALFTONE As Long = 1
Public Const RESIZE_BILINEAR As Long = 2
Public Const RESIZE_BSPLINE As Long = 3
Public Const RESIZE_BICUBIC_MITCHELL As Long = 4
Public Const RESIZE_BICUBIC_CATMULL As Long = 5
Public Const RESIZE_LANCZOS As Long = 6

'Constants used for reducing image colors
Public Const REDUCECOLORS_AUTO As Long = 0
Public Const REDUCECOLORS_MANUAL As Long = 1
Public Const REDUCECOLORS_MANUAL_ERRORDIFFUSION As Long = 2

'Constants for the drop shadow drawn around the image on the image canvas.  At some point these may become user-editable.
Public Const PD_CANVASSHADOWSIZE As Long = 5
Public Const PD_CANVASSHADOWSTRENGTH As Long = 70

'Constant for testing JP2/J2K support.  These may or may not become permanent pending the outcome of some rigorous testing.
Public Const JP2_ENABLED As Boolean = True

'Mathematic constants
Public Const PI As Double = 3.14159265358979
Public Const PI_HALF As Double = 1.5707963267949
Public Const PI_DOUBLE As Double = 6.28318530717958
Public Const PI_DIV_180 As Double = 0.017453292519943
Public Const EULER As Double = 2.71828182845905

'Data constants
Public Const LONG_MAX As Long = 2147483647

'Edge-handling methods for distort-style filters
Public Enum EDGE_OPERATOR
    EDGE_CLAMP = 0
    EDGE_REFLECT = 1
    EDGE_WRAP = 2
    EDGE_ERASE = 3
    EDGE_ORIGINAL = 4
End Enum
#If False Then
    Const EDGE_CLAMP = 0
    Const EDGE_REFLECT = 1
    Const EDGE_WRAP = 2
    Const EDGE_ERASE = 3
    Const EDGE_ORIGINAL = 4
#End If

'Maximum width (in pixels) for custom-built tooltips
Public Const PD_MAX_TOOLTIP_WIDTH As Long = 400

'Constants used for pulling up an API browse-for-folder box
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BFFM_INITIALIZED = 1
Public Const MAX_PATH_LEN = 260

'Orientation (used in a whole bunch of different situations)
Public Enum PD_ORIENTATION
    PD_HORIZONTAL = 0
    PD_VERTICAL = 1
    PD_BOTH = 2
End Enum

#If False Then
    Const PD_HORIZONTAL = 0, PD_VERTICAL = 1, PD_BOTH = 2
#End If

'Some PhotoDemon actions can operate on the whole image, or on just a specific layer (e.g. resize).  When initiating
' one of these dual-action operations, the constants below can be used to specify the mode.
Public Enum PD_ACTION_TARGET
    PD_AT_WHOLEIMAGE = 0
    PD_AT_SINGLELAYER = 1
End Enum

#If False Then
    Const PD_AT_WHOLEIMAGE = 0, PD_AT_SINGLELAYER = 1
#End If

'When an action triggers the creation of Undo/Redo data, it must specify what kind of Undo/Redo data it wants created.
' This type is used by PD to determine the most efficient way to store/restore previous actions.
Public Enum PD_UNDO_TYPE
    UNDO_NOTHING = -1
    UNDO_EVERYTHING = 0
    UNDO_IMAGE = 1
    UNDO_IMAGEHEADER = 2
    UNDO_LAYER = 3
    UNDO_LAYERHEADER = 4
    UNDO_SELECTION = 5
End Enum

#If False Then
    Const UNDO_NOTHING = -1, UNDO_EVERYTHING = 0, UNDO_IMAGE = 1, UNDO_IMAGEHEADER = 2, UNDO_LAYER = 3, UNDO_LAYERHEADER = 4, UNDO_SELECTION = 5
#End If

Public Type RGBQUAD
   Blue As Byte
   Green As Byte
   Red As Byte
   Alpha As Byte
End Type

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
Public Enum PD_EDGE_DETECTION
    PD_EDGE_ARTISTIC_CONTOUR = 0
    PD_EDGE_HILITE = 1
    PD_EDGE_LAPLACIAN = 2
    PD_EDGE_PHOTODEMON = 3
    PD_EDGE_PREWITT = 4
    PD_EDGE_ROBERTS = 5
    PD_EDGE_SOBEL = 6
End Enum

#If False Then
    Private Const PD_EDGE_ARTISTIC_CONTOUR = 0, PD_EDGE_HILITE = 1, PD_EDGE_LAPLACIAN = 2, PD_EDGE_PHOTODEMON = 3, PD_EDGE_PREWITT = 4, PD_EDGE_ROBERTS = 5, PD_EDGE_SOBEL = 6
#End If

Public Enum PD_EDGE_DETECTION_DIRECTION
    PD_EDGE_DIR_ALL = 0
    PD_EDGE_DIR_HORIZONTAL = 1
    PD_EDGE_DIR_VERTICAL = 2
End Enum

#If False Then
    Private Const PD_EDGE_DIR_ALL = 0, PD_EDGE_DIR_HORIZONTAL = 1, PD_EDGE_DIR_VERTICAL = 2
#End If
