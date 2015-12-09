Attribute VB_Name = "Public_Constants"
Option Explicit

Public Enum BUILD_QUALITY
    PD_PRE_ALPHA = 0
    PD_ALPHA = 1
    PD_BETA = 2
    PD_PRODUCTION = 3
End Enum

#If False Then
    Const PD_PRE_ALPHA = 0, PD_ALPHA = 1, PD_BETA = 2, PD_PRODUCTION = 3
#End If

'Quality of the current build.  This value automatically dictates a number of behaviors throughout the program,
' like reporting time-to-completion for effects and enabling detailed debug reports.  Do not change unless you
' fully understand the consequences!
'
'IMPORTANT NOTE!  In conjunction with this constant, a compile-time constant called "DEBUGMODE" should be set to 1
' for any non-production (release) builds.  This compile-time constant enables various conditional compilation
' segments through the program, including the bulk of PD's debugging code.
'
' Obvious corollary: ALWAYS SET DEBUGMODE TO 0 IN PRODUCTION BUILDS!
Public Const PD_BUILD_QUALITY As Long = PD_PRE_ALPHA

'Identifier for various PD-specific file types
Public Const PD_IMAGE_IDENTIFIER As Long = &H44494450   'pdImage data (ASCII characters "PDID", as hex, little-endian)
Public Const PD_LAYER_IDENTIFIER As Long = &H4C494450   'pdLayer data (ASCII characters "PDIL", as hex, little-endian)
Public Const PD_LANG_IDENTIFIER As Long = &H414C4450    'pdLanguage data (ASCII characters "PDLA", as hex, little-endian)
Public Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

'Magic number for errors that arise during pdPackage interactions
Public Const PDP_GENERIC_ERROR As Long = 9001

'Expected version numbers of plugins.  These are updated at each new PhotoDemon release (if a new version of
' the plugin is available, obviously).
Public Const EXPECTED_FREEIMAGE_VERSION As String = "3.18.0"
Public Const EXPECTED_ZLIB_VERSION As String = "1.2.8"
Public Const EXPECTED_EZTWAIN_VERSION As String = "1.18.0"
Public Const EXPECTED_PNGQUANT_VERSION As String = "2.3.1"
Public Const EXPECTED_EXIFTOOL_VERSION As String = "10.01"

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
Public Const DOUBLE_MAX As Double = 1.79769313486231E+308

'Maximum width (in pixels) for custom-built tooltips
Public Const PD_MAX_TOOLTIP_WIDTH As Long = 400

'Constants used for pulling up an API browse-for-folder box
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BFFM_INITIALIZED = 1
Public Const MAX_PATH_LEN = 260

'PhotoDemon's internal PDI format identifier.  We preface this with FIF_ because PhotoDemon uses FreeImage's format constants
' to track save state.  However, FreeImage does not include PDI support, and because their VB6 interface may change between
' versions, we don't want to store our constant in the FreeImage modules - so we keep it here!
Public Const FIF_PDI As Long = 100

'Some other FIF_ formats supported by PhotoDemon, but not by FreeImage
Public Const FIF_WMF As Long = 110
Public Const FIF_EMF As Long = 111

'When a UC with an image is hovered, we typically reflect this via some kind of "glow" state.  This constant controls
' the amount of brightness added to the image during a hover state.
Public Const UC_HOVER_BRIGHTNESS As Long = 50

'Virtual key constants
Public Const VK_LEFT As Long = &H25
Public Const VK_UP As Long = &H26
Public Const VK_RIGHT As Long = &H27
Public Const VK_DOWN As Long = &H28

Public Const VK_NUMLOCK As Long = &H90
Public Const VK_NUMPAD0 As Long = &H60
Public Const VK_NUMPAD1 As Long = &H61
Public Const VK_NUMPAD2 As Long = &H62
Public Const VK_NUMPAD3 As Long = &H63
Public Const VK_NUMPAD4 As Long = &H64
Public Const VK_NUMPAD5 As Long = &H65
Public Const VK_NUMPAD6 As Long = &H66
Public Const VK_NUMPAD7 As Long = &H67
Public Const VK_NUMPAD8 As Long = &H68
Public Const VK_NUMPAD9 As Long = &H69

Public Const VK_BACK As Long = &H8
Public Const VK_TAB As Long = &H9
Public Const VK_RETURN As Long = &HD
Public Const VK_SPACE As Long = &H20
Public Const VK_INSERT As Long = &H2D
Public Const VK_DELETE As Long = &H2E
Public Const VK_ESCAPE As Long = &H1B
Public Const VK_PAGEUP As Long = &H21
Public Const VK_PAGEDOWN As Long = &H22
Public Const VK_END As Long = &H23
Public Const VK_HOME As Long = &H24

Public Const VK_0 As Long = &H30
Public Const VK_1 As Long = &H31
Public Const VK_2 As Long = &H32
Public Const VK_3 As Long = &H33
Public Const VK_4 As Long = &H34
Public Const VK_5 As Long = &H35
Public Const VK_6 As Long = &H36
Public Const VK_7 As Long = &H37
Public Const VK_8 As Long = &H38
Public Const VK_9 As Long = &H39

'Old PDI files were not Unicode friendly.  When loading PDI files, we use this constant to determine whether
' ANSI or Unicode string behavior should be used.
Public Const PDPACKAGE_UNICODE_FRIENDLY_VERSION As Long = 66

'PD uses some of its own window messages to simplify things like notifications.
Public Const WM_APP As Long = &H8000&
Public Const WM_PD_PRIMARY_COLOR_CHANGE As Long = (WM_APP + 16&)
