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
Public Const EULER As Double = 2.71828182845905

'Data constants
Public Const LONG_MAX As Long = 2147483647
Public Const DOUBLE_MAX As Double = 1.79769313486231E+308

'Maximum width (in pixels) for custom-built tooltips
Public Const PD_MAX_TOOLTIP_WIDTH As Long = 400

'Standard API constants
Public Const MAX_PATH_LEN = 260

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

Public Const VK_OEM_PLUS As Long = 187   'Locale-inspecific + key
Public Const VK_OEM_MINUS As Long = 189  'Locale-inspecific - key
Public Const VK_OEM_4 As Long = 219      'For the US standard keyboard, the '[{' key.  (Varies internationally.)
Public Const VK_OEM_6 As Long = 221      'For the US standard keyboard, the ']}' key.  (Varies internationally.)

'Old PDI files were not Unicode friendly.  When loading PDI files, we use this constant to determine whether
' ANSI or Unicode string behavior should be used.
Public Const PDPACKAGE_UNICODE_FRIENDLY_VERSION As Long = 66

'PD uses some of its own window messages to simplify things like cross-control notifications.
Public Const WM_APP As Long = &H8000&
Public Const WM_PD_PRIMARY_COLOR_CHANGE As Long = (WM_APP + 16&)
Public Const WM_PD_COLOR_MANAGEMENT_CHANGE As Long = (WM_APP + 17&)

'Inside the IDE, we can't rely on PD's central themer for color decisions (as it won't be initialized).
' A few constants are used instead.
Public Const IDE_WHITE As String = "#ffffff"
Public Const IDE_BLUE As String = "#3296dc"
Public Const IDE_LIGHTBLUE As String = "#3cafe6"
Public Const IDE_GRAY As String = "#404040"
Public Const IDE_BLACK As String = "#000000"
Public Const IDE_RED As String = "#0000ff"
