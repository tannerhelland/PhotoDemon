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
Public Const EXPECTED_PNGQUANT_VERSION As String = "2.1.1"
Public Const EXPECTED_EXIFTOOL_VERSION As String = "9.70"

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

