Attribute VB_Name = "Public_Constants"
Option Explicit

'Enable this constant if you want PhotoDemon to report time-to-completion for filters and effects
Public Const DISPLAY_TIMINGS As Boolean = True

'Constants related to specialized mouse handling
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_MOUSEFORWARDBACK As Long = 793
Public Const WM_MOUSEKEYBACK As Long = -2147418112
Public Const WM_MOUSEKEYFORWARD As Long = -2147352576
Public Const WM_MOUSELEAVE As Long = &H2A3
Public Const TME_LEAVE As Long = &H2
Public Const TME_CANCEL As Long = &H80000000

'Constants used for pulling up an API browse-for-folder box
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BFFM_INITIALIZED = 1
Public Const MAX_PATH_LEN = 260

'Some constants used for general program changes (better to leave them as constants here, then to
' have to manually change them when I think up better or more appropriate ones)
Public Const PROGRAMNAME As String = "PhotoDemon"
Public Const FILTER_EXT As String * 3 = "thf"
Public Const MACRO_EXT As String * 3 = "thm"

'Constants used for passing image resize options
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
Public Const EULER As Double = 2.71828182845905

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
