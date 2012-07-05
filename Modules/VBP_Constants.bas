Attribute VB_Name = "Public_Constants"

Option Explicit

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
'Constants used for passing image resize options
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

