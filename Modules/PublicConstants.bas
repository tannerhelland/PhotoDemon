Attribute VB_Name = "Public_Constants"
Option Explicit

Public Enum BUILD_QUALITY
    PD_ALPHA = 1
    PD_BETA = 2
    PD_PRODUCTION = 3
End Enum

#If False Then
    Const PD_ALPHA = 1, PD_BETA = 2, PD_PRODUCTION = 3
#End If

'Quality of the current build.  This value automatically dictates a number of behaviors throughout the program,
' like reporting time-to-completion for effects and enabling detailed debug reports.  Do not change unless you
' fully understand the consequences!
Public Const PD_BUILD_QUALITY As Long = PD_ALPHA

'Identifier for various PD-specific file types
Public Const PD_IMAGE_IDENTIFIER As Long = &H44494450   'pdImage data (ASCII characters "PDID", as hex, little-endian)
Public Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

'Magic number for errors that arise during pdPackage interactions
Public Const PDP_GENERIC_ERROR As Long = 9001

'Some constants used for general program changes (better to leave them as constants here, then to
' have to manually change them when I think up better or more appropriate ones)
Public Const MACRO_EXT As String * 3 = "pdm"

'Maximum allowed image dimensions (width, height); used at run-time to set max values for things like Image > Resize
Public Const PD_MAX_IMAGE_DIMENSION As Long = 100000

'Standard API constants
Public Const MAX_PATH_LEN As Long = 260

'When a UC with an image is hovered, we typically reflect this via some kind of "glow" state.  This constant controls
' the amount of brightness added to the image during a hover state.
Public Const UC_HOVER_BRIGHTNESS As Long = 50

'Virtual key constants
Public Const VK_SHIFT As Long = &H10
Public Const VK_CONTROL As Long = &H11
Public Const VK_ALT As Long = &H12

Public Const VK_LEFT As Long = &H25
Public Const VK_UP As Long = &H26
Public Const VK_RIGHT As Long = &H27
Public Const VK_DOWN As Long = &H28

'Numeric consts are not currently used
'Public Const VK_0 As Long = &H30
'Public Const VK_1 As Long = &H31
'Public Const VK_2 As Long = &H32
'Public Const VK_3 As Long = &H33
'Public Const VK_4 As Long = &H34
'Public Const VK_5 As Long = &H35
'Public Const VK_6 As Long = &H36
'Public Const VK_7 As Long = &H37
'Public Const VK_8 As Long = &H38
'Public Const VK_9 As Long = &H39
'Public Const VK_NUMLOCK As Long = &H90
'Public Const VK_NUMPAD0 As Long = &H60
'Public Const VK_NUMPAD1 As Long = &H61
'Public Const VK_NUMPAD2 As Long = &H62
'Public Const VK_NUMPAD3 As Long = &H63
'Public Const VK_NUMPAD4 As Long = &H64
'Public Const VK_NUMPAD5 As Long = &H65
'Public Const VK_NUMPAD6 As Long = &H66
'Public Const VK_NUMPAD7 As Long = &H67
'Public Const VK_NUMPAD8 As Long = &H68
'Public Const VK_NUMPAD9 As Long = &H69

Public Const VK_BACK As Long = &H8
Public Const VK_TAB As Long = &H9
Public Const VK_RETURN As Long = &HD
Public Const VK_CAPITAL As Long = &H14
Public Const VK_SPACE As Long = &H20
Public Const VK_INSERT As Long = &H2D
Public Const VK_DELETE As Long = &H2E
Public Const VK_ESCAPE As Long = &H1B
Public Const VK_PAGEUP As Long = &H21
Public Const VK_PAGEDOWN As Long = &H22
Public Const VK_END As Long = &H23
Public Const VK_HOME As Long = &H24
Public Const VK_MULTIPLY As Long = &H6A     'Multiply key (numpad)
Public Const VK_ADD As Long = &H6B          'Add key (numpad)
Public Const VK_SUBTRACT As Long = &H6D     'Subtract key (numpad)

Public Const VK_OEM_COMMA As Long = 188     'For any country/region, the ',' key
Public Const VK_OEM_PERIOD As Long = &HBE   'For any country/region, the '.' key
Public Const VK_OEM_PLUS As Long = 187      'Locale-inspecific + key
Public Const VK_OEM_MINUS As Long = 189     'Locale-inspecific - key
Public Const VK_OEM_1 As Long = &HBA        'For the US standard keyboard, the ';:' key.  (Varies internationally.)
Public Const VK_OEM_4 As Long = 219         'For the US standard keyboard, the '[{' key.  (Varies internationally.)
Public Const VK_OEM_6 As Long = 221         'For the US standard keyboard, the ']}' key.  (Varies internationally.)
Public Const VK_OEM_7 As Long = &HDE        'For the US standard keyboard, the 'single-quote/double-quote' key.  (Varies internationally.)

'Old PDI files were not Unicode friendly.  When loading PDI files, we use this constant to determine whether
' ANSI or Unicode string behavior should be used.
Public Const PDPACKAGE_UNICODE_FRIENDLY_VERSION As Long = 66

'PD uses some of its own window messages to simplify things like cross-control notifications.
Public Const WM_APP As Long = &H8000&
Public Const WM_PD_PRIMARY_COLOR_CHANGE As Long = (WM_APP + 16&)
Public Const WM_PD_COLOR_MANAGEMENT_CHANGE As Long = (WM_APP + 17&)
Public Const WM_PD_DIALOG_NAVKEY As Long = (WM_APP + 18&)
Public Const WM_PD_PRIMARY_COLOR_APPLIED As Long = (WM_APP + 19&)
Public Const WM_PD_FOCUS_FROM_TAB_KEY As Long = (WM_APP + 20&)
Public Const WM_PD_TAB_KEY_TARGET As Long = (WM_APP + 21&)
Public Const WM_PD_SHIFT_TAB_KEY_TARGET As Long = (WM_APP + 22&)
Public Const WM_PD_FLASH_ACTIVE_LAYER As Long = (WM_APP + 23&)
Public Const WM_PD_DIALOG_RESIZE_FINISHED As Long = (WM_APP + 24&)
Public Const WM_PD_HIDECHILD As Long = (WM_APP + 25&)

'Inside the IDE, we can't rely on PD's central themer for color decisions (as it won't be initialized).
' A few constants are used instead.
Public Const IDE_WHITE As String = "#ffffff"
Public Const IDE_BLUE As String = "#3296dc"
Public Const IDE_GRAY As String = "#404040"
Public Const IDE_BLACK As String = "#000000"

'Default alpha cut-off when "auto" is selected
Public Const PD_DEFAULT_ALPHA_CUTOFF As Long = 64

'When applying localizations, individual controls can pass a unique name/ID string as the last parameter to
' TranslateMessage by using this prefix.  The translation engine can then match the object name against
' any special per-object (not per-caption) translations in the active language file.
Public Const SPECIAL_TRANSLATION_OBJECT_PREFIX As String = "obj-"
Public Const CONTROL_ARRAY_IDX_SEPARATOR As String = "."

'UI element sizes are standardized against 96-DPI.  It's up to the caller to adjust these at run-time as relevant.
Public Const SQUARE_CORNER_SIZE As Single = 12!
