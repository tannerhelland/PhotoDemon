Attribute VB_Name = "Public_Enums_and_Types"
Option Explicit

Public Type RGBQUAD
   Blue As Byte
   Green As Byte
   Red As Byte
   Alpha As Byte
End Type

'Currently supported tools; these numbers correspond to the index of the tool's command button on the main form.
' In theory, adding new tools should be as easy as changing these numbers.  All tool-related code is tied into these
' constants, so any changes here should automatically propagate throughout the software.  (In practice, be sure to
' double-check everything!!)
Public Enum PDTools
    NAV_DRAG = 0
    NAV_MOVE = 1
    SELECT_RECT = 2
    SELECT_CIRC = 3
    SELECT_LINE = 4
    SELECT_LASSO = 5
    QUICK_FIX_LIGHTING = 6
End Enum

#If False Then
    Const NAV_DRAG = 0, NAV_MOVE = 1, SELECT_RECT = 2, SELECT_CIRC = 3, SELECT_LINE = 4, SELECT_LASSO = 5, QUICK_FIX_LIGHTING = 6
#End If

'How should the selection be rendered?
Public Enum SelectionRender
    SELECTION_RENDER_LIGHTBOX = 0
    SELECTION_RENDER_HIGHLIGHT = 1
End Enum

#If False Then
    Const SELECTION_RENDER_LIGHTBOX = 0, SELECTION_RENDER_HIGHLIGHT = 1
#End If

'JPEG automatic quality detection modes
Public Enum jpegAutoQualityMode
    doNotUseAutoQuality = 0
    noDifference = 1
    tinyDifference = 2
    minorDifference = 3
    majorDifference = 4
End Enum

#If False Then
    Private Const doNotUseAutoQuality = 0, noDifference = 1, tinyDifference = 2, minorDifference = 3, majorDifference = 4
#End If

'PhotoDemon's language files provide a small amount of metadata to help the program know how to use them.  This type
' was previously declared inside the pdTranslate class, but with the addition of a Language Editor, I have moved it
' here, so the entire project can access the type.
Public Type pdLanguageFile
    Author As String
    FileName As String
    langID As String
    langName As String
    langType As String
    langVersion As String
    langStatus As String
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
    Const pdSE_AnySave = 0, pdSE_SavePDI = 1, pdSE_SaveFlat = 2
#End If

'Edge-handling methods for distort-style filters
Public Enum EDGE_OPERATOR
    EDGE_CLAMP = 0
    EDGE_REFLECT = 1
    EDGE_WRAP = 2
    EDGE_ERASE = 3
    EDGE_ORIGINAL = 4
End Enum

#If False Then
    Const EDGE_CLAMP = 0, EDGE_REFLECT = 1, EDGE_WRAP = 2, EDGE_ERASE = 3, EDGE_ORIGINAL = 4
#End If

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

'PhotoDemon performance settings are generally provided in three groups: Max Quality, Balanced, and Max Performance
Public Enum PD_PERFORMANCE_SETTING
    PD_PERF_BESTQUALITY = 0
    PD_PERF_BALANCED = 1
    PD_PERF_FASTEST = 2
End Enum

#If False Then
    Private Const PD_PERF_BESTQUALITY = 0, PD_PERF_BALANCED = 1, PD_PERF_FASTEST = 2
#End If

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)
Public Type undoEntry
    processID As String             'Name of the associated action (e.g. "Gaussian blur")
    processParamString As String    'Processor string supplied to the action
    undoType As PD_UNDO_TYPE        'What type of Undo/Redo data was stored for this action (e.g. Image or Selection data)
    undoLayerID As Long             'If the undoType is UNDO_LAYER or UNDO_LAYERHEADER, this value will note the ID (NOT THE INDEX) of the affected layer
    relevantTool As Long            'If a tool was associated with this action, it can be set here.  This value is not currently used.
    thumbnailSmall As pdDIB         'A small thumbnail associated with the current action.  In the future, this will be used by the Undo History window.
    thumbnailLarge As pdDIB         'A large thumbnail associated with the current action.
End Type

'PhotoDemon supports multiple image encoders and decoders.
Public Enum PD_IMAGE_DECODER_ENGINE
    PDIDE_INTERNAL = 0
    PDIDE_FREEIMAGE = 1
    PDIDE_GDIPLUS = 2
    PDIDE_VBLOADPICTURE = 3
End Enum

#If False Then
    Private Const PDIDE_INTERNAL = 0, PDIDE_FREEIMAGE = 1, PDIDE_GDIPLUS = 2, PDIDE_VBLOADPICTURE = 3
#End If

'Some UI DIBs are generated at run-time.  These DIBs can be requested by using the getRuntimeUIDIB() function.
Public Enum PD_RUNTIME_UI_DIB
    PDRUID_CHANNEL_RED = 0
    PDRUID_CHANNEL_GREEN = 1
    PDRUID_CHANNEL_BLUE = 2
    PDRUID_CHANNEL_RGB = 3
End Enum

#If False Then
    Private Const PDRUID_CHANNEL_RED = 0, PDRUID_CHANNEL_GREEN = 1, PDRUID_CHANNEL_BLUE = 2, PDRUID_CHANNEL_RGB = 3
#End If
