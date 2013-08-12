Attribute VB_Name = "Public_Variables"

'Contains any and all publicly-declared variables.  I am trying to move
' all public variables here (for obvious reasons), but the transition may
' not be completely done as long as this comment remains!

Option Explicit

'Main user preferences and settings handler
Public g_UserPreferences As pdPreferences

'Main file format compatibility handler
Public g_ImageFormats As pdFormats

'Main language and translation handler
Public g_Language As pdTranslate

'Progress bar class
Public g_ProgBar As cProgressBar

'Color variables
Public g_EmbossEngraveColor As Long 'last used emboss/engrave color

'Currently selected tool, previous tool
Public g_CurrentTool As PDTools
Public g_PreviousTool As PDTools

'Currently supported tools; these numbers correspond to the index of the tool's command button on the main form
Public Enum PDTools
    SELECT_RECT = 0
    SELECT_CIRC = 1
    SELECT_LINE = 2
End Enum

#If False Then
    Const SELECT_RECT = 0
    Const SELECT_CIRC = 1
    Const SELECT_LINE = 2
#End If

'Filter variables
'The array containing the filter data
Public g_FM() As Double
'The size (1x1, 3x3, 5x5, etc) of the filter array
Public g_FilterSize As Long
'The weight (i.e. / by)
Public g_FilterWeight As Double
'The bias (i.e. +/-)
Public g_FilterBias As Double

'Selection variables

'How should the selection be rendered?
Public Enum SelectionRender
    sLightbox = 0
    sHighlightBlue = 1
    sHighlightRed = 2
    sInvertRect = 3
End Enum

Public g_SelectionRenderPreference As SelectionRender


'Zoom data
Public Type ZoomData
    ZoomCount As Byte
    ZoomArray() As Double
    ZoomFactor() As Double
End Type

Public g_Zoom As ZoomData

'Whether or not to resize large images to fit on-screen (0 means "yes," 1 means "no")
Public g_AutosizeLargeImages As Long

'The path where DLLs and related support libraries are kept, currently "ProgramPath\App\PhotoDemon\Plugins\"
Public g_PluginPath As String

'Command line (used here for processing purposes)
Public g_CommandLine As String

'Is scanner/digital camera support enabled?
Public g_ScanEnabled As Boolean

'Is compression via zLib enabled?
Public g_ZLibEnabled As Boolean

'Is metadata handling via ExifTool enabled?
Public g_ExifToolEnabled As Boolean

'Whether or not the user has created a custom filter
Public g_HasCreatedFilter As Boolean

'How to draw the background of image forms; -1 is checkerboard, any other value is treated as an RGB long
Public g_CanvasBackground As Long

'Whether or not to render a drop shadow onto the canvas around the image
Public g_CanvasDropShadow As Boolean

'g_canvasShadow contains a pdShadow object that helps us render a drop shadow around the image, if the user requests it
Public g_CanvasShadow As pdShadow

'Does the user want us to prompt them when they try to close unsaved images?
Public g_ConfirmClosingUnsaved As Boolean

'Whether or not to log program messages in a separate file - this is useful for debugging
Public g_LogProgramMessages As Boolean

'Whether or not we are running in the IDE or compiled
Public g_IsProgramCompiled As Boolean

'Temporary loading variable to disable Autog_Zoom feature
Public g_FixScrolling As Boolean

'For the Open and Save common dialog boxes, it's polite to remember what format the user used last, then default
' the boxes to that.  (Note that these values are stored in the preferences file as well, but that is only accessed
' upon program load and unload.)
Public g_LastOpenFilter As Long
Public g_LastSaveFilter As Long

'Checkerboard mode for rendering transparency.  Possible values are:
' 0 - Light
' 1 - Midtones
' 2 - Dark
' 3 - Custom
Public g_AlphaCheckMode As Long

'Checkerboard colors for rendering transparency
Public g_AlphaCheckOne As Long
Public g_AlphaCheckTwo As Long

'Checkerboard size when rendering transparency.  Possible values are:
' 0 - Small (4x4 pixels per square)
' 1 - Medium (8x8 pixels per square)
' 2 - Large (16x16 pixels per square)
Public g_AlphaCheckSize As Long

'Is the current system running Vista, Windows 7, or later?  (Used to determine availability of certain system fonts)
Public g_IsVistaOrLater As Boolean

'Is theming enabled?  (Used to handle some menu icon rendering quirks)
Public g_IsThemingEnabled As Boolean

'Render the interface using Segoe UI if the user specifies as much in the Preferences dialog
Public g_UseFancyFonts As Boolean
Public g_InterfaceFont As String

'This g_cMonitors object contains data on all monitors on this system.  It is used to handle multiple monitor situations.
Public g_cMonitors As clsMonitors

'If the user attempts to close the program while multiple unsaved images are present, these values allow us to count
' a) how many unsaved images are present
' b) if the user wants to deal with all the images (if the "Repeat this action..." box is checked on the unsaved
'     image confirmation prompt) in the same fashion
' c) what the user's preference is for dealing with all the unsaved images
Public g_NumOfUnsavedImages As Long
Public g_DealWithAllUnsavedImages As Boolean
Public g_HowToDealWithAllUnsavedImages As VbMsgBoxResult

'When the entire program is being shut down, this variable is set
Public g_ProgramShuttingDown As Boolean

'The user is attempting to close all images (necessary for handling the "repeat for all images" check box)
Public g_ClosingAllImages As Boolean

'JPEG export options; these are set by the JPEG export dialog if the user clicks "OK" (not Cancel)
Public g_JPEGQuality As Long
Public g_JPEGFlags As Long
Public g_JPEGThumbnail As Long

'JPEG-2000 export compression ratio; this is set by the JP2 export dialog if the user clicks "OK" (not Cancel)
Public g_JP2Compression As Long

'Exported color depth
Public g_ColorDepth As Long

'Color count
Public g_LastColorCount As Long

'Is the current image grayscale?  This variable is set by the quick count colors routine.  Do not trust its
' state unless you have just called the quick count colors routine (otherwise it may be outdated).
Public g_IsImageGray As Boolean

'Is the current image black and white (literally, is it monochrome e.g. comprised of JUST black and JUST white)?
' This variable is set by the quick count colors routine.  Do not trust its state unless you have just called
' the quick count colors routine (otherwise it may be outdated).
Public g_IsImageMonochrome As Boolean

'What threshold should be used for simplifying an image's complex alpha channel?
' (This is set by the custom alpha cutoff dialog.)
Public g_AlphaCutoff As Byte

'When an image has its colors counted, the image's ID is stored here.  Other functions can use this to see if the
' current color count is relevant for a given image (e.g. if the image being worked on has just had its colors counted).
Public g_LastImageScanned As Long

'Some actions take a long time to execute.  This global variable can be used to track if a function is still running.
' Just make sure to initialize it properly (in case the last function didn't!).
'Public g_Processing As Boolean

'If this is the first time the user has run PhotoDemon (as determined by the lack of a preferences XML file), this
' variable will be set to TRUE early in the load process.  Other routines can then modify their behavior accordingly.
Public g_IsFirstRun As Boolean

'Drag and drop operations are allowed at certain times, but not others.  Any time a modal form is displayed, drag-and-drop
' must be disallowed - with the exception of common dialog boxes.  To make sure this behavior is carefully maintained,
' we track drag-and-drop enabling ourselves
Public g_AllowDragAndDrop As Boolean

'While Undo/Redo operations are active, certain tasks can be ignored.  This public value can be used to check Undo/Redo activity.
Public g_UndoRedoActive As Boolean
