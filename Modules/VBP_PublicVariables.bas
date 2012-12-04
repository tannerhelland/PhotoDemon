Attribute VB_Name = "Public_Variables"

'Contains any and all publicly-declared variables.  I am trying to move
' all public variables to here for obvious reasons, but the transition may
' not be completely done as long as this comment remains!

Option Explicit

'Main user preferences and settings handler
Public userPreferences As pdPreferences

'Main file format compatibility handler
Public imageFormats As pdFormats

'Progress bar class
Public cProgBar As cProgressBar

'Color variables
Public EmbossEngraveColor As Long 'last used emboss/engrave color

'Filter variables
'The array containing the filter data
Public FM() As Long
'The size (1x1, 3x3, 5x5, etc) of the filter array
Public FilterSize As Byte
'The weight (i.e. / by)
Public FilterWeight As Long
'The bias (i.e. +/-)
Public FilterBias As Long

'Selection variables

'How should the selection be rendered?
Public Enum SelectionRender
    sLightbox = 0
    sHighlightBlue = 1
    sHighlightRed = 2
    sInvertRect = 3
End Enum

Public selectionRenderPreference As SelectionRender


'Zoom data
Public Type ZoomData
    ZoomCount As Byte
    ZoomArray() As Double
    ZoomFactor() As Double
End Type

Public Zoom As ZoomData

'Whether or not to resize large images to fit on-screen (preference is stored in the INI file; 0 means "yes," 1 means "no")
Public AutosizeLargeImages As Long

'The path where DLLs and related support libraries are kept, currently "ProgramPath\Data\Plugins\"
Public PluginPath As String

'Command line (used here for processing purposes)
Public CommandLine As String

'Commonly used loop variables
Public x As Long
Public y As Long
Public z As Long

'Is scanner/digital camera support enabled?
Public ScanEnabled As Boolean

'Is compression via zLib enabled?
Public zLibEnabled As Boolean

'Whether or not the user has created a custom filter
Public HasCreatedFilter As Boolean

'How to draw the background of image forms; -1 is checkerboard, any other value is treated as an RGB long
Public CanvasBackground As Long

'Whether or not to render a drop shadow onto the canvas around the image
Public CanvasDropShadow As Boolean

'canvasShadow contains a pdShadow object that helps us render a drop shadow around the image, if the user requests it
Public canvasShadow As pdShadow

'Does the user want us to prompt them when they try to close unsaved images?
Public ConfirmClosingUnsaved As Boolean

'Whether or not to log program messages in a separate file - this is useful for debugging
Public LogProgramMessages As Boolean

'Whether or not we are running in the IDE or compiled
Public isProgramCompiled As Boolean

'Temporary loading variable to disable Autozoom feature
Public FixScrolling As Boolean

'For the Open and Save common dialog boxes, it's polite to remember what format the user used last, then default
' the boxes to that.  (Note that these values are stored in the INI file as well, but that is only accessed
' upon program load and unload.)
Public LastOpenFilter As Long
Public LastSaveFilter As Long

'Was the save dialog canceled?
Public saveDialogCanceled As Boolean

'Checkerboard mode for rendering transparency.  Possible values are:
' 0 - Light
' 1 - Midtones
' 2 - Dark
' 3 - Custom
Public AlphaCheckMode As Long

'Checkerboard colors for rendering transparency
Public AlphaCheckOne As Long
Public AlphaCheckTwo As Long

'Checkerboard size when rendering transparency.  Possible values are:
' 0 - Small (4x4 pixels per square)
' 1 - Medium (8x8 pixels per square)
' 2 - Large (16x16 pixels per square)
Public AlphaCheckSize As Long

'Is the current system running Vista, Windows 7, or later?  (Used to determine availability of certain system fonts)
Public isVistaOrLater As Boolean

'Is theming enabled?  (Used to handle some menu icon rendering quirks)
Public isThemingEnabled As Boolean

'Render the interface using Segoe UI if the user specifies as much in the Preferences dialog
Public useFancyFonts As Boolean

'This cMonitors object contains data on all monitors on this system.  It is used to handle multiple monitor situations.
Public cMonitors As clsMonitors

'If the user attempts to close the program while multiple unsaved images are present, these values allow us to count
' a) how many unsaved images are present
' b) if the user wants to deal with all the images (if the "Repeat this action..." box is checked on the unsaved
'     image confirmation prompt) in the same fashion
' c) what the user's preference is for dealing with all the unsaved images
Public numOfUnsavedImages As Long
Public dealWithAllUnsavedImages As Boolean
Public howToDealWithAllUnsavedImages As VbMsgBoxResult

'When the entire program is being shut down, this variable is set
Public programShuttingDown As Boolean

'The user is attempting to close all images (necessary for handling the "repeat for all images" check box)
Public closingAllImages As Boolean

'JPEG export options; these are set by the JPEG export dialog if the user clicks "OK" (not Cancel)
Public g_JPEGQuality As Long
Public g_JPEGFlags As Long
Public g_JPEGThumbnail As Long
