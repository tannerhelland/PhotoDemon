Attribute VB_Name = "Public_Variables"

'Contains any and all publicly-declared variables.  I am trying to move
' all public variables here (for obvious reasons), but the transition may
' not be completely done as long as this comment remains!

Option Explicit

'Lightweight ThunderMain listener
Public g_ThunderMain As pdThunderMain

'Main resource handler
Public g_Resources As pdResources

'Main language and translation handler
Public g_Language As pdTranslate

'Main clipboard handler
Public g_Clipboard As pdClipboardMain

'Currently selected tool, previous tool
Public g_CurrentTool As PDTools
Public g_PreviousTool As PDTools

'Does the user want us to prompt them when they try to close unsaved images?
Public g_ConfirmClosingUnsaved As Boolean

'For the Open and Save common dialog boxes, it's polite to remember what format the user used last, then default
' the boxes to that.  (Note that these values are stored in the preferences file as well, but that is only accessed
' upon program load and unload.)
Public Const PD_USER_PREF_UNKNOWN As Long = -1
Public g_LastOpenFilter As Long, g_LastSaveFilter As Long

'DIB that contains a 2x2 pattern of the alpha checkerboard.  Use it with CreatePatternBrush to paint the alpha
' checkerboard prior to rendering.
Public g_CheckerboardPattern As pdDIB

'Copy of g_CheckerboardPattern, above, but in pd2DBrush format.  The brush is pre-built as a GDI+ texture brush,
' which makes it preferable for painting on 32-bpp surfaces.
Public g_CheckerboardBrush As pd2DBrush

'This g_Displays object contains data on all display devices on this system.  It includes a ton of code to assist the program
' with managing multiple monitors and other display-related issues.
Public g_Displays As pdDisplays

'If the user attempts to close the program while multiple unsaved images are present, these values tell us...
' 1) if the user wants to deal with all unsaved images (e.g. if the "Repeat this action..." box was checked
'     on the "unsaved image confirmation" dialog)
' 2) what the user wants us to do with the various remaining unsaved images (e.g. save vs discard)
Public g_DealWithAllUnsavedImages As Boolean
Public g_HowToDealWithAllUnsavedImages As VbMsgBoxResult

'When the entire program is being shut down, this variable is set
Public g_ProgramShuttingDown As Boolean

'The user is attempting to close all images (necessary for handling the "repeat for all images" check box)
Public g_ClosingAllImages As Boolean

'If this is the first time the user has run PhotoDemon (as determined by the lack of a preferences XML file), this
' variable will be set to TRUE early in the load process.  Other routines can then modify their behavior accordingly.
Public g_IsFirstRun As Boolean

'Drag and drop operations are allowed at certain times, but not others.  Any time a modal form is displayed, drag-and-drop
' must be disallowed - with the exception of common dialog boxes.  To make sure this behavior is carefully maintained,
' we track drag-and-drop enabling ourselves
Public g_AllowDragAndDrop As Boolean

'This window manager handles positioning, layering, and sizing of the main canvas and all toolbars
Public g_WindowManager As pdWindowManager

'UI theme engine.
Public g_Themer As pdTheme

'"File > Open Recent" and "Tools > Recent Macros" dynamic menu managers
Public g_RecentFiles As pdRecentFiles
Public g_RecentMacros As pdMRUManager

'If a double-click action closes a window (e.g. double-clicking a file from a common dialog), Windows incorrectly
' forwards the second click to the window behind the closed dialog.  To avoid this "click-through" behavior,
' this variable can be set to TRUE, which will prevent the underlying canvas from accepting input.  Just make sure
' to restore this variable to FALSE when you're done, including catching any error states!
Public g_DisableUserInput As Boolean

'As of v6.4, PhotoDemon supports a number of performance-related preferences.  Because performance settings (obviously)
' affect performance-sensitive parts of the program, these preferences are cached to global variables (rather than
' constantly pulled on-demand from file, which is unacceptably slow for performance-sensitive pipelines).
Public g_ViewportPerformance As PD_PerformanceSetting, g_InterfacePerformance As PD_PerformanceSetting

'As of v6.4, PhotoDemon allows the user to specify compression settings for Undo/Redo data.  By default, Undo/Redo data is
' uncompressed, which takes up a lot of (cheap) disk space but provides excellent performance.  The user can modify this
' setting to their liking, but they'll have to live with the performance implications.  The default setting for this value
' is 0, for no compression.
Public g_UndoCompressionLevel As Long

'Set this value to TRUE if you want PhotoDemon to report time-to-completion for various program actions.
' NOTE: this value is currently set automatically, in the LoadTheProgram sub.  PRE-ALPHA and ALPHA builds will report
'       timing for a variety of actions; BETA and PRODUCTION builds will not.  This can be overridden by changing the
'       activation code in LoadTheProgram.
Public g_DisplayTimingReports As Boolean

'Asynchronous tasks may require a modal wait screen.  To unload them successfully, we use a global flag that other
' asynchronous methods (like timers) can trigger.
Public g_UnloadWaitWindow As Boolean
