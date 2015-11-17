Attribute VB_Name = "Public_Variables"

'Contains any and all publicly-declared variables.  I am trying to move
' all public variables here (for obvious reasons), but the transition may
' not be completely done as long as this comment remains!

Option Explicit


'The number of images PhotoDemon has loaded this session (always goes up, never down; starts at zero when the program is loaded).
' This value correlates to the upper bound of the primary pdImages array.  For performance reasons, that array is not dynamically
' resized when images are loaded - the array stays the same size, and entries are deactivated as needed.  Thus, WHENEVER YOU
' NEED TO ITERATE THROUGH ALL LOADED IMAGES, USE THIS VALUE INSTEAD OF g_OpenImageCount.
Public g_NumOfImagesLoaded As Long

'The ID number (e.g. index in the pdImages array) of image the user is currently interacting with (e.g. the currently active image
' window).  Whenever a function needs to access the current image, use pdImages(g_CurrentImage).
Public g_CurrentImage As Long

'Number of image windows CURRENTLY OPEN.  This value goes up and down as images are opened or closed.  Use it to test for no open
' images (e.g. If g_OpenImageCount = 0...).  Note that this value SHOULD NOT BE USED FOR ITERATING OPEN IMAGES.  Instead, use
' g_NumOfImagesLoaded, which will always match the upper bound of the pdImages() array, and never decrements, even when images
' are unloaded.
Public g_OpenImageCount As Long

'This array is the heart and soul of a given PD session.  Every time an image is loaded, all of its relevant data is stored within
' a new entry in this array.
Public pdImages() As pdImage

'Main user preferences and settings handler
Public g_UserPreferences As pdPreferences

'Main file format compatibility handler
Public g_ImageFormats As pdFormats

'Main language and translation handler
Public g_Language As pdTranslate

'Main clipboard handler
Public g_Clipboard As pdClipboardMain

'Currently selected tool, previous tool
Public g_CurrentTool As PDTools
Public g_PreviousTool As PDTools

'Primary zoom handler for the program
Public g_Zoom As pdZoom

'Whether or not to resize large images to fit on-screen (0 means "yes," 1 means "no")
Public g_AutozoomLargeImages As Long

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

'Because FreeImage is used far more than any other plugin, we no longer load it on-demand.  It is loaded once
' when the program starts, and released when the program ends.  This saves us from repeatedly having to load/free
' the entire library (which is fairly large).  This variable stores our received library handle.
Public g_FreeImageHandle As Long

'How to draw the background of image forms; -1 is checkerboard, any other value is treated as an RGB long
Public g_CanvasBackground As Long

'Does the user want us to prompt them when they try to close unsaved images?
Public g_ConfirmClosingUnsaved As Boolean

'Whether or not we are running in the IDE or compiled
Public g_IsProgramCompiled As Boolean

'Per the excellent advice of Kroc (camendesign.com), a custom UserMode variable is less prone to errors than the usual
' Ambient.UserMode value supplied to ActiveX controls.  This fixes a problem where ActiveX controls sometimes think they
' are being run in a compiled EXE, when actually their properties are just being written as part of .exe compiling.
Public g_IsProgramRunning As Boolean

'Temporary loading variable to disable Autozoom feature
Public g_AllowViewportRendering As Boolean

'For the Open and Save common dialog boxes, it's polite to remember what format the user used last, then default
' the boxes to that.  (Note that these values are stored in the preferences file as well, but that is only accessed
' upon program load and unload.)
Public g_LastOpenFilter As Long
Public g_LastSaveFilter As Long

'DIB that contains a 2x2 pattern of the alpha checkerboard.  Use it with CreatePatternBrush to paint the alpha
' checkerboard prior to rendering.
Public g_CheckerboardPattern As pdDIB

'Is the current system running Vista, 7, 8, or later?  (Used to determine availability of certain system features)
Public g_IsVistaOrLater As Boolean
Public g_IsWin7OrLater As Boolean
Public g_IsWin8OrLater As Boolean
Public g_IsWin81OrLater As Boolean
Public g_IsWin10OrLater As Boolean

'Is theming enabled?  (Used to handle some menu icon rendering quirks)
Public g_IsThemingEnabled As Boolean

'Render the interface using Segoe UI if available; g_UseFancyFonts will be set to FALSE if we have to fall back to Tahoma
Public g_UseFancyFonts As Boolean
Public g_InterfaceFont As String

'This g_Displays object contains data on all display devices on this system.  It includes a ton of code to assist the program
' with managing multiple monitors and other display-related issues.
Public g_Displays As pdDisplays

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
Public g_JPEGAutoQuality As jpegAutoQualityMode
Public g_JPEGAdvancedColorMatching As Boolean

'JPEG-2000 export compression ratio; this is set by the JP2 export dialog if the user clicks "OK" (not Cancel)
Public g_JP2Compression As Long

'WebP export compression ratio; this is set by the WebP export dialog if the user clicks "OK" (not Cancel)
Public g_WebPCompression As Long

'JPEG XR export settings; these are set by the JPEG XR export dialog if the user clicks "OK" (not Cancel)
Public g_JXRCompression As Long
Public g_JXRProgressive As Boolean

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

'What background color should be used for compositing an image's complex alpha channel?
' (This is also set by the custom alpha cutoff dialog.)
Public g_AlphaCompositeColor As Long

'When an image has its colors counted, the image's ID is stored here.  Other functions can use this to see if the
' current color count is relevant for a given image (e.g. if the image being worked on has just had its colors counted).
Public g_LastImageScanned As Long

'If this is the first time the user has run PhotoDemon (as determined by the lack of a preferences XML file), this
' variable will be set to TRUE early in the load process.  Other routines can then modify their behavior accordingly.
Public g_IsFirstRun As Boolean

'Drag and drop operations are allowed at certain times, but not others.  Any time a modal form is displayed, drag-and-drop
' must be disallowed - with the exception of common dialog boxes.  To make sure this behavior is carefully maintained,
' we track drag-and-drop enabling ourselves
Public g_AllowDragAndDrop As Boolean

'While Undo/Redo operations are active, certain tasks can be ignored.  This public value can be used to check Undo/Redo activity.
Public g_UndoRedoActive As Boolean

'GDI+ availability is determined at the very start of the program; we rely on it heavily, so expect problems if
' it can't be initialized!
Public g_GDIPlusAvailable As Boolean

'PhotoDemon's primary window manager.  This handles positioning, layering, and sizing of all windows in the project.
Public g_WindowManager As pdWindowManager

'PhotoDemon's visual theme engine.
Public g_Themer As pdVisualThemes

'PhotoDemon's recent files and recent macros managers.
' CHANGING: Replacing pdRecentFiles with pdMRUManager
Public g_RecentFiles As pdMRUManager
Public g_RecentMacros As pdMRUManager

'To improve mousewheel handling, we dynamically track the position of the mouse.  If it is over the image tabstrip,
' the main form will forward mousewheel events there; otherwise, the image window gets them.
Public g_MouseOverImageTabstrip As Boolean

'Global color management setting.  If the user has requested use of custom profiles, this will be set to FALSE.
Public g_UseSystemColorProfile As Boolean

'Mouse accuracy for collision detection with on-screen objects.  This is typically 6 pixels, but it's re-calculated
' at run-time to account for high-DPI screens.  (It may even be worthwhile to let users adjust this value, or to
' retrieve some system metric for it... if such a thing exists.)
Public g_MouseAccuracy As Double

'If a double-click action closes a window (e.g. double-clicking a file from a common dialog), Windows incorrectly
' forwards the second click to the window behind the closed dialog.  To avoid this "click-through" behavior,
' this variable can be set to TRUE, which will prevent the underlying canvas from accepting input.  Just make sure
' to restore this variable to FALSE when you're done, including catching any error states!
Public g_DisableUserInput As Boolean

'Last message sent to PD's central Message() function.  Note that this string *includes any custom attachments* and is calculated
' *post-translation*, e.g. instead of being "Error %1", the "%1" will be populated with whatever value was supplied, and "Error"
' will be translated into the currently active language.  The purpose of this variable is to assist asynchronous functions.  When
' such functions terminate, they can cache the previous message, display any relevant messages according to their asynchronous
' results, then restore the original message when done.  This makes the experience seamless for the user, but is hugely helpful
' to me when debugging asynchronous program behavior.
Public g_LastPostedMessage As String

'ID for this PD instance.  When started, each PhotoDemon instance is assigned a pseudo-random (GUID-based) session ID, which it
' then appends to things like Undo/Redo files.  This allows for multiple side-by-side program instances without collisions.
Public g_SessionID As String

'As of v6.4, PhotoDemon supports a number of performance-related preferences.  Because performance settings (obviously)
' affect performance-sensitive parts of the program, these preferences are cached to global variables (rather than
' constantly pulled on-demand from file, which is unacceptably slow for performance-sensitive pipelines).
Public g_ViewportPerformance As PD_PERFORMANCE_SETTING
Public g_ThumbnailPerformance As PD_PERFORMANCE_SETTING
Public g_InterfacePerformance As PD_PERFORMANCE_SETTING
Public g_ColorPerformance As PD_PERFORMANCE_SETTING

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

'PhotoDemon's central debugger.  This class is accessed by pre-alpha, alpha, and beta builds, and it is used to log
' generic debug messages on client PCs, which we can (hopefully) use to recreate crashes as necessary.
Public pdDebug As pdDebugger

'If FreeImage throws an error, the error string(s) will be stored here.  Make sure to clear it after reading to prevent future
' functions from mistakenly displaying the same message!
Public g_FreeImageErrorMessages() As String

'As part of an improved memory efficiency initiative, some global variables are used (during debug mode) to track how many
' GDI objects PD creates and destroys.
Public g_DIBsCreated As Long
Public g_DIBsDestroyed As Long
Public g_FontsCreated As Long
Public g_FontsDestroyed As Long
Public g_DCsCreated As Long
Public g_DCsDestroyed As Long

'If a modal window is active, this value will be set to TRUE.  This is helpful for controlling certain program flow issues.
Public g_ModalDialogActive As Boolean

'High-resolution input tracking allows for much more accurate reproduction of mouse values.  However, old PCs may struggle to
' cope with all the extra input data.  A user-facing preference allows for disabling this behavior.
Public g_HighResolutionInput As Boolean

'If an update notification is ready, but we can't display it (for example, because a modal dialog is active) this flag will
' be set to TRUE.  PD's central processor uses this to display the update notification as soon as it reasonably can.
Public g_ShowUpdateNotification As Boolean

'If an update has been successfully applied, the user is given the option to restart PD immediately.  If the user chooses
' to restart, this global value will be set to TRUE.
Public g_UserWantsRestart As Boolean

'If this PhotoDemon session was started by a restart (because an update patch was applied), this will be set to TRUE.
' PD uses this value to suspend any other automatic updates, as a precaution against any bugs in the updater.
Public g_ProgramStartedViaRestart As Boolean

