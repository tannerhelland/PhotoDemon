Attribute VB_Name = "Icon_and_Cursor_Handler"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright ©2012-2013 by Tanner Helland
'Created: 24/June/12
'Last updated: 08/May/13
'Last update: completed a very exciting function that allows me to use any 32bpp PNG file as a cursor.  No subclassing
'              required! First of its kind in VB, and a bitch to reverse-engineer... but well worth the effort, I think.
'
'Because VB6 doesn't provide many mechanisms for working with icons, I've had to manually add a number of icon-related
' functions to PhotoDemon.  First is a way to add icons/bitmaps to menus, as originally written by Leandro Ascierto.
' Menu icons are extracted from a resource file (where they're stored in PNG format) and rendered to the menu at run-time.
' See the clsMenuImage class for details on how this works. (A link to Leandro's original project can also be found there.)
'
'NOTE: Because the Windows XP version of Leandro's code utilizes potentially dirty subclassing, PhotoDemon automatically
' disables menu icons while running in the IDE on Windows XP.  Compile the project to see icons. (Windows Vista and 7 use
' a different mechanism, so menu icons are enabled in the IDE, and menu icons appear on all versions of Windows when compiled.)
'
'This module also handles the rendering of dynamic form, program, and taskbar icons.  (When an image is loaded and active,
' those icons can change to match the current image.)  As of February 2013, custom form icon generation has now been reworked
' based off this MSDN article: http://support.microsoft.com/kb/318876
' The new code is much leaner (and cleaner!) than past incarnations, and FreeImage is now required for the operation.  If
' FreeImage is not found, custom form icons will not be generated.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'SetWindowPos is used to force a repaint of the icon of maximized MDI child forms
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'API calls for building an icon at run-time
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'API call for manually setting a 32-bit icon to a form (as opposed to Form.Icon = ...)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'API needed for converting PNG data to icon or cursor format
Private Declare Sub CreateStreamOnHGlobal Lib "ole32" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hbmReturn As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hBmpReturn As Long, ByVal Background As Long) As GDIPlusStatus
Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal gdiImage As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal gdiImage As Long) As Long

'GDI+ types and constants
Private Const UnitPixel As Long = &H2&
Private Type RECTF
    fLeft As Single
    fTop As Single
    fWidth As Single
    fHeight As Single
End Type

'Type required to create an icon on-the-fly
Private Type ICONINFO
   fIcon As Boolean
   xHotspot As Long
   yHotspot As Long
   hbmMask As Long
   hbmColor As Long
End Type

'Used to apply and manage custom cursors (without subclassing)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Const IDC_APPSTARTING = 32650&
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

Private Const GCL_HCURSOR = (-12)

'These variables will hold the values of all custom-loaded cursors.
' They need to be deleted (via DestroyCursor) when the program exits; this is handled by unloadAllCursors.
Private hc_Handle_Arrow As Long
Private hc_Handle_Cross As Long
Public hc_Handle_Hand As Long       'The hand cursor handle is used by the jcButton control as well, so it is declared publicly.
Private hc_Handle_SizeAll As Long
Private hc_Handle_SizeNESW As Long
Private hc_Handle_SizeNS As Long
Private hc_Handle_SizeNWSE As Long
Private hc_Handle_SizeWE As Long

Dim numOfCustomCursors As Long
Dim customCursorNames() As String
Dim customCursorHandles() As Long

'This array will be used to store our dynamically created icon handles so we can delete them on program exit
Dim numOfIcons As Long
Dim iconHandles() As Long

'This constant is used for testing only.  It should always be set to TRUE for production code.
Public Const ALLOW_DYNAMIC_ICONS As Boolean = True

'This array tracks the resource identifiers and consequent numeric identifiers of all loaded icons.  The size of the array
' is arbitrary; just make sure it's larger than the max number of loaded icons.
Private iconNames(0 To 511) As String

'We also need to track how many icons have been loaded; this counter will also be used to reference icons in the database
Private curIcon As Long

'clsMenuImage does the heavy lifting for inserting icons into menus
Private cMenuImage As clsMenuImage

'A second class is used to manage the icons for the MRU list.
Private cMRUIcons As clsMenuImage

'Load all the menu icons from PhotoDemon's embedded resource file
Public Sub LoadMenuIcons()

    'If we are re-loading all icons instead of just loading them for the first time, clear out the old list
    If Not (cMenuImage Is Nothing) Then
        cMenuImage.Clear
        Set cMenuImage = Nothing
    End If
    
    'Reset the icon tracking array
    curIcon = 0
    Erase iconNames

    Set cMenuImage = New clsMenuImage
    
    With cMenuImage
            
        'Use Leandro's class to check if the current Windows install supports theming.
        g_IsThemingEnabled = .CanWeTheme
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not g_IsVistaOrLater) And (Not g_IsProgramCompiled) Then Exit Sub
        
        .Init FormMain.hWnd, 16, 16
        
    End With
            
    'Now that all menu icons are loaded, apply them to the proper menu entires
    ApplyAllMenuIcons
    
    'Finally, calculate where to place the "Clear MRU" menu item (this requires its own special handling).
    Dim numOfMRUFiles As Long
    numOfMRUFiles = MRU_ReturnCount()
    cMenuImage.PutImageToVBMenu 44, numOfMRUFiles + 1, 0, 1
    
    'And initialize the MRU icon handler.  (Unfortunately, MRU icons must be disabled on XP.  We can't
    ' double-subclass the main form, and using a single menu icon class isn't possible at present.)
    If g_IsVistaOrLater Then
        Set cMRUIcons = New clsMenuImage
        cMRUIcons.Init FormMain.hWnd, 64, 64
    End If
        
End Sub

'Apply (and if necessary, dynamically load) menu icons to their proper menu entries.
Public Sub ApplyAllMenuIcons()

    'Load every icon from the resource file.  (Yes, there are a LOT of icons!)
        
    'File Menu
    AddMenuIcon "OPENIMG", 0, 0       'Open Image
    AddMenuIcon "OPENREC", 0, 1       'Open recent
    AddMenuIcon "IMPORT", 0, 2        'Import
    AddMenuIcon "SAVE", 0, 4          'Save
    AddMenuIcon "SAVEAS", 0, 5        'Save As...
    AddMenuIcon "CLOSE", 0, 7         'Close
    AddMenuIcon "CLOSE", 0, 8         'Close All
    AddMenuIcon "BCONVERT", 0, 10     'Batch conversion
    AddMenuIcon "PRINT", 0, 12        'Print
    AddMenuIcon "EXIT", 0, 14         'Exit
    
    '--> Import Sub-Menu
    'NOTE: the specific menu values will be different if the scanner plugin (eztw32.dll) isn't found.
    If g_ScanEnabled Then
        AddMenuIcon "PASTE", 0, 2, 0      'From Clipboard (Paste as New Image)
        AddMenuIcon "SCANNER", 0, 2, 2    'Scan Image
        AddMenuIcon "SCANNERSEL", 0, 2, 3 'Select Scanner
        AddMenuIcon "DOWNLOAD", 0, 2, 5   'Online Image
        AddMenuIcon "SCREENCAP", 0, 2, 7  'Screen Capture
        AddMenuIcon "FRXIMPORT", 0, 2, 9  'Import from FRX
    Else
        AddMenuIcon "PASTE", 0, 2, 0      'From Clipboard (Paste as New Image)
        AddMenuIcon "DOWNLOAD", 0, 2, 2   'Online Image
        AddMenuIcon "SCREENCAP", 0, 2, 4  'Screen Capture
        AddMenuIcon "FRXIMPORT", 0, 2, 6  'Import from FRX
    End If
        
    'Edit Menu
    AddMenuIcon "UNDO", 1, 0           'Undo
    AddMenuIcon "REDO", 1, 1           'Redo
    AddMenuIcon "REPEAT", 1, 2         'Repeat Last Action
    AddMenuIcon "COPY", 1, 4           'Copy
    AddMenuIcon "PASTE", 1, 5          'Paste
    AddMenuIcon "CLEAR", 1, 6          'Empty Clipboard
    
    'View Menu
    AddMenuIcon "FITONSCREEN", 2, 0    'Fit on Screen
    AddMenuIcon "FITWINIMG", 2, 1      'Fit Viewport to Image
    AddMenuIcon "ZOOMIN", 2, 3         'Zoom In
    AddMenuIcon "ZOOMOUT", 2, 4        'Zoom Out
    AddMenuIcon "ZOOMACTUAL", 2, 10    'Zoom 100%
    AddMenuIcon "LEFTPANSHOW", 2, 16   'Show/Hide the left-hand panel
    AddMenuIcon "RIGHTPANSHOW", 2, 17  'Show/Hide the right-hand panel
    
    'Image Menu
    AddMenuIcon "DUPLICATE", 3, 0      'Duplicate
    AddMenuIcon "METADATA", 3, 2       'Metadata (top-level)
        '--> Metadata sub-menu
        AddMenuIcon "BROWSEMD", 3, 2, 0     'Browse metadata
        AddMenuIcon "MAPPHOTO", 3, 2, 2     'Map photo location
    AddMenuIcon "TRANSPARENCY", 3, 4   'Transparency
        '--> Image Mode sub-menu
        AddMenuIcon "ADDTRANS", 3, 4, 0     'Add alpha channel
        AddMenuIcon "REMOVETRANS", 3, 4, 1  'Remove alpha channel
    AddMenuIcon "RESIZE", 3, 6         'Resize
    AddMenuIcon "CROPSEL", 3, 7        'Crop to Selection
    AddMenuIcon "AUTOCROP", 3, 8       'Autocrop
    AddMenuIcon "MIRROR", 3, 10         'Mirror
    AddMenuIcon "FLIP", 3, 11           'Flip
    AddMenuIcon "ROTATECW", 3, 13      'Rotate Clockwise
    AddMenuIcon "ROTATECCW", 3, 14     'Rotate Counter-clockwise
    AddMenuIcon "ROTATE180", 3, 15     'Rotate 180
    'NOTE: the specific menu values will be different if the FreeImage plugin (FreeImage.dll) isn't found.
    If g_ImageFormats.FreeImageEnabled Then
        AddMenuIcon "ROTATEANY", 3, 16 'Rotate Arbitrary
        AddMenuIcon "ISOMETRIC", 3, 18 'Isometric
        AddMenuIcon "TILE", 3, 19      'Tile
    Else
        AddMenuIcon "ISOMETRIC", 3, 17 'Isometric
        AddMenuIcon "TILE", 3, 18      'Tile
    End If
    
    'Color Menu
    AddMenuIcon "BRIGHT", 4, 0         'Brightness/Contrast
    AddMenuIcon "COLORBALANCE", 4, 1   'Color balance
    AddMenuIcon "GAMMA", 4, 2          'Gamma Correction
    AddMenuIcon "HSL", 4, 3            'HSL adjustment
    AddMenuIcon "LEVELS", 4, 4         'Levels
    AddMenuIcon "SHDWHGHLGHT", 4, 5    'Shadow/Highlight
    AddMenuIcon "TEMPERATURE", 4, 6    'Temperature
    AddMenuIcon "WHITEBAL", 4, 7       'White Balance
    AddMenuIcon "HISTOGRAM", 4, 9      'Histogram
        '--> Histogram sub-menu
        AddMenuIcon "HISTOGRAM", 4, 9, 0  'Display Histogram
        AddMenuIcon "EQUALIZE", 4, 9, 2   'Equalize
        AddMenuIcon "STRETCH", 4, 9, 3    'Stretch
    AddMenuIcon "CHANNELMIX", 4, 11    'Components top-level
        '--> Components sub-menu
        AddMenuIcon "CHANNELMIX", 4, 11, 0   'Channel mixer
        AddMenuIcon "RECHANNEL", 4, 11, 1    'Rechannel
        AddMenuIcon "COLORSHIFTR", 4, 11, 3  'Shift Right
        AddMenuIcon "COLORSHIFTL", 4, 11, 4  'Shift Left
    AddMenuIcon "COLORIZE", 4, 13      'Colorize
    AddMenuIcon "ENHANCE", 4, 14       'Enhance
        '--> Enhance sub-menu
        AddMenuIcon "ENCONTRAST", 4, 14, 0    'Contrast
        AddMenuIcon "ENHIGHLIGHT", 4, 14, 1   'Highlights
        AddMenuIcon "ENMIDTONE", 4, 14, 2     'Midtones
        AddMenuIcon "ENSHADOW", 4, 14, 3      'Shadows
    AddMenuIcon "FADE", 4, 15           'Fade
        '--> Fade sub-menu
        AddMenuIcon "FADELOW", 4, 15, 0       'Low Fade
        AddMenuIcon "FADE", 4, 15, 1          'Medium Fade
        AddMenuIcon "FADEHIGH", 4, 15, 2      'High Fade
        AddMenuIcon "CUSTOMFADE", 4, 15, 3    'Custom Fade
        AddMenuIcon "UNFADE", 4, 15, 5        'Unfade
    AddMenuIcon "GRAYSCALE", 4, 16            'Grayscale
    AddMenuIcon "INVERT", 4, 17         'Invert
        '--> Invert sub-menu
        AddMenuIcon "INVCMYK", 4, 17, 0       'Invert CMYK
        AddMenuIcon "INVHUE", 4, 17, 1        'Invert Hue
        AddMenuIcon "INVRGB", 4, 17, 2        'Invert RGB
        AddMenuIcon "INVCOMPOUND", 4, 17, 4   'Compound Invert
    AddMenuIcon "MONOCHROME", 4, 18     'Monochrome
        '--> Monochrome sub-menu
        AddMenuIcon "COLORTOMONO", 4, 18, 0   'Color to monochrome
        AddMenuIcon "MONOTOCOLOR", 4, 18, 1   'Monochrome to grayscale
    AddMenuIcon "SEPIA", 4, 19          'Sepia
    AddMenuIcon "COUNTCOLORS", 4, 21    'Count Colors
    AddMenuIcon "REDUCECOLORS", 4, 22   'Reduce Colors
    
    'Filters Menu
    AddMenuIcon "FADELAST", 5, 0        'Fade Last
    AddMenuIcon "ARTISTIC", 5, 2        'Artistic
        '--> Artistic sub-menu
        AddMenuIcon "ANTIQUE", 5, 2, 0        'Antique (Sepia)
        AddMenuIcon "COMICBOOK", 5, 2, 1      'Comic book
        AddMenuIcon "FILMNOIR", 5, 2, 2       'Film Noir
        AddMenuIcon "MODERNART", 5, 2, 3      'Modern Art
        AddMenuIcon "PENCIL", 5, 2, 4         'Pencil
        AddMenuIcon "POSTERIZE", 5, 2, 5      'Posterize
        AddMenuIcon "RELIEF", 5, 2, 6         'Relief
    AddMenuIcon "BLUR", 5, 3            'Blur
        '--> Blur sub-menu
        AddMenuIcon "BOXBLUR", 5, 3, 0        'Box Blur
        AddMenuIcon "GAUSSBLUR", 5, 3, 1      'Gaussian Blur
        AddMenuIcon "GRIDBLUR", 5, 3, 2       'Grid Blur
        AddMenuIcon "PIXELATE", 5, 3, 3       'Pixelate (formerly Mosaic)
        AddMenuIcon "SMARTBLUR", 5, 3, 4      'Smart Blur
    AddMenuIcon "DISTORT", 5, 4      'Distort
        '--> Distort sub-menu
        AddMenuIcon "LENSDISTORT", 5, 4, 0    'Apply lens distortion
        AddMenuIcon "FIXLENS", 5, 4, 1        'Remove or correct existing lens distortion
        AddMenuIcon "FIGGLASS", 5, 4, 2       'Figured glass
        AddMenuIcon "KALEIDOSCOPE", 5, 4, 3   'Kaleidoscope
        AddMenuIcon "MISCDISTORT", 5, 4, 4    'Miscellaneous distort functions
        AddMenuIcon "PANANDZOOM", 5, 4, 5     'Pan and zoom
        AddMenuIcon "PERSPECTIVE", 5, 4, 6    'Perspective (free)
        AddMenuIcon "PINCHWHIRL", 5, 4, 7     'Pinch and whirl
        AddMenuIcon "POKE", 5, 4, 8           'Poke
        AddMenuIcon "POLAR", 5, 4, 9          'Polar conversion
        AddMenuIcon "RIPPLE", 5, 4, 10        'Ripple
        AddMenuIcon "SHEAR", 5, 4, 11         'Shear
        AddMenuIcon "SPHERIZE", 5, 4, 12      'Spherize
        AddMenuIcon "SQUISH", 5, 4, 13        'Squish (formerly Fixed Perspective)
        AddMenuIcon "SWIRL", 5, 4, 14         'Swirl
        AddMenuIcon "WAVES", 5, 4, 15         'Waves
        
    AddMenuIcon "EDGES", 5, 5        'Edges
        '--> Edges sub-menu
        AddMenuIcon "EMBOSS", 5, 5, 0         'Emboss / Engrave
        AddMenuIcon "EDGEENHANCE", 5, 5, 1    'Enhance Edges
        AddMenuIcon "EDGES", 5, 5, 2          'Find Edges
        AddMenuIcon "TRACECONTOUR", 5, 5, 3   'Trace Contour
    AddMenuIcon "OTHER", 5, 6        'Fun
        '--> Fun sub-menu
        AddMenuIcon "ALIEN", 5, 6, 0          'Alien
        AddMenuIcon "BLACKLIGHT", 5, 6, 1     'Blacklight
        AddMenuIcon "DREAM", 5, 6, 2          'Dream
        AddMenuIcon "RADIOACTIVE", 5, 6, 3    'Radioactive
        AddMenuIcon "SYNTHESIZE", 5, 6, 4     'Synthesize
        AddMenuIcon "HEATMAP", 5, 6, 5        'Thermograph
        AddMenuIcon "VIBRATE", 5, 6, 6        'Vibrate
    AddMenuIcon "NATURAL", 5, 7      'Natural
        '--> Natural sub-menu
        AddMenuIcon "ATMOSPHERE", 5, 7, 0     'Atmosphere
        AddMenuIcon "BURN", 5, 7, 1           'Burn
        AddMenuIcon "FOG", 5, 7, 2            'Fog
        AddMenuIcon "FREEZE", 5, 7, 3         'Freeze
        AddMenuIcon "LAVA", 5, 7, 4           'Lava
        AddMenuIcon "RAINBOW", 5, 7, 5        'Rainbow
        AddMenuIcon "STEEL", 5, 7, 6          'Steel
        AddMenuIcon "RAIN", 5, 7, 7           'Water
    AddMenuIcon "NOISE", 5, 8        'Noise
        '--> Noise sub-menu
        AddMenuIcon "FILMGRAIN", 5, 8, 0      'Film grain
        AddMenuIcon "ADDNOISE", 5, 8, 1       'Add Noise
        AddMenuIcon "DESPECKLE", 5, 8, 3      'Despeckle
        AddMenuIcon "MEDIAN", 5, 8, 4         'Median
        AddMenuIcon "REMOVEORPHAN", 5, 8, 5   'Remove Orphan
    AddMenuIcon "SHARPEN", 5, 9     'Sharpen
        '--> Sharpen sub-menu
        AddMenuIcon "SHARPEN", 5, 9, 0       'Sharpen
        AddMenuIcon "SHARPENMORE", 5, 9, 1   'Sharpen More
        AddMenuIcon "UNSHARP", 5, 9, 3       'Unsharp
    AddMenuIcon "STYLIZE", 5, 10     'Stylize
        '--> Stylize sub-menu
        AddMenuIcon "DIFFUSE", 5, 10, 0       'Diffuse
        AddMenuIcon "DILATE", 5, 10, 1        'Dilate
        AddMenuIcon "ERODE", 5, 10, 2         'Erode
        AddMenuIcon "PHOTOFILTER", 5, 10, 3   'Photo filters
        AddMenuIcon "SOLARIZE", 5, 10, 4      'Solarize
        AddMenuIcon "TWINS", 5, 10, 5         'Twins
        AddMenuIcon "VIGNETTE", 5, 10, 6      'Vignetting
    AddMenuIcon "CUSTFILTER", 5, 12  'Custom Filter
    
    'Tools Menu
    AddMenuIcon "LANGUAGES", 6, 0    'Languages
    AddMenuIcon "RECORD", 6, 2       'Macros
        '--> Macro sub-menu
        AddMenuIcon "OPENMACRO", 6, 2, 0      'Open Macro
        AddMenuIcon "RECORD", 6, 2, 2         'Start Recording
        AddMenuIcon "RECORDSTOP", 6, 2, 3     'Stop Recording
    AddMenuIcon "PREFERENCES", 6, 4           'Options (Preferences)
    AddMenuIcon "PLUGIN", 6, 5       'Plugin Manager
    
    'Window Menu
    AddMenuIcon "NEXTIMAGE", 7, 0    'Next image
    AddMenuIcon "PREVIMAGE", 7, 1    'Previous image
    AddMenuIcon "ARNGICONS", 7, 3    'Arrange Icons
    AddMenuIcon "CASCADE", 7, 4      'Cascade
    AddMenuIcon "TILEVER", 7, 5      'Tile Horizontally
    AddMenuIcon "TILEHOR", 7, 6      'Tile Vertically
    AddMenuIcon "MINALL", 7, 8       'Minimize All
    AddMenuIcon "RESTOREALL", 7, 9   'Restore All
    
    'Help Menu
    AddMenuIcon "FAVORITE", 8, 0     'Donate
    AddMenuIcon "UPDATES", 8, 2      'Check for updates
    AddMenuIcon "FEEDBACK", 8, 3     'Submit Feedback
    AddMenuIcon "BUG", 8, 4          'Submit Bug
    AddMenuIcon "PDWEBSITE", 8, 6    'Visit the PhotoDemon website
    AddMenuIcon "DOWNLOADSRC", 8, 7  'Download source code
    AddMenuIcon "LICENSE", 8, 8      'License
    AddMenuIcon "ABOUT", 8, 10       'About PD
    
End Sub

'This new, simpler technique for adding menu icons requires only the menu location (including sub-menus) and the icon's identifer
' in the resource file.  If the icon has already been loaded, it won't be loaded again; instead, the function will check the list
' of loaded icons and automatically fill in the numeric identifier as necessary.
Private Sub AddMenuIcon(ByVal resID As String, ByVal topMenu As Long, ByVal subMenu As Long, Optional ByVal subSubMenu As Long = -1)

    Dim i As Long
    Dim iconLocation As Long
    Dim iconAlreadyLoaded As Boolean
    
    iconAlreadyLoaded = False
    
    'Loop through all icons that have been loaded, and see if this one has been requested already.
    For i = 0 To curIcon
        
        If iconNames(i) = resID Then
            iconAlreadyLoaded = True
            iconLocation = i
            Exit For
        End If
        
    Next i
    
    'If the icon was not found, load it and add it to the list
    If Not iconAlreadyLoaded Then
        cMenuImage.AddImageFromStream LoadResData(resID, "CUSTOM")
        iconNames(curIcon) = resID
        iconLocation = curIcon
        curIcon = curIcon + 1
    End If
    
    'The position of menus changes if the MDI child is maximized.  When maximized, the form menu is given index 0, shifting
    ' everything to the right by one.
    
    'Thus, we must check for that and redraw the Undo/Redo menus accordingly
    Dim posModifier As Long
    posModifier = 0

    If NumOfWindows > 0 Then
        If FormMain.ActiveForm.WindowState = vbMaximized Then posModifier = 1
    End If
    
    'Place the icon onto the requested menu
    If subSubMenu = -1 Then
        cMenuImage.PutImageToVBMenu iconLocation, subMenu, topMenu + posModifier
    Else
        cMenuImage.PutImageToVBMenu iconLocation, subSubMenu, topMenu + posModifier, subMenu
    End If

End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' Note that to keep the code small, all changeable icons are refreshed whenever this is called.
Public Sub ResetMenuIcons()

    'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
    If (Not g_IsVistaOrLater) And (Not g_IsProgramCompiled) Then Exit Sub
        
    'Redraw the Undo/Redo menus
    With cMenuImage
        AddMenuIcon "UNDO", 1, 0     'Undo
        AddMenuIcon "REDO", 1, 1     'Redo
    End With
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    Dim numOfMRUFiles As Long
    numOfMRUFiles = MRU_ReturnCount()
    AddMenuIcon "CLEARRECENT", 0, 1, numOfMRUFiles + 1
    
    'Change the Show/Hide panel icon to match its current state
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideLeftPanel", False) Then
        AddMenuIcon "LEFTPANSHOW", 2, 16     'Show the panel
    Else
        AddMenuIcon "LEFTPANHIDE", 2, 16     'Hide the panel
    End If
    
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideRightPanel", False) Then
        AddMenuIcon "RIGHTPANSHOW", 2, 17   'Show the panel
    Else
        AddMenuIcon "RIGHTPANHIDE", 2, 17   'Hide the panel
    End If
        
    'If the OS is Vista or later, render MRU icons to the Open Recent menu
    If g_IsVistaOrLater Then
    
        'The position of menus changes if the MDI child is maximized.  When maximized, the form menu is given index 0, shifting
        ' everything to the right by one.
        
        'Thus, we must check for that and redraw the Undo/Redo menus accordingly
        Dim posModifier As Long
        posModifier = 0
    
        If NumOfWindows > 0 Then
            If FormMain.ActiveForm.WindowState = vbMaximized Then posModifier = 1
        End If
    
        cMRUIcons.Clear
        Dim tmpFilename As String
    
        'Load a placeholder image for missing MRU entries
        cMRUIcons.AddImageFromStream LoadResData("MRUHOLDER", "CUSTOM")
    
        'This counter will be used to track the current position of loaded thumbnail images into the icon collection
        Dim iconLocation As Long
        iconLocation = 0
    
        'Loop through the MRU list, and attempt to load thumbnail images for each entry
        Dim i As Long
        For i = 0 To numOfMRUFiles
        
            'Start by seeing if an image exists for this MRU entry
            tmpFilename = getMRUThumbnailPath(i)
        
            'If the file exists, add it to the MRU icon handler
            If FileExist(tmpFilename) Then
                
                iconLocation = iconLocation + 1
                cMRUIcons.AddImageFromFile tmpFilename
                cMRUIcons.PutImageToVBMenu iconLocation, i, 0 + posModifier, 1
            
            Else
                cMRUIcons.PutImageToVBMenu 0, i, 0 + posModifier, 1
            End If
        
        Next i
                
    End If
        
End Sub


'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
' Past versions of this sub included a pure-VB fallback if FreeImage wasn't found.  These have since been removed -
' so FreeImage is 100% required for this sub to operate.
Public Sub CreateCustomFormIcon(ByRef imgForm As FormImage)

    If Not ALLOW_DYNAMIC_ICONS Then Exit Sub
    If Not g_ImageFormats.FreeImageEnabled Then Exit Sub

    'Generating an icon requires many variables; see below for specific comments on each one
    Dim MonoBmp As Long
    Dim icoInfo As ICONINFO
    Dim generatedIcon As Long
   
    'The icon can be drawn at any size, but 16x16 is how it will (typically) end up on the form.  Since we are now rendering
    ' a dynamically generated icon to the task bar as well, we opt for 32x32, and from that we generate an additional 16x16 version.
    Dim icoSize As Long
    
    'If we are rendering a dynamic taskbar icon, we will perform two reductions - first to 32x32, second to 16x16
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) Then icoSize = 32 Else icoSize = 16

    'Determine aspect ratio
    Dim aspectRatio As Double
    aspectRatio = CSng(pdImages(imgForm.Tag).Width) / CSng(pdImages(imgForm.Tag).Height)
    
    'The target icon's width and height, x and y positioning
    Dim tIcoWidth As Double, tIcoHeight As Double, TX As Double, TY As Double
    
    'If the form is wider than it is tall...
    If aspectRatio > 1 Then
        
        'Determine proper sizes and (x, y) positioning so the icon will be centered
        tIcoWidth = icoSize
        tIcoHeight = icoSize * (1 / aspectRatio)
        TX = 0
        TY = (icoSize - tIcoHeight) / 2
        
    Else
    
        'Same thing, but with the math adjusted for images taller than they are wide
        tIcoHeight = icoSize
        tIcoWidth = icoSize * aspectRatio
        TY = 0
        TX = (icoSize - tIcoWidth) / 2
        
    End If
    
    'Load the FreeImage dll into memory
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
    
    'Convert our current layer to a FreeImage-type DIB
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
    
    'Use that handle to request an image resize
    If fi_DIB <> 0 Then
            
        'Rescale the image
        Dim returnDIB As Long
        returnDIB = FreeImage_RescaleByPixel(fi_DIB, CLng(tIcoWidth), CLng(tIcoHeight), True, FILTER_BILINEAR)
        
        'Make sure the image is 32bpp (returns a clone of the image if it's already 32bpp, so no harm done)
        Dim newDIB32 As Long
        newDIB32 = FreeImage_ConvertTo32Bits(returnDIB)
        
        'Unload the original DIB
        If newDIB32 <> returnDIB Then FreeImage_UnloadEx returnDIB
            
        'If the image isn't square-shaped, we need to enlarge the DIB accordingly. FreeImage provides a function for that.
        
        'Also, set the background of the enlarged area as transparent
        Dim newColor As RGBQUAD
        With newColor
            .rgbBlue = 255
            .rgbGreen = 255
            .rgbRed = 255
            .rgbReserved = 0
        End With
            
        'Enlarge the canvas as necessary
        Dim finalDIB As Long
        finalDIB = FreeImage_EnlargeCanvas(newDIB32, TX, TY, TX, TY, newColor, FI_COLOR_IS_RGBA_COLOR)
        
        'Unload the original DIB
        If finalDIB <> newDIB32 Then FreeImage_UnloadEx newDIB32
            
        'At this point, finalDIB contains the 32bpp alpha icon exactly how we want it.
        
        'If we are dynamically updating the taskbar icon to match the current image, we need to assign the 32x32 icon now
        If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) Then
                
            'Generate a blank monochrome mask to pass to the icon creation function.
            MonoBmp = CreateBitmap(icoSize, icoSize, 1, 1, ByVal 0&)
                        
            'With the transfer complete, release the FreeImage DIB and unload the library
            If finalDIB <> 0 Then
                With icoInfo
                    .fIcon = True
                    .xHotspot = icoSize
                    .yHotspot = icoSize
                    .hbmMask = MonoBmp
                    .hbmColor = FreeImage_GetBitmapForDevice(finalDIB)
                End With
            End If
                
            'Create the 32x32 icon
            generatedIcon = CreateIconIndirect(icoInfo)
            
            'Assign it to the taskbar
            setNewTaskbarIcon generatedIcon
            
            '...and remember it in our current icon collection
            addIconToList generatedIcon
                
            '...and the current form
            pdImages(imgForm.Tag).curFormIcon32 = generatedIcon
                                    
            'Now delete the temporary mask and bitmap
            DeleteObject MonoBmp
            DeleteObject icoInfo.hbmColor
            
        End If
            
        'Finally, resize the 32x32 icon to 16x16 so it will work as the current form icon as well
        icoSize = 16
        finalDIB = FreeImage_RescaleByPixel(finalDIB, 16, 16, True, FILTER_BILINEAR)
            
        'Generate a blank monochrome mask to pass to the icon creation function.
        MonoBmp = CreateBitmap(icoSize, icoSize, 1, 1, ByVal 0&)
            
        'With the transfer complete, release the FreeImage DIB and unload the library
        If finalDIB <> 0 Then
            With icoInfo
                .fIcon = True
                .xHotspot = icoSize
                .yHotspot = icoSize
                .hbmMask = MonoBmp
                .hbmColor = FreeImage_GetBitmapForDevice(finalDIB)
            End With
        End If
            
        'Render the icon to a handle and store it in our running list, so we can destroy it when the program is closed
        generatedIcon = CreateIconIndirect(icoInfo)
        addIconToList generatedIcon
        
        'If we are dynamically updating the taskbar icon to match the current image, we need to assign the 16x16 icon now
        If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) Then
            setNewAppIcon generatedIcon
            pdImages(imgForm.Tag).curFormIcon16 = generatedIcon
        End If
        
        'Clear out memory
        FreeImage_UnloadEx finalDIB
        FreeLibrary hLib
        DeleteObject MonoBmp
        DeleteObject icoInfo.hbmColor
        
        'Use the API to assign this new icon to the specified MDI child form
        SendMessageLong imgForm.hWnd, &H80, 0, generatedIcon
        
        'When an MDI child form is maximized, the icon is not updated properly - so we must force a manual refresh of the entire window frame.
        If imgForm.WindowState = vbMaximized Then SetWindowPos FormMain.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
        
    End If
       
End Sub
'Needs to be run only once, at the start of the program
Public Sub initializeIconHandler()
    numOfIcons = 0
End Sub

'Add another icon reference to the list
Private Sub addIconToList(ByVal hIcon As Long)

    ReDim Preserve iconHandles(0 To numOfIcons) As Long
    iconHandles(numOfIcons) = hIcon
    numOfIcons = numOfIcons + 1

End Sub

'Remove all icons generated since the program launched
Public Sub destroyAllIcons()

    If numOfIcons = 0 Then Exit Sub
    
    Dim i As Long
    For i = 0 To numOfIcons - 1
        DestroyIcon iconHandles(i)
    Next i
    
    numOfIcons = 0
    
    ReDim iconHandles(0) As Long

End Sub

'Given an image in the .exe's resource section (typically a PNG image), return an icon handle to it (hIcon).
' The calling function is responsible for deleting this object once they are done with it.
Public Function createIconFromResource(ByVal resTitle As String) As Long
    
    'Start by extracting the PNG data into a bytestream
    Dim ImageData() As Byte
    ImageData() = LoadResData(resTitle, "CUSTOM")
    
    Dim IStream As IUnknown
    Dim hBitmap As Long, hIcon As Long
        
    CreateStreamOnHGlobal ImageData(0), 0&, IStream
    
    If Not IStream Is Nothing Then
        
        'Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine
        If GdipLoadImageFromStream(IStream, hBitmap) = 0 Then
        
            'hBitmap now contains the PNG file as an hBitmap (obviously).  Now we need to convert it to icon format.
            If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
                createIconFromResource = hIcon
            Else
                createIconFromResource = 0
            End If
            
            GdipDisposeImage hBitmap
                
        End If
    
        Set IStream = Nothing
    
    End If
    
    Exit Function
    
End Function

'Given an image in the .exe's resource section (typically a PNG image), return it as a cursor object.
' The calling function is responsible for deleting the cursor once they are done with it.
Public Function createCursorFromResource(ByVal resTitle As String, Optional ByVal curHotSpotX As Long = 8, Optional ByVal curHotSpotY As Long = 16) As Long
    
    'Start by extracting the PNG data into a bytestream
    Dim ImageData() As Byte
    ImageData() = LoadResData(resTitle, "CUSTOM")
    
    Dim IStream As IUnknown
    Dim tmpRect As RECTF
    Dim gdiBitmap As Long, hBitmap As Long, hIcon As Long
        
    CreateStreamOnHGlobal ImageData(0), 0&, IStream
    
    If Not IStream Is Nothing Then
        
        'Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine
        If GdipLoadImageFromStream(IStream, gdiBitmap) = 0 Then
        
            'Retrieve the image's size
            GdipGetImageBounds gdiBitmap, tmpRect, UnitPixel
        
            'Convert the GDI+ bitmap to a standard Windows hBitmap
            If GdipCreateHBITMAPFromBitmap(gdiBitmap, hBitmap, vbBlack) = 0 Then
                        
                'Generate a blank monochrome mask to pass to the icon creation function.
                Dim MonoBmp As Long
                MonoBmp = CreateBitmap(CLng(tmpRect.fWidth), CLng(tmpRect.fHeight), 1, 1, ByVal 0&)
                            
                'With the transfer complete, release the FreeImage DIB and unload the library
                Dim icoInfo As ICONINFO
                With icoInfo
                    .fIcon = False
                    .xHotspot = curHotSpotX
                    .yHotspot = curHotSpotY
                    .hbmMask = MonoBmp
                    .hbmColor = hBitmap
                End With
                    
                'Create the 32x32 cursor
                createCursorFromResource = CreateIconIndirect(icoInfo)
                
                DeleteObject MonoBmp
                DeleteObject hBitmap
            
            End If
            
            GdipDisposeImage gdiBitmap
                
        End If
    
        Set IStream = Nothing
    
    End If
    
    Exit Function
    
End Function

'Load all system cursors into memory
Public Sub InitAllCursors()

    hc_Handle_Arrow = LoadCursor(0, IDC_ARROW)
    hc_Handle_Cross = LoadCursor(0, IDC_CROSS)
    hc_Handle_Hand = LoadCursor(0, IDC_HAND)
    hc_Handle_SizeAll = LoadCursor(0, IDC_SIZEALL)
    hc_Handle_SizeNESW = LoadCursor(0, IDC_SIZENESW)
    hc_Handle_SizeNS = LoadCursor(0, IDC_SIZENS)
    hc_Handle_SizeNWSE = LoadCursor(0, IDC_SIZENWSE)
    hc_Handle_SizeWE = LoadCursor(0, IDC_SIZEWE)

End Sub

'Remove the hand cursor from memory
Public Sub unloadAllCursors()
    DestroyCursor hc_Handle_Hand
    DestroyCursor hc_Handle_Arrow
    DestroyCursor hc_Handle_Cross
    DestroyCursor hc_Handle_SizeAll
    DestroyCursor hc_Handle_SizeNESW
    DestroyCursor hc_Handle_SizeNS
    DestroyCursor hc_Handle_SizeNWSE
    DestroyCursor hc_Handle_SizeWE
    
    Dim i As Long
    For i = 0 To numOfCustomCursors - 1
        DestroyCursor customCursorHandles(i)
    Next i
    
End Sub

'Use any 32bpp PNG resource as a cursor (yes, it's amazing!)
Public Sub setPNGCursorToHwnd(ByVal dstHwnd As Long, ByVal pngTitle As String)
    SetClassLong dstHwnd, GCL_HCURSOR, requestCustomCursor(pngTitle)
End Sub

'Set a single object to use the hand cursor
Public Sub setHandCursor(ByRef tControl As Control)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Hand
End Sub

Public Sub setHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, hc_Handle_Hand
End Sub

'Set a single object to use the arrow cursor
Public Sub setArrowCursorToObject(ByRef tControl As Control)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

Public Sub setArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

'Set a single form to use the arrow cursor
Public Sub setArrowCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

'Set a single form to use the cross cursor
Public Sub setCrossCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Cross
End Sub
    
'Set a single form to use the Size All cursor
Public Sub setSizeAllCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeAll
End Sub

'Set a single form to use the Size NESW cursor
Public Sub setSizeNESWCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNESW
End Sub

'Set a single form to use the Size NS cursor
Public Sub setSizeNSCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNS
End Sub

'Set a single form to use the Size NWSE cursor
Public Sub setSizeNWSECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNWSE
End Sub

'Set a single form to use the Size WE cursor
Public Sub setSizeWECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeWE
End Sub

'If a custom PNG cursor has not been loaded, this function will load the PNG, convert it to cursor format, then store
' the cursor resource for future reference (so the image doesn't have to be loaded again).
Private Function requestCustomCursor(ByVal cursorName As String) As Long

    Dim i As Long
    Dim cursorLocation As Long
    Dim cursorAlreadyLoaded As Boolean
    
    cursorLocation = 0
    cursorAlreadyLoaded = False
    
    'Loop through all cursors that have been loaded, and see if this one has been requested already.
    If numOfCustomCursors > 0 Then
    
        For i = 0 To numOfCustomCursors - 1
        
            If customCursorNames(i) = cursorName Then
                cursorAlreadyLoaded = True
                cursorLocation = i
                Exit For
            End If
        
        Next i
    
    End If
    
    'If the cursor was not found, load it and add it to the list
    If cursorAlreadyLoaded Then
        requestCustomCursor = customCursorHandles(cursorLocation)
    Else
        Dim tmpHandle As Long
        tmpHandle = createCursorFromResource(cursorName)
        
        ReDim Preserve customCursorNames(0 To numOfCustomCursors) As String
        ReDim Preserve customCursorHandles(0 To numOfCustomCursors) As Long
        
        customCursorNames(numOfCustomCursors) = cursorName
        customCursorHandles(numOfCustomCursors) = tmpHandle
        
        numOfCustomCursors = numOfCustomCursors + 1
        
        requestCustomCursor = tmpHandle
    End If

End Function
