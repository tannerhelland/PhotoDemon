Attribute VB_Name = "Icon_and_Cursor_Handler"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright ©2012-2013 by Tanner Helland
'Created: 24/June/12
'Last updated: 12/October/13
'Last update: remove some MDI-specific code that is no longer needed
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API calls for building an icon at run-time
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'API call for manually setting a 32-bit icon to a form (as opposed to Form.Icon = ...)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'API needed for converting PNG data to icon or cursor format
Private Declare Sub CreateStreamOnHGlobal Lib "ole32" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus" (ByVal Stream As Any, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus" (ByVal gdiBitmap As Long, ByRef hbmReturn As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus" (ByVal gdiBitmap As Long, ByRef hBmpReturn As Long, ByVal Background As Long) As GDIPlusStatus
Private Declare Function GdipGetImageBounds Lib "GdiPlus" (ByVal gdiBitmap As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal gdiBitmap As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GdiPlus" (ByVal gdiBitmap As Long, ByRef PixelFormat As Long) As Long

'Used to check GDI+ images for alpha channels
Private Const PixelFormatAlpha = &H40000             ' Has an alpha component
Private Const PixelFormatPAlpha = &H80000            ' Pre-multiplied alpha

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

'Some functions in this module take a long time to apply.  In order to refresh a generic progress bar on the "please wait" dialog,
' this module-level variable can be set to TRUE.
Private m_refreshOutsideProgressBar As Boolean

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
Public Sub ApplyAllMenuIcons(Optional ByVal useDoEvents As Boolean = False)

    m_refreshOutsideProgressBar = useDoEvents

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
        AddMenuIcon "PASTE", 0, 2, 0       'From Clipboard (Paste as New Image)
        AddMenuIcon "SCANNER", 0, 2, 2     'Scan Image
        AddMenuIcon "SCANNERSEL", 0, 2, 3  'Select Scanner
        AddMenuIcon "DOWNLOAD", 0, 2, 5    'Online Image
        AddMenuIcon "SCREENCAP", 0, 2, 7   'Screen Capture
    Else
        AddMenuIcon "PASTE", 0, 2, 0       'From Clipboard (Paste as New Image)
        AddMenuIcon "DOWNLOAD", 0, 2, 2    'Online Image
        AddMenuIcon "SCREENCAP", 0, 2, 4   'Screen Capture
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
    
    'Image Menu
    AddMenuIcon "DUPLICATE", 3, 0      'Duplicate
    AddMenuIcon "TRANSPARENCY", 3, 2   'Transparency
        '--> Image Mode sub-menu
        AddMenuIcon "ADDTRANS", 3, 2, 0      'Add alpha channel
        AddMenuIcon "GREENSCREEN", 3, 2, 1      'Color to alpha
        AddMenuIcon "REMOVETRANS", 3, 2, 3   'Remove alpha channel
    AddMenuIcon "RESIZE", 3, 4         'Resize
    AddMenuIcon "CANVASSIZE", 3, 5     'Canvas resize
    AddMenuIcon "CROPSEL", 3, 7        'Crop to Selection
    AddMenuIcon "AUTOCROP", 3, 8       'Autocrop
    AddMenuIcon "ROTATECW", 3, 10      'Rotate top-level
        '--> Rotate sub-menu
        AddMenuIcon "ROTATECW", 3, 10, 0     'Rotate Clockwise
        AddMenuIcon "ROTATECCW", 3, 10, 1    'Rotate Counter-clockwise
        AddMenuIcon "ROTATE180", 3, 10, 2    'Rotate 180
        If g_ImageFormats.FreeImageEnabled Then AddMenuIcon "ROTATEANY", 3, 10, 3   'Rotate Arbitrary
    AddMenuIcon "MIRROR", 3, 11        'Mirror
    AddMenuIcon "FLIP", 3, 12          'Flip
    AddMenuIcon "ISOMETRIC", 3, 13     'Isometric
    AddMenuIcon "REDUCECOLORS", 3, 15  'Indexed color (Reduce Colors)
    If g_ImageFormats.FreeImageEnabled Then FormMain.MnuImage(15).Enabled = True Else FormMain.MnuImage(15).Enabled = False
    AddMenuIcon "TILE", 3, 16          'Tile
    AddMenuIcon "METADATA", 3, 18      'Metadata (top-level)
        '--> Metadata sub-menu
        AddMenuIcon "BROWSEMD", 3, 18, 0     'Browse metadata
        AddMenuIcon "COUNTCOLORS", 3, 18, 2  'Count Colors
        AddMenuIcon "MAPPHOTO", 3, 18, 3     'Map photo location
    
    'Select Menu
    AddMenuIcon "SELECTALL", 4, 0       'Select all
    AddMenuIcon "SELECTNONE", 4, 1      'Select none
    AddMenuIcon "SELECTINVERT", 4, 2    'Invert selection
    AddMenuIcon "SELECTGROW", 4, 4      'Grow selection
    AddMenuIcon "SELECTSHRINK", 4, 5    'Shrink selection
    AddMenuIcon "SELECTBORDER", 4, 6    'Border selection
    AddMenuIcon "SELECTFTHR", 4, 7      'Feather selection
    AddMenuIcon "SELECTSHRP", 4, 8      'Sharpen selection
    AddMenuIcon "SELECTLOAD", 4, 10     'Load selection from file
    AddMenuIcon "SELECTSAVE", 4, 11     'Save selection to file
    
    'Adjustments Menu
    
    'Adjustment shortcuts (top-level menu items)
    AddMenuIcon "GRAYSCALE", 5, 0       'Black and white
    AddMenuIcon "BRIGHT", 5, 1          'Brightness/Contrast
    AddMenuIcon "COLORBALANCE", 5, 2    'Color balance
    AddMenuIcon "CURVES", 5, 3          'Curves
    AddMenuIcon "LEVELS", 5, 4          'Levels
    AddMenuIcon "VIBRANCE", 5, 5        'Vibrance
    AddMenuIcon "WHITEBAL", 5, 6        'White Balance
       
    'Channels
    AddMenuIcon "CHANNELMIX", 5, 8     'Channels top-level
        AddMenuIcon "CHANNELMIX", 5, 8, 0    'Channel mixer
        AddMenuIcon "RECHANNEL", 5, 8, 1     'Rechannel
        AddMenuIcon "CHANNELMAX", 5, 8, 3    'Channel max
        AddMenuIcon "CHANNELMIN", 5, 8, 4    'Channel min
        AddMenuIcon "COLORSHIFTL", 5, 8, 6   'Shift Left
        AddMenuIcon "COLORSHIFTR", 5, 8, 7   'Shift Right
            
    'Color
    AddMenuIcon "HSL", 5, 9            'Color balance
        AddMenuIcon "COLORBALANCE", 5, 9, 0  'Color balance
        AddMenuIcon "WHITEBAL", 5, 9, 1      'White Balance
        AddMenuIcon "HSL", 5, 9, 3           'HSL adjustment
        AddMenuIcon "PHOTOFILTER", 5, 9, 4   'Photo filters
        AddMenuIcon "VIBRANCE", 5, 9, 5      'Vibrance
        AddMenuIcon "GRAYSCALE", 5, 9, 7     'Black and white
        AddMenuIcon "COLORIZE", 5, 9, 8      'Colorize
        AddMenuIcon "SEPIA", 5, 9, 9         'Sepia
    
    'Histogram
    AddMenuIcon "HISTOGRAM", 5, 10      'Histogram top-level
        AddMenuIcon "HISTOGRAM", 5, 10, 0     'Display Histogram
        AddMenuIcon "EQUALIZE", 5, 10, 2      'Equalize
        AddMenuIcon "STRETCH", 5, 10, 3       'Stretch
    
    'Invert
    AddMenuIcon "INVERT", 5, 11         'Invert top-level
        AddMenuIcon "INVCMYK", 5, 11, 0     'Invert CMYK
        AddMenuIcon "INVHUE", 5, 11, 1       'Invert Hue
        AddMenuIcon "INVRGB", 5, 11, 2       'Invert RGB
        AddMenuIcon "INVCOMPOUND", 5, 11, 4  'Compound Invert
        
    'Lighting
    AddMenuIcon "LIGHTING", 5, 12       'Lighting top-level
        AddMenuIcon "BRIGHT", 5, 12, 0       'Brightness/Contrast
        AddMenuIcon "CURVES", 5, 12, 1       'Curves
        AddMenuIcon "EXPOSURE", 5, 12, 2     'Exposure
        AddMenuIcon "GAMMA", 5, 12, 3        'Gamma Correction
        AddMenuIcon "LEVELS", 5, 12, 4       'Levels
        AddMenuIcon "SHDWHGHLGHT", 5, 12, 5  'Shadow/Highlight
        AddMenuIcon "TEMPERATURE", 5, 12, 6  'Temperature
    
    'Monochrome
    AddMenuIcon "MONOCHROME", 5, 13      'Monochrome
        AddMenuIcon "COLORTOMONO", 5, 13, 0   'Color to monochrome
        AddMenuIcon "MONOTOCOLOR", 5, 13, 1   'Monochrome to grayscale
    
    
    'Effects (Filters) Menu
    AddMenuIcon "FADELAST", 6, 0        'Fade Last
    AddMenuIcon "ARTISTIC", 6, 2        'Artistic
        '--> Artistic sub-menu
        AddMenuIcon "COMICBOOK", 6, 2, 0      'Comic book
        AddMenuIcon "FIGGLASS", 6, 2, 1       'Figured glass
        AddMenuIcon "FILMNOIR", 6, 2, 2       'Film Noir
        AddMenuIcon "KALEIDOSCOPE", 6, 2, 3   'Kaleidoscope
        AddMenuIcon "MODERNART", 6, 2, 4      'Modern Art
        AddMenuIcon "OILPAINTING", 6, 2, 5    'Oil painting
        AddMenuIcon "PENCIL", 6, 2, 6         'Pencil
        AddMenuIcon "POSTERIZE", 6, 2, 7      'Posterize
        AddMenuIcon "RELIEF", 6, 2, 8         'Relief
    AddMenuIcon "BLUR", 6, 3            'Blur
        '--> Blur sub-menu
        AddMenuIcon "BOXBLUR", 6, 3, 0        'Box Blur
        AddMenuIcon "GAUSSBLUR", 6, 3, 1      'Gaussian Blur
        AddMenuIcon "GRIDBLUR", 6, 3, 2       'Grid Blur
        AddMenuIcon "MOTIONBLUR", 6, 3, 3     'Motion Blur
        AddMenuIcon "PIXELATE", 6, 3, 4       'Pixelate (formerly Mosaic)
        AddMenuIcon "RADIALBLUR", 6, 3, 5     'Radial Blur
        AddMenuIcon "SMARTBLUR", 6, 3, 6      'Smart Blur
        AddMenuIcon "ZOOMBLUR", 6, 3, 7       'Zoom Blur
    AddMenuIcon "DISTORT", 6, 4         'Distort
        '--> Distort sub-menu
        AddMenuIcon "LENSDISTORT", 6, 4, 0    'Apply lens distortion
        AddMenuIcon "FIXLENS", 6, 4, 1        'Remove or correct existing lens distortion
        AddMenuIcon "MISCDISTORT", 6, 4, 2    'Miscellaneous distort functions
        AddMenuIcon "PANANDZOOM", 6, 4, 3     'Pan and zoom
        AddMenuIcon "PERSPECTIVE", 6, 4, 4    'Perspective (free)
        AddMenuIcon "PINCHWHIRL", 6, 4, 5     'Pinch and whirl
        AddMenuIcon "POKE", 6, 4, 6           'Poke
        AddMenuIcon "POLAR", 6, 4, 7          'Polar conversion
        AddMenuIcon "RIPPLE", 6, 4, 8         'Ripple
        AddMenuIcon "ROTATECW", 6, 4, 9       'Rotate
        AddMenuIcon "SHEAR", 6, 4, 10         'Shear
        AddMenuIcon "SPHERIZE", 6, 4, 11      'Spherize
        AddMenuIcon "SQUISH", 6, 4, 12        'Squish (formerly Fixed Perspective)
        AddMenuIcon "SWIRL", 6, 4, 13         'Swirl
        AddMenuIcon "WAVES", 6, 4, 14         'Waves
        
    AddMenuIcon "EDGES", 6, 5           'Edges
        '--> Edges sub-menu
        AddMenuIcon "EMBOSS", 6, 5, 0         'Emboss / Engrave
        AddMenuIcon "EDGEENHANCE", 6, 5, 1    'Enhance Edges
        AddMenuIcon "EDGES", 6, 5, 2          'Find Edges
        AddMenuIcon "TRACECONTOUR", 6, 5, 3   'Trace Contour
    AddMenuIcon "OTHER", 6, 6           'Fun
        '--> Fun sub-menu
        AddMenuIcon "ALIEN", 6, 6, 0          'Alien
        AddMenuIcon "BLACKLIGHT", 6, 6, 1     'Blacklight
        AddMenuIcon "DREAM", 6, 6, 2          'Dream
        AddMenuIcon "RADIOACTIVE", 6, 6, 3    'Radioactive
        AddMenuIcon "SYNTHESIZE", 6, 6, 4     'Synthesize
        AddMenuIcon "HEATMAP", 6, 6, 5        'Thermograph
        AddMenuIcon "VIBRATE", 6, 6, 6        'Vibrate
    AddMenuIcon "NATURAL", 6, 7         'Natural
        '--> Natural sub-menu
        AddMenuIcon "ATMOSPHERE", 6, 7, 0     'Atmosphere
        AddMenuIcon "BURN", 6, 7, 1           'Burn
        AddMenuIcon "FOG", 6, 7, 2            'Fog
        AddMenuIcon "FREEZE", 6, 7, 3         'Freeze
        AddMenuIcon "LAVA", 6, 7, 4           'Lava
        AddMenuIcon "RAINBOW", 6, 7, 5        'Rainbow
        AddMenuIcon "STEEL", 6, 7, 6          'Steel
        AddMenuIcon "RAIN", 6, 7, 7           'Water
    AddMenuIcon "NOISE", 6, 8           'Noise
        '--> Noise sub-menu
        AddMenuIcon "FILMGRAIN", 6, 8, 0      'Film grain
        AddMenuIcon "ADDNOISE", 6, 8, 1       'Add Noise
        AddMenuIcon "MEDIAN", 6, 8, 3         'Median
    AddMenuIcon "SHARPEN", 6, 9         'Sharpen
        '--> Sharpen sub-menu
        AddMenuIcon "SHARPEN", 6, 9, 0       'Sharpen
        AddMenuIcon "UNSHARP", 6, 9, 1       'Unsharp
    AddMenuIcon "STYLIZE", 6, 10        'Stylize
        '--> Stylize sub-menu
        AddMenuIcon "ANTIQUE", 6, 10, 0       'Antique (Sepia)
        AddMenuIcon "DIFFUSE", 6, 10, 1       'Diffuse
        AddMenuIcon "DILATE", 6, 10, 2        'Dilate
        AddMenuIcon "ERODE", 6, 10, 3         'Erode
        AddMenuIcon "SOLARIZE", 6, 10, 4      'Solarize
        AddMenuIcon "TWINS", 6, 10, 5         'Twins
        AddMenuIcon "VIGNETTE", 6, 10, 6      'Vignetting
    AddMenuIcon "CUSTFILTER", 6, 12     'Custom Filter
    
    'Tools Menu
    AddMenuIcon "LANGUAGES", 7, 0       'Languages
    AddMenuIcon "LANGEDITOR", 7, 1      'Language editor
    AddMenuIcon "RECORD", 7, 3          'Macros
        '--> Macro sub-menu
        AddMenuIcon "OPENMACRO", 7, 3, 0      'Open Macro
        AddMenuIcon "RECORD", 7, 3, 2         'Start Recording
        AddMenuIcon "RECORDSTOP", 7, 3, 3     'Stop Recording
    AddMenuIcon "PREFERENCES", 7, 5           'Options (Preferences)
    AddMenuIcon "PLUGIN", 7, 6          'Plugin Manager
    
    'Window Menu
    AddMenuIcon "NEXTIMAGE", 8, 3       'Next image
    AddMenuIcon "PREVIMAGE", 8, 4       'Previous image
    AddMenuIcon "CASCADE", 8, 6         'Cascade
    AddMenuIcon "TILEVER", 8, 7         'Tile Horizontally
    AddMenuIcon "TILEHOR", 8, 8         'Tile Vertically
    AddMenuIcon "MINALL", 8, 10         'Minimize All
    AddMenuIcon "RESTOREALL", 8, 11     'Restore All
    
    'Help Menu
    AddMenuIcon "FAVORITE", 9, 0        'Donate
    AddMenuIcon "UPDATES", 9, 2         'Check for updates
    AddMenuIcon "FEEDBACK", 9, 3        'Submit Feedback
    AddMenuIcon "BUG", 9, 4             'Submit Bug
    AddMenuIcon "PDWEBSITE", 9, 6       'Visit the PhotoDemon website
    AddMenuIcon "DOWNLOADSRC", 9, 7     'Download source code
    AddMenuIcon "LICENSE", 9, 8         'License
    AddMenuIcon "ABOUT", 9, 10          'About PD
    
    'When we're done, reset the doEvents tracker
    m_refreshOutsideProgressBar = False
    
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

    If g_OpenImageCount > 0 Then
        If Not (pdImages(g_CurrentImage).containingForm Is Nothing) Then
            If pdImages(g_CurrentImage).containingForm.WindowState = vbMaximized Then posModifier = 1
        End If
    End If
    
    'Place the icon onto the requested menu
    If subSubMenu = -1 Then
        cMenuImage.PutImageToVBMenu iconLocation, subMenu, topMenu + posModifier
    Else
        cMenuImage.PutImageToVBMenu iconLocation, subSubMenu, topMenu + posModifier, subMenu
    End If
    
    'If an outside progress bar needs to refresh, do so now
    If m_refreshOutsideProgressBar Then DoEvents

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
    
    'Vista+ gets a nice, large icon added later in the process.  XP is stuck with a 16x16 one, which we add now.
    If Not g_IsVistaOrLater Then AddMenuIcon "CLEARRECENT", 0, 1, numOfMRUFiles + 1
    
    'Change the Show/Hide panel icon to match its current state
    If g_UserPreferences.GetPref_Boolean("Core", "Hide Left Panel", False) Then
        AddMenuIcon "LEFTPANSHOW", 2, 16     'Show the panel
    Else
        AddMenuIcon "LEFTPANHIDE", 2, 16     'Hide the panel
    End If
    
    If g_UserPreferences.GetPref_Boolean("Core", "Hide Right Panel", False) Then
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
    
        If g_OpenImageCount > 0 Then
            If Not (pdImages(g_CurrentImage).containingForm Is Nothing) Then
                If pdImages(g_CurrentImage).containingForm.WindowState = vbMaximized Then posModifier = 1
            End If
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
        
        'Vista+ users now get their nice, large clear icon.
        If g_IsVistaOrLater Then
            cMRUIcons.AddImageFromStream LoadResData("CLEARRECLRG", "CUSTOM")
            cMRUIcons.PutImageToVBMenu iconLocation + 1, numOfMRUFiles + 1, 0 + posModifier, 1
        End If
                
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
    If g_UserPreferences.GetPref_Boolean("Interface", "Dynamic Taskbar Icon", True) Then icoSize = 32 Else icoSize = 16

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
    fi_DIB = FreeImage_CreateFromDC(pdImages(g_CurrentImage).mainLayer.getLayerDC)
    
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
        If g_UserPreferences.GetPref_Boolean("Interface", "Dynamic Taskbar Icon", True) Then
                
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
            setNewTaskbarIcon generatedIcon, imgForm.hWnd
            
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
        If g_UserPreferences.GetPref_Boolean("Interface", "Dynamic Taskbar Icon", True) Then
            setNewAppIcon generatedIcon
            pdImages(imgForm.Tag).curFormIcon16 = generatedIcon
        End If
        
        'Clear out memory
        FreeImage_UnloadEx finalDIB
        FreeLibrary hLib
        DeleteObject MonoBmp
        DeleteObject icoInfo.hbmColor
        
        'Use the API to assign this new icon to the specified child form
        SendMessageLong imgForm.hWnd, &H80, 0, generatedIcon
        
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
Public Function createCursorFromResource(ByVal resTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0) As Long
    
    'Start by extracting the PNG data into a bytestream
    Dim ImageData() As Byte
    ImageData() = LoadResData(resTitle, "CUSTOM")
    
    Dim IStream As IUnknown
    Dim tmpRect As RECTF
    Dim gdiBitmap As Long, hBitmap As Long
        
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
                    .xHotspot = curHotspotX
                    .yHotspot = curHotspotY
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

'Load all relevant program cursors into memory
Public Sub InitAllCursors()

    'Previously, system cursors were cached here.  This is no longer needed per https://github.com/tannerhelland/PhotoDemon/issues/78
    ' I am leaving this sub in case I need to pre-load tool cursors in the future.
    
    'Note that unloadAllCursors below is still required, as the program may dynamically generate custom cursors while running, and
    ' these cursors will not be automatically deleted by the system.

End Sub

'Unload any custom cursors from memory
Public Sub unloadAllCursors()
    
    Dim i As Long
    For i = 0 To numOfCustomCursors - 1
        DestroyCursor customCursorHandles(i)
    Next i
    
End Sub

'Use any 32bpp PNG resource as a cursor (yes, it's amazing!)
Public Sub setPNGCursorToHwnd(ByVal dstHwnd As Long, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    SetClassLong dstHwnd, GCL_HCURSOR, requestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Set a single object to use the hand cursor
Public Sub setHandCursor(ByRef tControl As Control)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub setHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

'Set a single object to use the arrow cursor
Public Sub setArrowCursorToObject(ByRef tControl As Control)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

Public Sub setArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the arrow cursor
Public Sub setArrowCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the cross cursor
Public Sub setCrossCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_CROSS)
End Sub
    
'Set a single form to use the Size All cursor
Public Sub setSizeAllCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZEALL)
End Sub

'Set a single form to use the Size NESW cursor
Public Sub setSizeNESWCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENESW)
End Sub

'Set a single form to use the Size NS cursor
Public Sub setSizeNSCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENS)
End Sub

'Set a single form to use the Size NWSE cursor
Public Sub setSizeNWSECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENWSE)
End Sub

'Set a single form to use the Size WE cursor
Public Sub setSizeWECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZEWE)
End Sub

'If a custom PNG cursor has not been loaded, this function will load the PNG, convert it to cursor format, then store
' the cursor resource for future reference (so the image doesn't have to be loaded again).
Private Function requestCustomCursor(ByVal CursorName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0) As Long

    Dim i As Long
    Dim cursorLocation As Long
    Dim cursorAlreadyLoaded As Boolean
    
    cursorLocation = 0
    cursorAlreadyLoaded = False
    
    'Loop through all cursors that have been loaded, and see if this one has been requested already.
    If numOfCustomCursors > 0 Then
    
        For i = 0 To numOfCustomCursors - 1
        
            If customCursorNames(i) = CursorName Then
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
        tmpHandle = createCursorFromResource(CursorName, cursorHotspotX, cursorHotspotY)
        
        ReDim Preserve customCursorNames(0 To numOfCustomCursors) As String
        ReDim Preserve customCursorHandles(0 To numOfCustomCursors) As Long
        
        customCursorNames(numOfCustomCursors) = CursorName
        customCursorHandles(numOfCustomCursors) = tmpHandle
        
        numOfCustomCursors = numOfCustomCursors + 1
        
        requestCustomCursor = tmpHandle
    End If

End Function

'Given an image in the .exe's resource section (typically a PNG image), load it to a pdLayer object.
' The calling function is responsible for deleting the layer once they are done with it.
Public Function loadResourceToLayer(ByVal resTitle As String, ByRef dstLayer As pdLayer, Optional ByVal vbSupportedFormat As Boolean = False) As Boolean
    
    'If the requested image is in a VB-compatible format (e.g. BMP), we don't need to use GDI+
    If vbSupportedFormat Then
    
        'Load the requested image into a temporary StdPicture object
        Dim tmppic As StdPicture
        Set tmppic = New StdPicture
        Set tmppic = LoadResPicture(resTitle, 0)
        
        'Copy that image into the supplied layer
        If dstLayer.CreateFromPicture(tmppic) Then
            loadResourceToLayer = True
        Else
            loadResourceToLayer = False
        End If
        
        Exit Function
        
    Else
    
        'Start by extracting the PNG data into a bytestream
        Dim ImageData() As Byte
        ImageData() = LoadResData(resTitle, "CUSTOM")
        
        Dim IStream As IUnknown
        Dim tmpRect As RECTF
        Dim gdiBitmap As Long, hBitmap As Long
            
        CreateStreamOnHGlobal ImageData(0), 0&, IStream
        
        If Not IStream Is Nothing Then
            
            'Use GDI+ to convert the bytestream into a usable image
            ' (Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine)
            If GdipLoadImageFromStream(IStream, gdiBitmap) = 0 Then
            
                'Retrieve the image's size and pixel format
                GdipGetImageBounds gdiBitmap, tmpRect, UnitPixel
                
                Dim gdiPixelFormat As Long
                GdipGetImagePixelFormat gdiBitmap, gdiPixelFormat
                
                'If the image has an alpha channel, create a 32bpp layer to receive it
                If (gdiPixelFormat And PixelFormatAlpha <> 0) Or (gdiPixelFormat And PixelFormatPAlpha <> 0) Then
                    dstLayer.createBlank tmpRect.fWidth, tmpRect.fHeight, 32
                Else
                    dstLayer.createBlank tmpRect.fWidth, tmpRect.fHeight, 24
                End If
                
                'Convert the GDI+ bitmap to a standard Windows hBitmap
                If GdipCreateHBITMAPFromBitmap(gdiBitmap, hBitmap, vbBlack) = 0 Then
                
                    'Select the hBitmap into a new DC so we can BitBlt it into the layer
                    Dim gdiDC As Long
                    gdiDC = CreateCompatibleDC(0)
                    SelectObject gdiDC, hBitmap
                    
                    'Copy the GDI+ bitmap into the layer
                    BitBlt dstLayer.getLayerDC, 0, 0, tmpRect.fWidth, tmpRect.fHeight, gdiDC, 0, 0, vbSrcCopy
                    
                    'Verify the alpha channel
                    If Not dstLayer.verifyAlphaChannel Then dstLayer.convertTo24bpp
                    
                    'Release the Windows-format bitmap and temporary device context
                    DeleteObject hBitmap
                    DeleteDC gdiDC
                    
                    'Release the GDI+ bitmap as well
                    GdipDisposeImage gdiBitmap
                    
                    'Free the memory stream
                    Set IStream = Nothing
                    
                    loadResourceToLayer = True
                    Exit Function
                
                End If
                
                'Release the GDI+ bitmap and mark the load as failed
                GdipDisposeImage gdiBitmap
                loadResourceToLayer = False
                Exit Function
                    
            End If
        
            'Free the memory stream
            Set IStream = Nothing
            loadResourceToLayer = False
            Exit Function
        
        End If
        
        loadResourceToLayer = False
        Exit Function
    
    End If
        
End Function
