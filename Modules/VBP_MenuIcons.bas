Attribute VB_Name = "Icon_and_Cursor_Handler"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright ©2012-2014 by Tanner Helland
'Created: 24/June/12
'Last updated: 13/October/13
'Last update: rework custom form icon code to be much cleaner and self-contained
'
'Because VB6 doesn't provide many mechanisms for working with icons, I've had to manually add a number of icon-related
' functions to PhotoDemon.  First is a way to add icons/bitmaps to menus, as originally written by Leandro Ascierto.
' Menu icons are extracted from a resource file (where they're stored in PNG format) and rendered to the menu at run-time.
' See the clsMenuImage class for details on how this works. (A link to Leandro's original project can also be found there.)
'
'This module also handles the rendering of dynamic form, program, and taskbar icons.  (When an image is loaded and active,
' those icons can change to match the current image.)  As of February 2013, custom form icon generation has now been reworked
' based off this MSDN article: http://support.microsoft.com/kb/318876
' The new code is much leaner (and cleaner!) than past incarnations.
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
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hbmReturn As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hBmpReturn As Long, ByVal Background As Long) As GDIPlusStatus
Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal gdiBitmap As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef PixelFormat As Long) As Long

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

Public Enum SystemCursorConstant
    IDC_DEFAULT = 0&
    IDC_APPSTARTING = 32650&
    IDC_HAND = 32649&
    IDC_ARROW = 32512&
    IDC_CROSS = 32515&
    IDC_IBEAM = 32513&
    IDC_ICON = 32641&
    IDC_NO = 32648&
    IDC_SIZEALL = 32646&
    IDC_SIZENESW = 32643&
    IDC_SIZENS = 32645&
    IDC_SIZENWSE = 32642&
    IDC_SIZEWE = 32644&
    IDC_UPARROW = 32516&
    IDC_WAIT = 32514&
End Enum

#If False Then
    Private Const IDC_DEFAULT = 0&, IDC_APPSTARTING = 32650&, IDC_HAND = 32649&, IDC_ARROW = 32512&, IDC_CROSS = 32515&, IDC_IBEAM = 32513&, IDC_ICON = 32641&, IDC_NO = 32648&, IDC_SIZEALL = 32646&, IDC_SIZENESW = 32643&, IDC_SIZENS = 32645&, IDC_SIZENWSE = 32642&, IDC_SIZEWE = 32644&, IDC_UPARROW = 32516&, IDC_WAIT = 32514&
#End If

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
Public Sub loadMenuIcons()

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
        If (Not g_IsVistaOrLater) And (Not g_IsProgramCompiled) Then
            Debug.Print "XP + IDE detected.  Menu icons will be disabled for this session."
            Exit Sub
        End If
        
        .Init FormMain.hWnd, fixDPI(16), fixDPI(16)
        
    End With
            
    'Now that all menu icons are loaded, apply them to the proper menu entires
    applyAllMenuIcons
        
    '...and initialize the separate MRU icon handler.
    Set cMRUIcons = New clsMenuImage
    If g_IsVistaOrLater Then
        cMRUIcons.Init FormMain.hWnd, fixDPI(64), fixDPI(64)
    Else
        cMRUIcons.Init FormMain.hWnd, fixDPI(16), fixDPI(16)
    End If
        
End Sub

'Apply (and if necessary, dynamically load) menu icons to their proper menu entries.
Public Sub applyAllMenuIcons(Optional ByVal useDoEvents As Boolean = False)

    m_refreshOutsideProgressBar = useDoEvents

    'Load every icon from the resource file.  (Yes, there are a LOT of icons!)
        
    'File Menu
    addMenuIcon "OPENIMG", 0, 0       'Open Image
    addMenuIcon "OPENREC", 0, 1       'Open recent
    addMenuIcon "IMPORT", 0, 2        'Import
    addMenuIcon "CLOSE", 0, 4         'Close
    addMenuIcon "CLOSE", 0, 5         'Close All
    addMenuIcon "SAVE", 0, 7          'Save
    addMenuIcon "SAVEAS", 0, 8        'Save As...
    addMenuIcon "REVERT", 0, 9        'Revert
    addMenuIcon "BCONVERT", 0, 11     'Batch conversion
    addMenuIcon "PRINT", 0, 13        'Print
    addMenuIcon "EXIT", 0, 15         'Exit
    
    '--> Import Sub-Menu
    'NOTE: the specific menu values will be different if the scanner plugin (eztw32.dll) isn't found.
    If g_ScanEnabled Then
        addMenuIcon "PASTE_IMAGE", 0, 2, 0 'From Clipboard (Paste as New Image)
        addMenuIcon "SCANNER", 0, 2, 2     'Scan Image
        addMenuIcon "SCANNERSEL", 0, 2, 3  'Select Scanner
        addMenuIcon "DOWNLOAD", 0, 2, 5    'Online Image
        addMenuIcon "SCREENCAP", 0, 2, 7   'Screen Capture
    Else
        addMenuIcon "PASTE_IMAGE", 0, 2, 0 'From Clipboard (Paste as New Image)
        addMenuIcon "DOWNLOAD", 0, 2, 2    'Online Image
        addMenuIcon "SCREENCAP", 0, 2, 4   'Screen Capture
    End If
        
    'Edit Menu
    addMenuIcon "UNDO", 1, 0           'Undo
    addMenuIcon "REDO", 1, 1           'Redo
    addMenuIcon "REPEAT", 1, 2         'Repeat Last Action
    addMenuIcon "COPY", 1, 4           'Copy
    addMenuIcon "COPY_MERGED", 1, 5    'Copy merged
    addMenuIcon "PASTE_LAYER", 1, 6    'Paste as new layer
    addMenuIcon "PASTE_IMAGE", 1, 7    'Paste as new image
    addMenuIcon "CLEAR", 1, 9          'Empty Clipboard
    
    'View Menu
    addMenuIcon "FITONSCREEN", 2, 0    'Fit on Screen
    addMenuIcon "ZOOMIN", 2, 2         'Zoom In
    addMenuIcon "ZOOMOUT", 2, 3        'Zoom Out
    addMenuIcon "ZOOMACTUAL", 2, 9     'Zoom 100%
    
    'Image Menu
    addMenuIcon "DUPLICATE", 3, 0      'Duplicate
    addMenuIcon "RESIZE", 3, 2         'Resize
    addMenuIcon "CANVASSIZE", 3, 4     'Canvas resize
    'addMenuIcon "CANVASSIZE", 3, 5     'Fit canvas to active layer
    'addMenuIcon "CANVASSIZE", 3, 6     'Fit canvas around all layers
    addMenuIcon "CROPSEL", 3, 8        'Crop to Selection
    addMenuIcon "AUTOCROP", 3, 9      'Trim
    addMenuIcon "ROTATECW", 3, 11     'Rotate top-level
        '--> Rotate sub-menu
        addMenuIcon "STRAIGHTEN", 3, 11, 0  'Straighten
        addMenuIcon "ROTATECW", 3, 11, 2    'Rotate Clockwise
        addMenuIcon "ROTATECCW", 3, 11, 3   'Rotate Counter-clockwise
        addMenuIcon "ROTATE180", 3, 11, 4   'Rotate 180
        If g_ImageFormats.FreeImageEnabled Then addMenuIcon "ROTATEANY", 3, 11, 5  'Rotate Arbitrary
    addMenuIcon "MIRROR", 3, 12        'Mirror
    addMenuIcon "FLIP", 3, 13          'Flip
    'addMenuIcon "ISOMETRIC", 3, 12     'Isometric      'NOTE: isometric was removed in v6.4.
    addMenuIcon "REDUCECOLORS", 3, 15  'Indexed color (Reduce Colors)
    If g_ImageFormats.FreeImageEnabled Then FormMain.MnuImage(15).Enabled = True Else FormMain.MnuImage(15).Enabled = False
    addMenuIcon "TILE", 3, 16          'Tile
    addMenuIcon "METADATA", 3, 18      'Metadata (top-level)
        '--> Metadata sub-menu
        addMenuIcon "BROWSEMD", 3, 18, 0     'Browse metadata
        addMenuIcon "COUNTCOLORS", 3, 18, 2  'Count Colors
        addMenuIcon "MAPPHOTO", 3, 18, 3     'Map photo location
    
    'Layer menu
    addMenuIcon "ADDLAYER", 4, 0        'Add layer (top-level)
        '--> Add layer sub-menu
        addMenuIcon "ADDLAYER", 4, 0, 0             'Add blank layer
        addMenuIcon "DUPL_LAYER", 4, 0, 1          'Add duplicate layer
        addMenuIcon "PASTE_LAYER", 4, 0, 3          'Add layer from clipboard
        addMenuIcon "ADDLAYERFILE", 4, 0, 4             'Add layer from file
    addMenuIcon "DELLAYER", 4, 1        'Delete layer (top-level)
        '--> Delete layer sub-menu
        addMenuIcon "DELLAYER", 4, 1, 0       'Delete current layer
        addMenuIcon "DELLAYERHDN", 4, 1, 1       'Delete all hidden layers
    addMenuIcon "MERGE_UP", 4, 3         'Merge up
    addMenuIcon "MERGE_DOWN", 4, 4       'Merge down
    addMenuIcon "LAYERORDER", 4, 5      'Order (top-level)
        '--> Order layer sub-menu
        addMenuIcon "LAYERUP", 4, 5, 0     'Raise layer
        addMenuIcon "LAYERDOWN", 4, 5, 1     'Lower layer
        addMenuIcon "LAYERTOTOP", 4, 5, 3     'Raise to top
        addMenuIcon "LAYERTOBTM", 4, 5, 4     'Lower to bottom
    addMenuIcon "ROTATECW", 4, 7         'Layer Orientation (top-level)
        '--> Orientation sub-menu
        addMenuIcon "STRAIGHTEN", 4, 7, 0   'Straighten
        addMenuIcon "ROTATECW", 4, 7, 2     'Rotate Clockwise
        addMenuIcon "ROTATECCW", 4, 7, 3    'Rotate Counter-clockwise
        addMenuIcon "ROTATE180", 4, 7, 4    'Rotate 180
        If g_ImageFormats.FreeImageEnabled Then addMenuIcon "ROTATEANY", 4, 7, 5   'Rotate Arbitrary
        addMenuIcon "MIRROR", 4, 7, 7       'Mirror
        addMenuIcon "FLIP", 4, 7, 8         'Flip
    addMenuIcon "RESIZE", 4, 8           'Layer Size (top-level)
        '--> Size sub-menu
        addMenuIcon "RESETSIZE", 4, 8, 0        'Reset to original size
        addMenuIcon "RESIZE", 4, 8, 2        'Resize
        addMenuIcon "SMRTRESIZE", 4, 8, 3    'Content-aware resize
    addMenuIcon "TRANSPARENCY", 4, 10    'Layer Transparency
        '--> Transparency sub-menu
        addMenuIcon "ADDTRANS", 4, 10, 0     'Add alpha channel
        addMenuIcon "GREENSCREEN", 4, 10, 1  'Color to alpha
        addMenuIcon "REMOVETRANS", 4, 10, 3  'Remove alpha channel
    addMenuIcon "FLATTEN", 4, 12         'Flatten image
    addMenuIcon "MERGEVISIBLE", 4, 13   'Merge visible layers
    
    'Select Menu
    addMenuIcon "SELECTALL", 5, 0       'Select all
    addMenuIcon "SELECTNONE", 5, 1      'Select none
    addMenuIcon "SELECTINVERT", 5, 2    'Invert selection
    addMenuIcon "SELECTGROW", 5, 4      'Grow selection
    addMenuIcon "SELECTSHRINK", 5, 5    'Shrink selection
    addMenuIcon "SELECTBORDER", 5, 6    'Border selection
    addMenuIcon "SELECTFTHR", 5, 7      'Feather selection
    addMenuIcon "SELECTSHRP", 5, 8      'Sharpen selection
    addMenuIcon "SELECTLOAD", 5, 10     'Load selection from file
    addMenuIcon "SELECTSAVE", 5, 11     'Save selection to file
    addMenuIcon "SELECTEXPORT", 5, 12   'Export selection (top-level)
        '--> Export Selection sub-menu
        addMenuIcon "EXPRTSELAREA", 5, 12, 0  'Export selected area as image
        addMenuIcon "EXPRTSELMASK", 5, 12, 1  'Export selection mask as image
    
    'Adjustments Menu
    
    'Adjustment shortcuts (top-level menu items)
    addMenuIcon "GRAYSCALE", 6, 0       'Black and white
    addMenuIcon "BRIGHT", 6, 1          'Brightness/Contrast
    addMenuIcon "COLORBALANCE", 6, 2    'Color balance
    addMenuIcon "CURVES", 6, 3          'Curves
    addMenuIcon "LEVELS", 6, 4          'Levels
    addMenuIcon "VIBRANCE", 6, 5        'Vibrance
    addMenuIcon "WHITEBAL", 6, 6        'White Balance
       
    'Channels
    addMenuIcon "CHANNELMIX", 6, 8     'Channels top-level
        addMenuIcon "CHANNELMIX", 6, 8, 0    'Channel mixer
        addMenuIcon "RECHANNEL", 6, 8, 1     'Rechannel
        addMenuIcon "CHANNELMAX", 6, 8, 3    'Channel max
        addMenuIcon "CHANNELMIN", 6, 8, 4    'Channel min
        addMenuIcon "COLORSHIFTL", 6, 8, 6   'Shift Left
        addMenuIcon "COLORSHIFTR", 6, 8, 7   'Shift Right
            
    'Color
    addMenuIcon "HSL", 6, 9            'Color balance
        addMenuIcon "COLORBALANCE", 6, 9, 0  'Color balance
        addMenuIcon "WHITEBAL", 6, 9, 1      'White Balance
        addMenuIcon "HSL", 6, 9, 3           'HSL adjustment
        addMenuIcon "VIBRANCE", 6, 9, 4      'Vibrance
        addMenuIcon "GRAYSCALE", 6, 9, 6     'Black and white
        addMenuIcon "COLORIZE", 6, 9, 7      'Colorize
        addMenuIcon "REPLACECLR", 6, 9, 8    'Replace color
        addMenuIcon "SEPIA", 6, 9, 9         'Sepia
    
    'Histogram
    addMenuIcon "HISTOGRAM", 6, 10      'Histogram top-level
        addMenuIcon "HISTOGRAM", 6, 10, 0     'Display Histogram
        addMenuIcon "EQUALIZE", 6, 10, 2      'Equalize
        addMenuIcon "STRETCH", 6, 10, 3       'Stretch
    
    'Invert
    addMenuIcon "INVERT", 6, 11         'Invert top-level
        addMenuIcon "INVCMYK", 6, 11, 0     'Invert CMYK
        addMenuIcon "INVHUE", 6, 11, 1       'Invert Hue
        addMenuIcon "INVRGB", 6, 11, 2       'Invert RGB
        addMenuIcon "INVCOMPOUND", 6, 11, 4  'Compound Invert
        
    'Lighting
    addMenuIcon "LIGHTING", 6, 12       'Lighting top-level
        addMenuIcon "BRIGHT", 6, 12, 0       'Brightness/Contrast
        addMenuIcon "CURVES", 6, 12, 1       'Curves
        addMenuIcon "GAMMA", 6, 12, 2        'Gamma Correction
        addMenuIcon "LEVELS", 6, 12, 3       'Levels
        addMenuIcon "SHDWHGHLGHT", 6, 12, 4  'Shadow/Highlight
        addMenuIcon "TEMPERATURE", 6, 12, 5  'Temperature
    
    'Monochrome
    addMenuIcon "MONOCHROME", 6, 13      'Monochrome
        addMenuIcon "COLORTOMONO", 6, 13, 0   'Color to monochrome
        addMenuIcon "MONOTOCOLOR", 6, 13, 1   'Monochrome to grayscale
        
    'Photography
    addMenuIcon "PHOTOFILTER", 6, 14      'Photography top-level
        addMenuIcon "EXPOSURE", 6, 14, 0     'Exposure
        addMenuIcon "PHOTOFILTER", 6, 14, 1  'Photo filters
        addMenuIcon "SPLITTONE", 6, 14, 2    'Split-toning
    
    
    'Effects (Filters) Menu
    'addMenuIcon "FADELAST", 7, 0        'Fade Last
    addMenuIcon "ARTISTIC", 7, 2        'Artistic
        '--> Artistic sub-menu
        addMenuIcon "COMICBOOK", 7, 2, 0      'Comic book
        addMenuIcon "FIGGLASS", 7, 2, 1       'Figured glass
        addMenuIcon "FILMNOIR", 7, 2, 2       'Film Noir
        addMenuIcon "GLASSTILES", 7, 2, 3     'Glass tiles
        addMenuIcon "KALEIDOSCOPE", 7, 2, 4   'Kaleidoscope
        addMenuIcon "MODERNART", 7, 2, 5      'Modern Art
        addMenuIcon "OILPAINTING", 7, 2, 6    'Oil painting
        addMenuIcon "PENCIL", 7, 2, 7         'Pencil
        addMenuIcon "POSTERIZE", 7, 2, 8      'Posterize
        addMenuIcon "RELIEF", 7, 2, 9         'Relief
    addMenuIcon "BLUR", 7, 3            'Blur
        '--> Blur sub-menu
        addMenuIcon "BOXBLUR", 7, 3, 0        'Box Blur
        addMenuIcon "GAUSSBLUR", 7, 3, 1      'Gaussian Blur
        addMenuIcon "SMARTBLUR", 7, 3, 2      'Surface Blur (formerly Smart Blur)
        addMenuIcon "MOTIONBLUR", 7, 3, 4     'Motion Blur
        addMenuIcon "RADIALBLUR", 7, 3, 5     'Radial Blur
        addMenuIcon "ZOOMBLUR", 7, 3, 6       'Zoom Blur
        addMenuIcon "CHROMABLUR", 7, 3, 8     'Chroma Blur
        addMenuIcon "FRAGMENT", 7, 3, 9       'Fragment
        addMenuIcon "GRIDBLUR", 7, 3, 10      'Grid Blur
        addMenuIcon "PIXELATE", 7, 3, 11      'Pixelate (formerly Mosaic)
        
    addMenuIcon "DISTORT", 7, 4         'Distort
        '--> Distort sub-menu
        addMenuIcon "LENSDISTORT", 7, 4, 0    'Apply lens distortion
        addMenuIcon "FIXLENS", 7, 4, 1        'Remove or correct existing lens distortion
        addMenuIcon "MISCDISTORT", 7, 4, 2    'Miscellaneous distort functions
        addMenuIcon "PANANDZOOM", 7, 4, 3     'Pan and zoom
        addMenuIcon "PERSPECTIVE", 7, 4, 4    'Perspective (free)
        addMenuIcon "PINCHWHIRL", 7, 4, 5     'Pinch and whirl
        addMenuIcon "POKE", 7, 4, 6           'Poke
        addMenuIcon "POLAR", 7, 4, 7          'Polar conversion
        addMenuIcon "RIPPLE", 7, 4, 8         'Ripple
        addMenuIcon "ROTATECW", 7, 4, 9       'Rotate
        addMenuIcon "SHEAR", 7, 4, 10         'Shear
        addMenuIcon "SPHERIZE", 7, 4, 11      'Spherize
        addMenuIcon "SQUISH", 7, 4, 12        'Squish (formerly Fixed Perspective)
        addMenuIcon "SWIRL", 7, 4, 13         'Swirl
        addMenuIcon "WAVES", 7, 4, 14         'Waves
        
    addMenuIcon "EDGES", 7, 5           'Edges
        '--> Edges sub-menu
        addMenuIcon "EMBOSS", 7, 5, 0         'Emboss / Engrave
        addMenuIcon "EDGEENHANCE", 7, 5, 1    'Enhance Edges
        addMenuIcon "EDGES", 7, 5, 2          'Find Edges
        addMenuIcon "TRACECONTOUR", 7, 5, 3   'Trace Contour
        
    addMenuIcon "SUNSHINE", 7, 6        'Lights and shadows
        '--> Lights and shadows sub-menu
        addMenuIcon "BLACKLIGHT", 7, 6, 0     'Blacklight
        addMenuIcon "DILATE", 7, 6, 1         'Dilate
        addMenuIcon "ERODE", 7, 6, 2          'Erode
        addMenuIcon "RAINBOW", 7, 6, 3        'Rainbow
        addMenuIcon "SUNSHINE", 7, 6, 4       'Sunshine
    
    addMenuIcon "NATURAL", 7, 7         'Natural
        '--> Natural sub-menu
        addMenuIcon "ATMOSPHERE", 7, 7, 0     'Atmosphere
        addMenuIcon "BURN", 7, 7, 1           'Burn
        addMenuIcon "FOG", 7, 7, 2            'Fog
        addMenuIcon "FREEZE", 7, 7, 3         'Freeze
        addMenuIcon "LAVA", 7, 7, 4           'Lava
        addMenuIcon "STEEL", 7, 7, 5          'Steel
        addMenuIcon "RAIN", 7, 7, 6           'Water
        
    addMenuIcon "NOISE", 7, 8           'Noise
        '--> Noise sub-menu
        addMenuIcon "FILMGRAIN", 7, 8, 0      'Film grain
        addMenuIcon "ADDNOISE", 7, 8, 1       'Add Noise
        addMenuIcon "MEDIAN", 7, 8, 3         'Median
        
    addMenuIcon "SHARPEN", 7, 9         'Sharpen
        '--> Sharpen sub-menu
        addMenuIcon "SHARPEN", 7, 9, 0       'Sharpen
        addMenuIcon "UNSHARP", 7, 9, 1       'Unsharp
        
    addMenuIcon "STYLIZE", 7, 10        'Stylize
        '--> Stylize sub-menu
        addMenuIcon "ANTIQUE", 7, 10, 0       'Antique (Sepia)
        addMenuIcon "DIFFUSE", 7, 10, 1       'Diffuse
        addMenuIcon "SOLARIZE", 7, 10, 2      'Solarize
        addMenuIcon "TWINS", 7, 10, 3         'Twins
        addMenuIcon "VIGNETTE", 7, 10, 4      'Vignetting
    addMenuIcon "CUSTFILTER", 7, 12     'Custom Filter
    
    addMenuIcon "OTHER", 7, 14           'Experimental
        '--> Experimental sub-menu
        addMenuIcon "ALIEN", 7, 14, 0          'Alien
        addMenuIcon "DREAM", 7, 14, 1          'Dream
        addMenuIcon "RADIOACTIVE", 7, 14, 2    'Radioactive
        addMenuIcon "SYNTHESIZE", 7, 14, 3     'Synthesize
        addMenuIcon "HEATMAP", 7, 14, 4        'Thermograph
        addMenuIcon "VIBRATE", 7, 14, 5        'Vibrate
    
    'Tools Menu
    addMenuIcon "LANGUAGES", 8, 0       'Languages
    addMenuIcon "LANGEDITOR", 8, 1      'Language editor
    addMenuIcon "RECORD", 8, 3          'Macros
        '--> Macro sub-menu
        addMenuIcon "OPENMACRO", 8, 3, 0      'Open Macro
        addMenuIcon "RECORD", 8, 3, 2         'Start Recording
        addMenuIcon "RECORDSTOP", 8, 3, 3     'Stop Recording
    addMenuIcon "PREFERENCES", 8, 5           'Options (Preferences)
    addMenuIcon "PLUGIN", 8, 6          'Plugin Manager
    
    'Window Menu
    addMenuIcon "NEXTIMAGE", 9, 7       'Next image
    addMenuIcon "PREVIMAGE", 9, 8       'Previous image
    
    'Help Menu
    addMenuIcon "FAVORITE", 10, 0        'Donate
    addMenuIcon "UPDATES", 10, 2         'Check for updates
    addMenuIcon "FEEDBACK", 10, 3        'Submit Feedback
    addMenuIcon "BUG", 10, 4             'Submit Bug
    addMenuIcon "PDWEBSITE", 10, 6       'Visit the PhotoDemon website
    addMenuIcon "DOWNLOADSRC", 10, 7     'Download source code
    addMenuIcon "LICENSE", 10, 8         'License
    addMenuIcon "ABOUT", 10, 10          'About PD
    
    'When we're done, reset the doEvents tracker
    m_refreshOutsideProgressBar = False
    
End Sub

'This new, simpler technique for adding menu icons requires only the menu location (including sub-menus) and the icon's identifer
' in the resource file.  If the icon has already been loaded, it won't be loaded again; instead, the function will check the list
' of loaded icons and automatically fill in the numeric identifier as necessary.
Private Sub addMenuIcon(ByVal resID As String, ByVal topMenu As Long, ByVal subMenu As Long, Optional ByVal subSubMenu As Long = -1)

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
        
    'Place the icon onto the requested menu
    If subSubMenu = -1 Then
        cMenuImage.PutImageToVBMenu iconLocation, subMenu, topMenu
    Else
        cMenuImage.PutImageToVBMenu iconLocation, subSubMenu, topMenu, subMenu
    End If
    
    'If an outside progress bar needs to refresh, do so now
    If m_refreshOutsideProgressBar Then DoEvents

End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' Note that to keep the code small, all changeable icons are refreshed whenever this is called.
Public Sub resetMenuIcons()
        
    'Redraw the Undo/Redo menus
    addMenuIcon "UNDO", 1, 0     'Undo
    addMenuIcon "REDO", 1, 1     'Redo
    
    'Redraw the Window menu, as some of its menus will be en/disabled according to the docking status of image windows
    addMenuIcon "NEXTIMAGE", 9, 7       'Next image
    addMenuIcon "PREVIMAGE", 9, 8       'Previous image
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    Dim numOfMRUFiles As Long
    numOfMRUFiles = g_RecentFiles.MRU_ReturnCount()
    
    'Vista+ gets nice, large icons added later in the process.  XP is stuck with 16x16 ones, which we add now.
    If Not g_IsVistaOrLater Then
        addMenuIcon "LOADALL", 0, 1, numOfMRUFiles + 1
        addMenuIcon "CLEARRECENT", 0, 1, numOfMRUFiles + 2
    End If
    
    'Clear the current MRU icon cache.
    ' (Note added 01 Jan 2014 - RR has reported an IDE error on the following line, which means this function is somehow being
    '  triggered before loadMenuIcons above.  I cannot reproduce this behavior, so instead, we now perform a single initialization
    '  check before attempting to load MRU icons.)
    If Not cMRUIcons Is Nothing Then
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
            tmpFilename = g_RecentFiles.getMRUThumbnailPath(i)
            
            'If the file exists, add it to the MRU icon handler
            If FileExist(tmpFilename) Then
                    
                iconLocation = iconLocation + 1
                cMRUIcons.AddImageFromFile tmpFilename
                cMRUIcons.PutImageToVBMenu iconLocation, i, 0, 1
            
            'If a thumbnail for this file does not exist, supply a placeholder image (Vista+ only; on XP it will simply be blank)
            Else
                If g_IsVistaOrLater Then cMRUIcons.PutImageToVBMenu 0, i, 0, 1
            End If
            
        Next i
            
        'Vista+ users now get their nice, large "load all recent files" and "clear list" icons.
        If g_IsVistaOrLater Then
            cMRUIcons.AddImageFromStream LoadResData("LOADALLLRG", "CUSTOM")
            cMRUIcons.PutImageToVBMenu iconLocation + 1, numOfMRUFiles + 1, 0, 1
            
            cMRUIcons.AddImageFromStream LoadResData("CLEARRECLRG", "CUSTOM")
            cMRUIcons.PutImageToVBMenu iconLocation + 2, numOfMRUFiles + 2, 0, 1
        End If
        
    End If
        
End Sub

'Convert a DIB - any DIB! - to an icon via CreateIconIndirect.  Transparency will be preserved, and by default, the icon will be created
' at the current image's size (though you can specify a custom size if you wish).  Ideally, the passed DIB will have been created using
' the pdImage function "requestThumbnail".
'FreeImage is currently required for this function, because it provides a simple way to move between DIBs and DDBs.  I could rewrite
' the function without FreeImage's help, but frankly don't consider it worth the trouble.
Public Function getIconFromDIB(ByRef srcDIB As pdDIB, Optional iconSize As Long = 0) As Long

    If Not g_ImageFormats.FreeImageEnabled Then
        getIconFromDIB = 0
        Exit Function
    End If
    
    Dim fi_DIB As Long
    fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
    
    'If the iconSize parameter is 0, use the current DIB's dimensions.  Otherwise, resize it as requested.
    If iconSize = 0 Then
        iconSize = srcDIB.getDIBWidth
    Else
        fi_DIB = FreeImage_RescaleByPixel(fi_DIB, iconSize, iconSize, True, FILTER_BILINEAR)
    End If
    
    If fi_DIB <> 0 Then
    
        'Icon generation has a number of quirks.  One is that even if you want a 32bpp icon, you still must supply a blank
        ' monochrome mask for the icon, even though the API just discards it.  Prepare such a mask now.
        Dim monoBmp As Long
        monoBmp = CreateBitmap(iconSize, iconSize, 1, 1, ByVal 0&)
        
        'Create a header for the icon we desire, then use CreateIconIndirect to create it.
        Dim icoInfo As ICONINFO
        With icoInfo
            .fIcon = True
            .xHotspot = iconSize
            .yHotspot = iconSize
            .hbmMask = monoBmp
            .hbmColor = FreeImage_GetBitmapForDevice(fi_DIB)
        End With
        
        getIconFromDIB = CreateIconIndirect(icoInfo)
        
        'Delete the temporary monochrome mask and DDB
        DeleteObject monoBmp
        DeleteObject icoInfo.hbmColor
    
    Else
        getIconFromDIB = 0
    End If
    
    'Release FreeImage's copy of the source DIB
    FreeImage_UnloadEx fi_DIB
    
End Function

'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
' Note that this function currently requires the FreeImage plugin to be present on the system.
Public Sub createCustomFormIcon(ByRef srcImage As pdImage)

    If Not ALLOW_DYNAMIC_ICONS Then Exit Sub
    If Not g_ImageFormats.FreeImageEnabled Then Exit Sub
    If srcImage Is Nothing Then Exit Sub

    'Taskbar icons are generally 32x32.  Form titlebar icons are generally 16x16.
    Dim hIcon32 As Long, hIcon16 As Long

    Dim thumbDIB As pdDIB
    Set thumbDIB = New pdDIB

    'Request a 32x32 thumbnail version of the current image
    If srcImage.requestThumbnail(thumbDIB, 32) Then

        'Request an icon-format version of the generated thumbnail
        hIcon32 = getIconFromDIB(thumbDIB)

        'Assign the new icon to the taskbar
        'setNewTaskbarIcon hIcon32, imgForm.hWnd

        '...and remember it in our current icon collection
        addIconToList hIcon32

        '...and the current form
        srcImage.curFormIcon32 = hIcon32

        'Now repeat the same steps, but for a 16x16 icon to be used in the form's title bar.
        hIcon16 = getIconFromDIB(thumbDIB, 16)
        addIconToList hIcon16
        srcImage.curFormIcon16 = hIcon16

        'Apply the 16x16 icon to the title bar of the specified form
        'SendMessageLong imgForm.hWnd, &H80, 0, hIcon16

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
    
    'Start by extracting the PNG resource data into a pdLayer object.
    Dim resDIB As pdDIB
    Set resDIB = New pdDIB
                
    If loadResourceToDIB(resTitle, resDIB) Then
    
        'If the user is running at non-96 DPI, we now need to perform DPI correction on the cursor.
        ' NOTE: this line will need to be revisited if high-DPI cursors are added to the resource file.  The reason
        '       I'm not making the change now is because PD's current cursor are not implemented uniformly, so I
        '       need to standardize their size and layout before implementing a universal "resize per DPI" check.
        '       The proper way to do this would be to retrieve cursor size from the system, then resize anything
        '       that isn't already that size - I've made a note to do this soon.
        If fixDPI(96) <> 96 Then
        
            'Create a temporary copy of the image
            Dim dpiDIB As pdDIB
            Set dpiDIB = New pdDIB
            
            dpiDIB.createFromExistingDIB resDIB
            
            'Erase and resize the primary DIB
            resDIB.createBlank fixDPI(dpiDIB.getDIBWidth), fixDPI(dpiDIB.getDIBHeight), dpiDIB.getDIBColorDepth
            
            'Use GDI+ to resize the cursor from dpiDIB into resDIB
            GDIPlusResizeDIB resDIB, 0, 0, resDIB.getDIBWidth, resDIB.getDIBHeight, dpiDIB, 0, 0, dpiDIB.getDIBWidth, dpiDIB.getDIBHeight, InterpolationModeNearestNeighbor
        
            'Release our temporary DIB
            Set dpiDIB = Nothing
        
        End If
        
        'Generate a blank monochrome mask to pass to the icon creation function.
        ' (This is a stupid Windows workaround for 32bpp cursors.  The cursor creation function always assumes
        '  the presence of a mask bitmap, so we have to submit one even if we want the PNG's alpha channel
        '  used for transparency!)
        Dim monoBmp As Long
        monoBmp = CreateBitmap(resDIB.getDIBWidth, resDIB.getDIBHeight, 1, 1, ByVal 0&)
        
        'Create an icon header and point it at our temp mask bitmap and our PNG resource
        Dim icoInfo As ICONINFO
        With icoInfo
            .fIcon = False
            .xHotspot = fixDPI(curHotspotX)
            .yHotspot = fixDPI(curHotspotY)
            .hbmMask = monoBmp
            .hbmColor = resDIB.getDIBHandle
        End With
                    
        'Create the cursor
        createCursorFromResource = CreateIconIndirect(icoInfo)
        
        'Release our temporary mask and resource container, as Windows has now made its own copies
        DeleteObject monoBmp
        Set resDIB = Nothing
        
        'Debug.Print "Custom cursor (""" & resTitle & """ : " & createCursorFromResource & ") created successfully!"
        
    Else
        Debug.Print "GDI+ couldn't load resource image - sorry!"
    End If
    
    Exit Function
    
End Function

'Load all relevant program cursors into memory
Public Sub initAllCursors()

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

'Use any 32bpp PNG resource as a cursor (yes, it's amazing!).  When setting the mouse pointer of VB objects, please use
' setPNGCursorToObject, below.
Public Sub setPNGCursorToHwnd(ByVal dstHwnd As Long, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    SetClassLong dstHwnd, GCL_HCURSOR, requestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Use any 32bpp PNG resource as a cursor (yes, it's amazing!).  Use this function preferentially over the previous one, if
' you can.  If a VB object does not have its MousePointer property set to "custom", it will override our attempts to set
' a custom mouse icon.
Public Sub setPNGCursorToObject(ByRef srcObject As Object, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    
    'Force VB to use a custom cursor
    srcObject.MousePointer = vbCustom
    
    SetClassLong srcObject.hWnd, GCL_HCURSOR, requestCustomCursor(pngTitle, curHotspotX, curHotspotY)
    
End Sub

'Set a single object to use the hand cursor
Public Sub setHandCursor(ByRef tControl As Object)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 99
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub setHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub setArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the arrow cursor
Public Sub setArrowCursor(ByRef tControl As Object)
    tControl.MousePointer = vbCustom
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the cross cursor
Public Sub setCrossCursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_CROSS)
End Sub
    
'Set a single form to use the Size All cursor
Public Sub setSizeAllCursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZEALL)
End Sub

'Set a single form to use the Size NESW cursor
Public Sub setSizeNESWCursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENESW)
End Sub

'Set a single form to use the Size NS cursor
Public Sub setSizeNSCursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENS)
End Sub

'Set a single form to use the Size NWSE cursor
Public Sub setSizeNWSECursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZENWSE)
End Sub

'Set a single form to use the Size WE cursor
Public Sub setSizeWECursor(ByRef tControl As Object)
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_SIZEWE)
End Sub

'If a custom PNG cursor has not been loaded, this function will load the PNG, convert it to cursor format, then store
' the cursor resource for future reference (so the image doesn't have to be loaded again).
Public Function requestCustomCursor(ByVal CursorName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0) As Long

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

'Given an image in the .exe's resource section (typically a PNG image), load it to a pdDIB object.
' The calling function is responsible for deleting the DIB once they are done with it.
Public Function loadResourceToDIB(ByVal resTitle As String, ByRef dstDIB As pdDIB, Optional ByVal vbSupportedFormat As Boolean = False) As Boolean
    
    'If the requested image is in a VB-compatible format (e.g. BMP), we don't need to use GDI+
    If vbSupportedFormat Then
    
        'Load the requested image into a temporary StdPicture object
        Dim tmppic As StdPicture
        Set tmppic = New StdPicture
        Set tmppic = LoadResPicture(resTitle, 0)
        
        'Copy that image into the supplied DIB
        If dstDIB.CreateFromPicture(tmppic) Then
            loadResourceToDIB = True
        Else
            loadResourceToDIB = False
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
                
                'If the image has an alpha channel, create a 32bpp DIB to receive it
                If (gdiPixelFormat And PixelFormatAlpha <> 0) Or (gdiPixelFormat And PixelFormatPAlpha <> 0) Then
                    dstDIB.createBlank tmpRect.fWidth, tmpRect.fHeight, 32
                Else
                    dstDIB.createBlank tmpRect.fWidth, tmpRect.fHeight, 24
                End If
                
                'Convert the GDI+ bitmap to a standard Windows hBitmap
                If GdipCreateHBITMAPFromBitmap(gdiBitmap, hBitmap, vbBlack) = 0 Then
                
                    'Select the hBitmap into a new DC so we can BitBlt it into the DIB
                    Dim gdiDC As Long
                    gdiDC = CreateCompatibleDC(0)
                    SelectObject gdiDC, hBitmap
                    
                    'Copy the GDI+ bitmap into the DIB
                    BitBlt dstDIB.getDIBDC, 0, 0, tmpRect.fWidth, tmpRect.fHeight, gdiDC, 0, 0, vbSrcCopy
                    
                    'Verify the alpha channel
                    If Not dstDIB.verifyAlphaChannel Then dstDIB.convertTo24bpp
                    
                    'Release the Windows-format bitmap and temporary device context
                    DeleteObject hBitmap
                    DeleteDC gdiDC
                    
                    'Release the GDI+ bitmap as well
                    GdipDisposeImage gdiBitmap
                    
                    'Free the memory stream
                    Set IStream = Nothing
                    
                    loadResourceToDIB = True
                    Exit Function
                
                Else
                    Debug.Print "GDI+ failed to create an HBITMAP for requested resource " & resTitle & " stream."
                End If
                
                'Release the GDI+ bitmap and mark the load as failed
                GdipDisposeImage gdiBitmap
                loadResourceToDIB = False
                Exit Function
                    
            Else
                Debug.Print "GDI+ failed to load requested resource " & resTitle & " stream."
            End If
        
            'Free the memory stream
            Set IStream = Nothing
            loadResourceToDIB = False
            Exit Function
        
        Else
            Debug.Print "Could not load requested resource " & resTitle & " from file."
        End If
        
        loadResourceToDIB = False
        Exit Function
    
    End If
        
End Function
