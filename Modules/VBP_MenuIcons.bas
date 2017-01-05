Attribute VB_Name = "Icons_and_Cursors"
'***************************************************************************
'PhotoDemon Icon and Cursor Handler
'Copyright 2012-2017 by Tanner Helland
'Created: 24/June/12
'Last updated: 22/Tanner/15
'Last updated by: Tanner
'Last update: Shuffle all icons to account for changes to the Adjustments menu
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
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

'System constants for retrieving system default icon sizes and related metrics
Private Const SM_CXICON As Long = 11
Private Const SM_CYICON As Long = 12
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50
Private Const LR_SHARED As Long = &H8000&
Private Const IMAGE_ICON As Long = 1
Private Const WM_SETICON As Long = &H80
Private Const ICON_SMALL As Long = 0
Private Const ICON_BIG As Long = 1

'API needed for converting PNG data to icon or cursor format
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef mImage As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hbmReturn As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef hBmpReturn As Long, ByVal Background As Long) As GP_Result
Private Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal gdiBitmap As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal gdiBitmap As Long, ByRef PixelFormat As Long) As Long

'Used to check GDI+ images for alpha channels
Private Const PixelFormatAlpha = &H40000             ' Has an alpha component
Private Const PixelFormatPAlpha = &H80000            ' Pre-multiplied alpha

'GDI+ types and constants
Private Const UnitPixel As Long = &H2&

'Type required to create an icon on-the-fly
Private Type ICONINFO
   fIcon As Boolean
   xHotspot As Long
   yHotspot As Long
   hbmMask As Long
   hbmColor As Long
End Type

'Used to apply and manage custom cursors (without subclassing)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
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

Private numOfCustomCursors As Long
Private customCursorNames() As String
Private customCursorHandles() As Long

'This array will be used to store our dynamically created icon handles so we can delete them on program exit
Private Const INITIAL_ICON_CACHE_SIZE As Long = 16
Private m_numOfIcons As Long
Private m_iconHandles() As Long

'As of v7.0, icon creation and destruction is tracked locally.
Private m_IconsCreated As Long, m_IconsDestroyed As Long

'This constant is used for testing only.  It should always be set to TRUE for production code.
Private Const ALLOW_DYNAMIC_ICONS As Boolean = True

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

'PD's default large and small application icons.  These are cached for the duration of the current session.
Private m_DefaultIconLarge As Long, m_DefaultIconSmall As Long

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
        If (Not g_IsVistaOrLater) And (Not g_IsProgramCompiled) Then
            Debug.Print "XP + IDE detected.  Menu icons will be disabled for this session."
            Exit Sub
        End If
        
        .Init FormMain.hWnd, FixDPI(16), FixDPI(16)
        
    End With
            
    'Now that all menu icons are loaded, apply them to the proper menu entires
    ApplyAllMenuIcons
        
    '...and initialize the separate MRU icon handler.
    Set cMRUIcons = New clsMenuImage
    If g_IsVistaOrLater Then
        cMRUIcons.Init FormMain.hWnd, FixDPI(64), FixDPI(64)
    Else
        cMRUIcons.Init FormMain.hWnd, FixDPI(16), FixDPI(16)
    End If
        
End Sub

'Apply (and if necessary, dynamically load) menu icons to their proper menu entries.
Public Sub ApplyAllMenuIcons(Optional ByVal useDoEvents As Boolean = False)
    
    m_refreshOutsideProgressBar = useDoEvents

    'Load every icon from the resource file.  (Yes, there are a LOT of icons!)
        
    'File Menu
    'AddMenuIcon "NEWIMAGE", 0, 0      'New
    'AddMenuIcon "OPENIMG", 0, 1       'Open Image
    AddMenuIcon "file_new", 0, 0      'New
    AddMenuIcon "file_open", 0, 1       'Open Image
    
    AddMenuIcon "OPENREC", 0, 2       'Open recent
    AddMenuIcon "IMPORT", 0, 3        'Import
        '--> Import sub-menu
        AddMenuIcon "PASTE_IMAGE", 0, 3, 0 'From Clipboard (Paste as New Image)
        AddMenuIcon "SCANNER", 0, 3, 2     'Scan Image
        AddMenuIcon "SCANNERSEL", 0, 3, 3  'Select Scanner
        AddMenuIcon "DOWNLOAD", 0, 3, 5    'Online Image
        AddMenuIcon "SCREENCAP", 0, 3, 7   'Screen Capture
    
    'AddMenuIcon "CLOSE", 0, 5         'Close
    'AddMenuIcon "CLOSE", 0, 6         'Close All
    'AddMenuIcon "SAVE", 0, 8          'Save
    'AddMenuIcon "SAVECOPY", 0, 9      'Save copy
    'AddMenuIcon "SAVEAS", 0, 10       'Save As...
    
    AddMenuIcon "file_close", 0, 5         'Close
    'AddMenuIcon "file_close", 0, 6         'Close All
    AddMenuIcon "file_save", 0, 8          'Save
    AddMenuIcon "file_savedup", 0, 9      'Save copy
    AddMenuIcon "file_saveas", 0, 10       'Save As...
    
    AddMenuIcon "REVERT", 0, 11       'Revert
    AddMenuIcon "BCONVERT", 0, 13     'Batch operations
        '--> Batch sub-menu
        AddMenuIcon "BCONVERT", 0, 13, 0    'Batch process
        'AddMenuIcon "BCONVERT", 0, 13, 1    'Batch repair
    AddMenuIcon "PRINT", 0, 15        'Print
    AddMenuIcon "EXIT", 0, 17         'Exit
    
        
    'Edit Menu
    'AddMenuIcon "UNDO", 1, 0           'Undo
    'AddMenuIcon "REDO", 1, 1           'Redo
    AddMenuIcon "edit_undo", 1, 0           'Undo
    AddMenuIcon "edit_redo", 1, 1           'Redo
    AddMenuIcon "UNDOHISTORY", 1, 2    'Undo history browser
    
    AddMenuIcon "REPEAT", 1, 4         'Repeat previous action
    AddMenuIcon "FADE", 1, 5           'Fade previous action...
    
    AddMenuIcon "CUT", 1, 7            'Cut
    AddMenuIcon "CUT_LAYER", 1, 8      'Cut from layer
    AddMenuIcon "COPY", 1, 9           'Copy
    AddMenuIcon "COPY_LAYER", 1, 10    'Copy from layer
    AddMenuIcon "PASTE_IMAGE", 1, 11   'Paste as new image
    AddMenuIcon "PASTE_LAYER", 1, 12   'Paste as new layer
    AddMenuIcon "CLEAR", 1, 14         'Empty Clipboard
    
    'View Menu
    AddMenuIcon "zoom_fit", 2, 0    'Fit on Screen
    AddMenuIcon "zoom_in", 2, 2         'Zoom In
    AddMenuIcon "zoom_out", 2, 3        'Zoom Out
    AddMenuIcon "zoom_actual", 2, 9     'Zoom 100%
    
    'Image Menu
    AddMenuIcon "DUPLICATE", 3, 0      'Duplicate
    AddMenuIcon "RESIZE", 3, 2         'Resize
    AddMenuIcon "SMRTRESIZE", 3, 3     'Content-aware resize
    AddMenuIcon "CANVASSIZE", 3, 5     'Canvas resize
    AddMenuIcon "FITTOLAYER", 3, 6     'Fit canvas to active layer
    AddMenuIcon "FITALLLAYERS", 3, 7   'Fit canvas around all layers
    AddMenuIcon "CROPSEL", 3, 9        'Crop to Selection
    AddMenuIcon "TRIMEMPTY", 3, 10      'Trim
    AddMenuIcon "ROTATECW", 3, 12      'Rotate top-level
        '--> Rotate sub-menu
        AddMenuIcon "STRAIGHTEN", 3, 12, 0  'Straighten
        AddMenuIcon "ROTATECW", 3, 12, 2    'Rotate Clockwise
        AddMenuIcon "ROTATECCW", 3, 12, 3   'Rotate Counter-clockwise
        AddMenuIcon "ROTATE180", 3, 12, 4   'Rotate 180
        If g_ImageFormats.FreeImageEnabled Then AddMenuIcon "ROTATEANY", 3, 12, 5  'Rotate Arbitrary
    AddMenuIcon "MIRROR", 3, 13        'Mirror
    AddMenuIcon "FLIP", 3, 14          'Flip
    'addMenuIcon "ISOMETRIC", 3, 12     'Isometric      'NOTE: isometric was removed in v6.4.
    AddMenuIcon "REDUCECOLORS", 3, 16  'Indexed color (Reduce Colors)
    If g_ImageFormats.FreeImageEnabled Then FormMain.MnuImage(16).Enabled = True Else FormMain.MnuImage(16).Enabled = False
    AddMenuIcon "METADATA", 3, 18      'Metadata (top-level)
        '--> Metadata sub-menu
        AddMenuIcon "BROWSEMD", 3, 18, 0     'Browse metadata
        AddMenuIcon "COUNTCOLORS", 3, 18, 2  'Count Colors
        AddMenuIcon "MAPPHOTO", 3, 18, 3     'Map photo location
    
    'Layer menu
    AddMenuIcon "layer_add", 4, 0        'Add layer (top-level)
        '--> Add layer sub-menu
        AddMenuIcon "ADDLAYER", 4, 0, 0             'Add blank layer
        AddMenuIcon "DUPL_LAYER", 4, 0, 1          'Add duplicate layer
        AddMenuIcon "PASTE_LAYER", 4, 0, 3          'Add layer from clipboard
        AddMenuIcon "ADDLAYERFILE", 4, 0, 4             'Add layer from file
    AddMenuIcon "layer_delete", 4, 1        'Delete layer (top-level)
        '--> Delete layer sub-menu
        AddMenuIcon "DELLAYER", 4, 1, 0       'Delete current layer
        AddMenuIcon "DELLAYERHDN", 4, 1, 1       'Delete all hidden layers
    AddMenuIcon "MERGE_UP", 4, 3         'Merge up
    AddMenuIcon "MERGE_DOWN", 4, 4       'Merge down
    AddMenuIcon "LAYERORDER", 4, 5      'Order (top-level)
        '--> Order layer sub-menu
        AddMenuIcon "layer_up", 4, 5, 0     'Raise layer
        AddMenuIcon "layer_down", 4, 5, 1     'Lower layer
        AddMenuIcon "LAYERTOTOP", 4, 5, 3     'Raise to top
        AddMenuIcon "LAYERTOBTM", 4, 5, 4     'Lower to bottom
    AddMenuIcon "ROTATECW", 4, 7         'Layer Orientation (top-level)
        '--> Orientation sub-menu
        AddMenuIcon "STRAIGHTEN", 4, 7, 0   'Straighten
        AddMenuIcon "ROTATECW", 4, 7, 2     'Rotate Clockwise
        AddMenuIcon "ROTATECCW", 4, 7, 3    'Rotate Counter-clockwise
        AddMenuIcon "ROTATE180", 4, 7, 4    'Rotate 180
        If g_ImageFormats.FreeImageEnabled Then AddMenuIcon "ROTATEANY", 4, 7, 5   'Rotate Arbitrary
        AddMenuIcon "MIRROR", 4, 7, 7       'Mirror
        AddMenuIcon "FLIP", 4, 7, 8         'Flip
    AddMenuIcon "RESIZE", 4, 8           'Layer Size (top-level)
        '--> Size sub-menu
        AddMenuIcon "RESETSIZE", 4, 8, 0        'Reset to original size
        AddMenuIcon "RESIZE", 4, 8, 2        'Resize
        AddMenuIcon "SMRTRESIZE", 4, 8, 3    'Content-aware resize
    AddMenuIcon "CROPSEL", 4, 9          'Crop to Selection
    AddMenuIcon "TRANSPARENCY", 4, 11    'Layer Transparency
        '--> Transparency sub-menu
        AddMenuIcon "ADDTRANS", 4, 11, 0     'Add alpha channel
        AddMenuIcon "GREENSCREEN", 4, 11, 1  'Color to alpha
        AddMenuIcon "REMOVETRANS", 4, 11, 3  'Remove alpha channel
    'addMenuIcon "RASTERIZE", 4, 13       'Rasterize layer
    AddMenuIcon "FLATTEN", 4, 15         'Flatten image
    AddMenuIcon "MERGEVISIBLE", 4, 16    'Merge visible layers
    
    'Select Menu
    AddMenuIcon "SELECTALL", 5, 0       'Select all
    AddMenuIcon "SELECTNONE", 5, 1      'Select none
    AddMenuIcon "SELECTINVERT", 5, 2    'Invert selection
    AddMenuIcon "SELECTGROW", 5, 4      'Grow selection
    AddMenuIcon "SELECTSHRINK", 5, 5    'Shrink selection
    AddMenuIcon "SELECTBORDER", 5, 6    'Border selection
    AddMenuIcon "SELECTFTHR", 5, 7      'Feather selection
    AddMenuIcon "SELECTSHRP", 5, 8      'Sharpen selection
    AddMenuIcon "SELECTERASE", 5, 10    'Erase selected area
    AddMenuIcon "SELECTLOAD", 5, 12     'Load selection from file
    AddMenuIcon "SELECTSAVE", 5, 13     'Save selection to file
    AddMenuIcon "SELECTEXPORT", 5, 14   'Export selection (top-level)
        '--> Export Selection sub-menu
        AddMenuIcon "EXPRTSELAREA", 5, 14, 0  'Export selected area as image
        AddMenuIcon "EXPRTSELMASK", 5, 14, 1  'Export selection mask as image
    
    'Adjustments Menu
    
    'Auto correct
    AddMenuIcon "AUTOCORRECT", 6, 0     'Auto-correct (top-level)
        AddMenuIcon "HSL", 6, 0, 0          'Color
        AddMenuIcon "BRIGHT", 6, 0, 1       'Contrast
        AddMenuIcon "LIGHTING", 6, 0, 2     'Lighting
        AddMenuIcon "SHDWHGHLGHT", 6, 0, 3  'Shadow/Highlight
        
    'Auto enhance
    AddMenuIcon "AUTOENHANCE", 6, 1     'Auto-enhance (top-level)
        AddMenuIcon "HSL", 6, 1, 0          'Color
        AddMenuIcon "BRIGHT", 6, 1, 1       'Contrast
        AddMenuIcon "LIGHTING", 6, 1, 2     'Lighting
        AddMenuIcon "SHDWHGHLGHT", 6, 1, 3  'Shadow/Highlight
        
    'Adjustment shortcuts (top-level menu items)
    AddMenuIcon "GRAYSCALE", 6, 3       'Black and white
    AddMenuIcon "BRIGHT", 6, 4          'Brightness/Contrast
    AddMenuIcon "COLORBALANCE", 6, 5    'Color balance
    AddMenuIcon "CURVES", 6, 6          'Curves
    AddMenuIcon "LEVELS", 6, 7          'Levels
    AddMenuIcon "SHDWHGHLGHT", 6, 8     'Shadow/highlight
    AddMenuIcon "VIBRANCE", 6, 9        'Vibrance
    AddMenuIcon "WHITEBAL", 6, 10       'White Balance
       
    'Channels
    AddMenuIcon "CHANNELMIX", 6, 12    'Channels top-level
        AddMenuIcon "CHANNELMIX", 6, 12, 0   'Channel mixer
        AddMenuIcon "RECHANNEL", 6, 12, 1    'Rechannel
        AddMenuIcon "CHANNELMAX", 6, 12, 3   'Channel max
        AddMenuIcon "CHANNELMIN", 6, 12, 4   'Channel min
        AddMenuIcon "COLORSHIFTL", 6, 12, 6  'Shift Left
        AddMenuIcon "COLORSHIFTR", 6, 12, 7  'Shift Right
            
    'Color
    AddMenuIcon "HSL", 6, 13           'Color balance
        AddMenuIcon "COLORBALANCE", 6, 13, 0 'Color balance
        AddMenuIcon "WHITEBAL", 6, 13, 1     'White Balance
        AddMenuIcon "HSL", 6, 13, 3          'HSL adjustment
        AddMenuIcon "TEMPERATURE", 6, 13, 4  'Temperature
        AddMenuIcon "TINT", 6, 13, 5         'Tint
        AddMenuIcon "VIBRANCE", 6, 13, 6     'Vibrance
        AddMenuIcon "GRAYSCALE", 6, 13, 8    'Black and white
        AddMenuIcon "COLORIZE", 6, 13, 9     'Colorize
        AddMenuIcon "REPLACECLR", 6, 13, 10  'Replace color
        AddMenuIcon "SEPIA", 6, 13, 11       'Sepia
    
    'Histogram
    AddMenuIcon "HISTOGRAM", 6, 14      'Histogram top-level
        AddMenuIcon "HISTOGRAM", 6, 14, 0     'Display Histogram
        AddMenuIcon "EQUALIZE", 6, 14, 2      'Equalize
        AddMenuIcon "STRETCH", 6, 14, 3       'Stretch
    
    'Invert
    AddMenuIcon "INVERT", 6, 15         'Invert top-level
        AddMenuIcon "INVCMYK", 6, 15, 0     'Invert CMYK
        AddMenuIcon "INVHUE", 6, 15, 1       'Invert Hue
        AddMenuIcon "INVRGB", 6, 15, 2       'Invert RGB
        AddMenuIcon "INVCOMPOUND", 6, 15, 4  'Compound Invert
        
    'Lighting
    AddMenuIcon "LIGHTING", 6, 16       'Lighting top-level
        AddMenuIcon "BRIGHT", 6, 16, 0       'Brightness/Contrast
        AddMenuIcon "CURVES", 6, 16, 1       'Curves
        AddMenuIcon "GAMMA", 6, 16, 2        'Gamma Correction
        AddMenuIcon "LEVELS", 6, 16, 3       'Levels
        AddMenuIcon "SHDWHGHLGHT", 6, 16, 4  'Shadow/Highlight
        
    'Monochrome
    AddMenuIcon "MONOCHROME", 6, 17      'Monochrome
        AddMenuIcon "COLORTOMONO", 6, 17, 0   'Color to monochrome
        AddMenuIcon "MONOTOCOLOR", 6, 17, 1   'Monochrome to grayscale
        
    'Photography
    AddMenuIcon "PHOTOFILTER", 6, 18      'Photography top-level
        AddMenuIcon "EXPOSURE", 6, 18, 0     'Exposure
        AddMenuIcon "HDR", 6, 18, 1          'HDR
        AddMenuIcon "PHOTOFILTER", 6, 18, 2  'Photo filters
        'addMenuIcon "REDEYE", 6, 18, 3       'Red-eye removal
        AddMenuIcon "SPLITTONE", 6, 18, 4    'Split-toning
    
    
    'Effects (Filters) Menu
    AddMenuIcon "ARTISTIC", 7, 0        'Artistic
        '--> Artistic sub-menu
        AddMenuIcon "PENCIL", 7, 0, 0         'Pencil
        AddMenuIcon "COMICBOOK", 7, 0, 1      'Comic book
        AddMenuIcon "FIGGLASS", 7, 0, 2       'Figured glass
        AddMenuIcon "FILMNOIR", 7, 0, 3       'Film Noir
        AddMenuIcon "GLASSTILES", 7, 0, 4     'Glass tiles
        AddMenuIcon "KALEIDOSCOPE", 7, 0, 5   'Kaleidoscope
        AddMenuIcon "MODERNART", 7, 0, 6      'Modern Art
        AddMenuIcon "OILPAINTING", 7, 0, 7    'Oil painting
        AddMenuIcon "POSTERIZE", 7, 0, 8      'Posterize
        AddMenuIcon "RELIEF", 7, 0, 9         'Relief
        AddMenuIcon "STAINEDGLASS", 7, 0, 10  'Stained glass
    
    AddMenuIcon "BLUR", 7, 1            'Blur
        '--> Blur sub-menu
        AddMenuIcon "BOXBLUR", 7, 1, 0        'Box Blur
        AddMenuIcon "GAUSSBLUR", 7, 1, 1      'Gaussian Blur
        AddMenuIcon "SMARTBLUR", 7, 1, 2      'Surface Blur (formerly Smart Blur)
        AddMenuIcon "MOTIONBLUR", 7, 1, 4     'Motion Blur
        AddMenuIcon "RADIALBLUR", 7, 1, 5     'Radial Blur
        AddMenuIcon "ZOOMBLUR", 7, 1, 6       'Zoom Blur
        AddMenuIcon "CHROMABLUR", 7, 1, 8     'Kuwahara
        
    AddMenuIcon "DISTORT", 7, 2         'Distort
        '--> Distort sub-menu
        AddMenuIcon "LENSDISTORT", 7, 2, 0    'Apply lens distortion
        AddMenuIcon "FIXLENS", 7, 2, 1        'Remove or correct existing lens distortion
        
        'addMenuIcon "DONUT", 7, 2, 3          'Donut
        AddMenuIcon "PINCHWHIRL", 7, 2, 4     'Pinch and whirl
        AddMenuIcon "POKE", 7, 2, 5           'Poke
        AddMenuIcon "RIPPLE", 7, 2, 6         'Ripple
        AddMenuIcon "SQUISH", 7, 2, 7         'Squish (formerly Fixed Perspective)
        AddMenuIcon "SWIRL", 7, 2, 8          'Swirl
        AddMenuIcon "WAVES", 7, 2, 9          'Waves
        
        AddMenuIcon "MISCDISTORT", 7, 2, 11   'Miscellaneous distort functions
                
    AddMenuIcon "EDGES", 7, 3           'Edges
        '--> Edges sub-menu
        AddMenuIcon "EMBOSS", 7, 3, 0         'Emboss / Engrave
        AddMenuIcon "EDGEENHANCE", 7, 3, 1    'Enhance Edges
        AddMenuIcon "EDGES", 7, 3, 2          'Find Edges
        'addMenuIcon "RANGEFILTER", 7, 3, 4    'Range filter
        AddMenuIcon "TRACECONTOUR", 7, 3, 4   'Trace Contour
        
    AddMenuIcon "SUNSHINE", 7, 4        'Lights and shadows
        '--> Lights and shadows sub-menu
        AddMenuIcon "BLACKLIGHT", 7, 4, 0     'Blacklight
        AddMenuIcon "CROSSSCREEN", 7, 4, 1    'Cross-screen (stars)
        AddMenuIcon "LENSFLARE", 7, 4, 2      'Lens flare
        AddMenuIcon "RAINBOW", 7, 4, 3        'Rainbow
        AddMenuIcon "SUNSHINE", 7, 4, 4       'Sunshine
        AddMenuIcon "DILATE", 7, 4, 6         'Dilate
        AddMenuIcon "ERODE", 7, 4, 7          'Erode
    
    AddMenuIcon "NATURAL", 7, 5         'Natural
        '--> Natural sub-menu
        AddMenuIcon "ATMOSPHERE", 7, 5, 0     'Atmosphere
        AddMenuIcon "FOG", 7, 5, 1            'Fog
        AddMenuIcon "FREEZE", 7, 5, 2         'Freeze
        AddMenuIcon "BURN", 7, 5, 3           'Ignite
        AddMenuIcon "LAVA", 7, 5, 4           'Lava
        AddMenuIcon "STEEL", 7, 5, 5          'Steel
        AddMenuIcon "RAIN", 7, 5, 6           'Water
        
    AddMenuIcon "NOISE", 7, 6           'Noise
        '--> Noise sub-menu
        AddMenuIcon "FILMGRAIN", 7, 6, 0      'Film grain
        AddMenuIcon "ADDNOISE", 7, 6, 1       'Add Noise
        
        'addMenuIcon "ANISOTROPIC", 7, 6, 3    'Anisotropic diffusion
        AddMenuIcon "BILATERAL", 7, 6, 4      'Bilateral smoothing
        'addMenuIcon "MEANSHIFT", 7, 6, 5      'Mean-shift
        AddMenuIcon "MEDIAN", 7, 6, 6         'Median
        
    AddMenuIcon "PIXELATE", 7, 7        'Pixelate
        '--> Pixelate sub-menu
        'addMenuIcon "CLRHALFTONE", 7, 7, 0   'Color halftone (TODO)
        'addMenuIcon "CRYTALLIZE", 7, 7, 1    'Crystallize (TODO)
        AddMenuIcon "FRAGMENT", 7, 7, 2      'Fragment
        'addMenuIcon "MEZZOTINT", 7, 7, 3     'Mezzotint (TODO)
        AddMenuIcon "PIXELATE", 7, 7, 4      'Mosaic (formerly Pixelate)
    
    AddMenuIcon "SHARPEN", 7, 8         'Sharpen
        '--> Sharpen sub-menu
        AddMenuIcon "SHARPEN", 7, 8, 0       'Sharpen
        AddMenuIcon "UNSHARP", 7, 8, 1       'Unsharp
        
    AddMenuIcon "STYLIZE", 7, 9        'Stylize
        '--> Stylize sub-menu
        AddMenuIcon "ANTIQUE", 7, 9, 0       'Antique (Sepia)
        AddMenuIcon "DIFFUSE", 7, 9, 1       'Diffuse
        'addMenuIcon "PORTGLOW", 7, 9, 2      'Portrait glow
        AddMenuIcon "SOLARIZE", 7, 9, 3      'Solarize
        AddMenuIcon "TWINS", 7, 9, 4         'Twins
        AddMenuIcon "VIGNETTE", 7, 9, 5      'Vignetting
        
    AddMenuIcon "PANANDZOOM", 7, 10        'Transform
        '--> Transform sub-menu
        AddMenuIcon "PANANDZOOM", 7, 10, 0    'Pan and zoom
        AddMenuIcon "PERSPECTIVE", 7, 10, 1   'Perspective (free)
        AddMenuIcon "POLAR", 7, 10, 2         'Polar conversion
        AddMenuIcon "ROTATECW", 7, 10, 3      'Rotate
        AddMenuIcon "SHEAR", 7, 10, 4         'Shear
        AddMenuIcon "SPHERIZE", 7, 10, 5      'Spherize
        
    AddMenuIcon "CUSTFILTER", 7, 12     'Custom Filter
    
    'addMenuIcon "OTHER", 7, 14           'Experimental
        '--> Experimental sub-menu
        'addMenuIcon "ALIEN", 7, 14, 0          'Alien
        'addMenuIcon "DREAM", 7, 14, 1          'Dream
        'addMenuIcon "RADIOACTIVE", 7, 14, 2    'Radioactive
        'addMenuIcon "SYNTHESIZE", 7, 14, 3     'Synthesize
        'addMenuIcon "HEATMAP", 7, 14, 4        'Thermograph
        'addMenuIcon "VIBRATE", 7, 14, 5        'Vibrate
    
    'Tools Menu
    AddMenuIcon "LANGUAGES", 8, 0       'Languages
    AddMenuIcon "LANGEDITOR", 8, 1      'Language editor
    
    AddMenuIcon "RECORDMACRO", 8, 3      'Macros
        '--> Macro sub-menu
        AddMenuIcon "RECORDMACRO", 8, 3, 0    'Start Recording
        AddMenuIcon "RECORDSTOP", 8, 3, 1     'Stop Recording
    AddMenuIcon "PLAYMACRO", 8, 4       'Play saved macro
    AddMenuIcon "RECENTMACROS", 8, 5    'Recent macros
    
    AddMenuIcon "PREFERENCES", 8, 7     'Options (Preferences)
    AddMenuIcon "PLUGIN", 8, 8          'Plugin Manager
    
    'Window Menu
    AddMenuIcon "NEXTIMAGE", 9, 7       'Next image
    AddMenuIcon "PREVIMAGE", 9, 8       'Previous image
    
    'Help Menu
    AddMenuIcon "FAVORITE", 10, 0        'Donate
    AddMenuIcon "UPDATES", 10, 2         'Check for updates
    AddMenuIcon "FEEDBACK", 10, 3        'Submit Feedback
    AddMenuIcon "BUG", 10, 4             'Submit Bug
    AddMenuIcon "PDWEBSITE", 10, 6       'Visit the PhotoDemon website
    AddMenuIcon "DOWNLOADSRC", 10, 7     'Download source code
    AddMenuIcon "LICENSE", 10, 8         'License
    AddMenuIcon "ABOUT", 10, 10          'About PD
    
    'When we're done, reset the doEvents tracker
    m_refreshOutsideProgressBar = False
    
End Sub

'This new, simpler technique for adding menu icons requires only the menu location (including sub-menus) and the icon's identifer
' in the resource file.  If the icon has already been loaded, it won't be loaded again; instead, the function will check the list
' of loaded icons and automatically fill in the numeric identifier as necessary.
Private Sub AddMenuIcon(ByVal resID As String, ByVal topMenu As Long, ByVal subMenu As Long, Optional ByVal subSubMenu As Long = -1)
    
    On Error GoTo MenuIconNotFound
    
    Dim i As Long
    Dim iconLocation As Long
    Dim iconAlreadyLoaded As Boolean
    
    iconAlreadyLoaded = False
    
    'Loop through all icons that have been loaded, and see if this one has been requested already.
    For i = 0 To curIcon
    
        If (iconNames(i) = resID) Then
            iconAlreadyLoaded = True
            iconLocation = i
            Exit For
        End If
    
    Next i
    
    'If the icon was not found, load it and add it to the list
    If (Not iconAlreadyLoaded) Then
        
        If (Not cMenuImage Is Nothing) Then
            
            'Attempt to load the image from our internal resource manager
            Dim loadedInternally As Boolean: loadedInternally = False
            If (Not g_Resources Is Nothing) Then
                If g_Resources.AreResourcesAvailable Then
                    Dim tmpDIB As pdDIB
                    loadedInternally = g_Resources.LoadImageResource(resID, tmpDIB, 16, 16, 0.5, True)
                    If loadedInternally Then cMenuImage.AddImageFromDIB tmpDIB
                End If
            End If
            
            If (Not loadedInternally) Then
                cMenuImage.AddImageFromStream LoadResData(resID, "CUSTOM")
            End If
            
            iconNames(curIcon) = resID
            iconLocation = curIcon
            curIcon = curIcon + 1
        End If
        
    End If
        
    'Place the icon onto the requested menu
    If (Not cMenuImage Is Nothing) Then
    
        If (subSubMenu = -1) Then
            cMenuImage.PutImageToVBMenu iconLocation, subMenu, topMenu
        Else
            cMenuImage.PutImageToVBMenu iconLocation, subSubMenu, topMenu, subMenu
        End If
        
    End If
    
    'If an outside progress bar needs to refresh, do so now
    If m_refreshOutsideProgressBar Then DoEvents

MenuIconNotFound:

End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' Note that to keep the code small, all changeable icons are refreshed whenever this is called.
Public Sub ResetMenuIcons()
        
    'Redraw the Undo/Redo menus
    'AddMenuIcon "UNDO", 1, 0     'Undo
    'AddMenuIcon "REDO", 1, 1     'Redo
    AddMenuIcon "edit_undo", 1, 0     'Undo
    AddMenuIcon "edit_redo", 1, 1     'Redo
    
    'Redraw the Repeat and Fade menus
    AddMenuIcon "REPEAT", 1, 4         'Repeat previous action
    AddMenuIcon "FADE", 1, 5           'Fade previous action...
    
    'NOTE! In the future, when icons are available for the Repeat and Fade menu items, we will need to add their refreshes
    ' to this list (as their captions dynamically change at run-time).
    
    'Redraw the Window menu, as some of its menus will be en/disabled according to the docking status of image windows
    AddMenuIcon "NEXTIMAGE", 9, 7       'Next image
    AddMenuIcon "PREVIMAGE", 9, 8       'Previous image
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    If Not (g_RecentFiles Is Nothing) Then
    
        Dim numOfMRUFiles As Long
        numOfMRUFiles = g_RecentFiles.MRU_ReturnCount()
        
        'Vista+ gets nice, large icons added later in the process.  XP is stuck with 16x16 ones, which we add now.
        If Not g_IsVistaOrLater Then
            AddMenuIcon "LOADALL", 0, 2, numOfMRUFiles + 1
            AddMenuIcon "CLEARRECENT", 0, 2, numOfMRUFiles + 2
        End If
        
        'Repeat the same steps for the Recent Macro list.  Note that a larger icon is never used for this list, because we don't have
        ' large thumbnail images present.
        Dim numOfMRUFiles_Macro As Long
        numOfMRUFiles_Macro = g_RecentMacros.MRU_ReturnCount
        AddMenuIcon "CLEARRECENT", 8, 5, numOfMRUFiles_Macro + 1
        
    End If
    
    'Clear the current MRU icon cache.
    ' (Note added 01 Jan 2014 - RR has reported an IDE error on the following line, which means this function is somehow being
    '  triggered before loadMenuIcons above.  I cannot reproduce this behavior, so instead, we now perform a single initialization
    '  check before attempting to load MRU icons.)
    If (Not cMRUIcons Is Nothing) Then
        
        cMRUIcons.Clear
        Dim tmpFilename As String
        
        'Load a placeholder image for missing MRU entries
        cMRUIcons.AddImageFromStream LoadResData("MRUHOLDER", "CUSTOM")
        
        'This counter will be used to track the current position of loaded thumbnail images into the icon collection
        Dim iconLocation As Long
        iconLocation = 0
        
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        'Loop through the MRU list, and attempt to load thumbnail images for each entry
        Dim i As Long
        For i = 0 To numOfMRUFiles
        
            'Start by seeing if an image exists for this MRU entry
            tmpFilename = g_RecentFiles.GetMRUThumbnailPath(i)
            
            If (Len(tmpFilename) <> 0) Then
            
                'If the file exists, add it to the MRU icon handler
                If cFile.FileExist(tmpFilename) Then
                    
                    iconLocation = iconLocation + 1
                    cMRUIcons.AddImageFromFile tmpFilename
                    cMRUIcons.PutImageToVBMenu iconLocation, i, 0, 2
                
                'If a thumbnail for this file does not exist, supply a placeholder image (Vista+ only; on XP it will simply be blank)
                Else
                    If g_IsVistaOrLater Then cMRUIcons.PutImageToVBMenu 0, i, 0, 2
                End If
                
            End If
            
        Next i
            
        'Vista+ users now get their nice, large "load all recent files" and "clear list" icons.
        If g_IsVistaOrLater Then
            cMRUIcons.AddImageFromStream LoadResData("LOADALLLRG", "CUSTOM")
            cMRUIcons.PutImageToVBMenu iconLocation + 1, numOfMRUFiles + 1, 0, 2
            
            cMRUIcons.AddImageFromStream LoadResData("CLEARRECLRG", "CUSTOM")
            cMRUIcons.PutImageToVBMenu iconLocation + 2, numOfMRUFiles + 2, 0, 2
        End If
        
    End If
        
End Sub

'Convert a DIB - any DIB! - to an icon via CreateIconIndirect.  Transparency will be preserved, and by default, the icon will be created
' at the current image's size (though you can specify a custom size if you wish).  Ideally, the passed DIB will have been created using
' the pdImage function "RequestThumbnail".
'
'FreeImage is currently required for this function, because it provides a simple way to move between DIBs and DDBs.  I could rewrite
' the function without FreeImage's help, but frankly don't consider it worth the trouble.
Public Function GetIconFromDIB(ByRef srcDIB As pdDIB, Optional iconSize As Long = 0, Optional ByVal flipVertically As Boolean = False) As Long

    If Not g_ImageFormats.FreeImageEnabled Then
        GetIconFromDIB = 0
        Exit Function
    End If
    
    Dim fi_DIB As Long
    fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB, flipVertically)
    
    'If the iconSize parameter is 0, use the current DIB's dimensions.  Otherwise, resize it as requested.
    If iconSize = 0 Then
        iconSize = srcDIB.GetDIBWidth
    Else
        fi_DIB = FreeImage_RescaleByPixel(fi_DIB, iconSize, iconSize, True, FILTER_BILINEAR)
    End If
    
    If (fi_DIB <> 0) Then
    
        'Icon generation has a number of quirks.  One is that even if you want a 32bpp icon, you still must supply a blank
        ' monochrome mask for the icon, even though the API just discards it.  Prepare such a mask now.
        Dim monoBmp As Long
        monoBmp = CreateBitmap(iconSize, iconSize, 1&, 1&, ByVal 0&)
        
        'Create a header for the icon we desire, then use CreateIconIndirect to create it.
        Dim icoInfo As ICONINFO
        With icoInfo
            .fIcon = True
            .xHotspot = iconSize
            .yHotspot = iconSize
            .hbmMask = monoBmp
            .hbmColor = FreeImage_GetBitmapForDevice(fi_DIB)
        End With
        
        GetIconFromDIB = CreateNewIcon(icoInfo)
        
        'Delete the temporary monochrome mask and DDB
        DeleteObject monoBmp
        DeleteObject icoInfo.hbmColor
    
    Else
        GetIconFromDIB = 0
    End If
    
    'Release FreeImage's copy of the source DIB
    If fi_DIB <> 0 Then FreeImage_UnloadEx fi_DIB
    
End Function

'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
' Note that this function currently requires the FreeImage plugin to be present on the system.
Public Sub CreateCustomFormIcons(ByRef srcImage As pdImage)

    If (Not ALLOW_DYNAMIC_ICONS) Then Exit Sub
    If (Not g_ImageFormats.FreeImageEnabled) Then Exit Sub
    If (srcImage Is Nothing) Then Exit Sub
    
    Dim thumbDIB As pdDIB
    
    'Request a 32x32 thumbnail version of the current image
    If srcImage.RequestThumbnail(thumbDIB, 32) Then

        'Request two icon-format versions of the generated thumbnail.
        ' (Taskbar icons are generally 32x32.  Form titlebar icons are generally 16x16.)
        Dim hIcon32 As Long, hIcon16 As Long
        hIcon32 = GetIconFromDIB(thumbDIB, , True)
        hIcon16 = GetIconFromDIB(thumbDIB, 16, False)   'Truthfully, I have no idea why this icon must be treated as upside-down.  FreeImage bug, perhaps?
        
        'Each pdImage instance stores its custom icon handles, which simplifies the process of synchronizing PD's icons
        ' to any given image if the user is working with multiple images at once.  Retrieve the old handles now, so we
        ' can free them after we set the new ones.
        Dim oldIcon32 As Long, oldIcon16 As Long
        oldIcon32 = srcImage.curFormIcon32
        oldIcon16 = srcImage.curFormIcon16
        
        'Set the new icons, then free the old ones
        srcImage.curFormIcon32 = hIcon32
        srcImage.curFormIcon16 = hIcon16
        If (oldIcon32 <> 0) Then ReleaseIcon oldIcon32
        If (oldIcon16 <> 0) Then ReleaseIcon oldIcon16
        
    Else
        Debug.Print "WARNING!  Image refused to provide a thumbnail!"
    End If

End Sub

Private Function CreateNewIcon(ByRef icoStruct As ICONINFO) As Long
    CreateNewIcon = CreateIconIndirect(icoStruct)
    If ((CreateNewIcon <> 0) And icoStruct.fIcon) Then m_IconsCreated = m_IconsCreated + 1
End Function

Public Sub ReleaseIcon(ByVal hIcon As Long)
    If (hIcon <> 0) Then
        DestroyIcon hIcon
        m_IconsDestroyed = m_IconsDestroyed + 1
    End If
End Sub

Public Function GetCreatedIconCount(Optional ByRef iconsCreated As Long, Optional ByRef iconsDestroyed As Long) As Long
    iconsCreated = m_IconsCreated
    iconsDestroyed = m_IconsDestroyed
    GetCreatedIconCount = m_IconsCreated - m_IconsDestroyed
End Function

'Needs to be run only once, at the start of the program
Public Sub InitializeIconHandler()
    m_numOfIcons = 0
    ReDim m_iconHandles(0 To INITIAL_ICON_CACHE_SIZE - 1) As Long
End Sub

Private Sub AddIconToList(ByVal hIcon As Long)
    
    If m_numOfIcons > UBound(m_iconHandles) Then
        ReDim Preserve m_iconHandles(0 To UBound(m_iconHandles) * 2 + 1) As Long
    End If
    
    m_iconHandles(m_numOfIcons) = hIcon
    m_numOfIcons = m_numOfIcons + 1

End Sub

'Remove all icons generated since the program launched
Public Sub DestroyAllIcons()

    If (m_numOfIcons = 0) Then Exit Sub
    
    Dim i As Long
    For i = 0 To m_numOfIcons - 1
        If m_iconHandles(i) <> 0 Then ReleaseIcon m_iconHandles(i)
    Next i
    
    'Reinitialize the icon handler, which will also reset the icon count and handle array
    InitializeIconHandler

End Sub

'Given an image in the .exe's resource section (typically a PNG image), return an icon handle to it (hIcon).
' The calling function is responsible for deleting this object once they are done with it.
Public Function CreateIconFromResource(ByVal resTitle As String) As Long
    
    'Start by extracting the PNG data into a bytestream
    Dim ImageData() As Byte
    ImageData() = LoadResData(resTitle, "CUSTOM")
    
    Dim hBitmap As Long, hIcon As Long
    
    Dim IStream As IUnknown
    Set IStream = VB_Hacks.GetStreamFromVBArray(VarPtr(ImageData(0)), UBound(ImageData) - LBound(ImageData) + 1)
    
    If Not (IStream Is Nothing) Then
        
        'Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine
        If GdipLoadImageFromStream(IStream, hBitmap) = 0 Then
        
            'hBitmap now contains the PNG file as an hBitmap (obviously).  Now we need to convert it to icon format.
            If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
                CreateIconFromResource = hIcon
            Else
                CreateIconFromResource = 0
            End If
            
            GdipDisposeImage hBitmap
                
        End If
    
        Set IStream = Nothing
    
    End If
    
    Exit Function
    
End Function

'Given an image in the .exe's resource section (typically a PNG image), return it as a cursor object.
' The calling function is responsible for deleting the cursor once they are done with it.
Public Function CreateCursorFromResource(ByVal resTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0) As Long
    
    'Start by extracting the PNG resource data into a pdLayer object.
    Dim resDIB As pdDIB
    Set resDIB = New pdDIB
                
    If LoadResourceToDIB(resTitle, resDIB) Then
    
        'If the user is running at non-96 DPI, we now need to perform DPI correction on the cursor.
        ' NOTE: this line will need to be revisited if high-DPI cursors are added to the resource file.  The reason
        '       I'm not making the change now is because PD's current cursor are not implemented uniformly, so I
        '       need to standardize their size and layout before implementing a universal "resize per DPI" check.
        '       The proper way to do this would be to retrieve cursor size from the system, then resize anything
        '       that isn't already that size - I've made a note to do this eventually.
        If FixDPI(96) <> 96 Then
        
            'Create a temporary copy of the image
            Dim dpiDIB As pdDIB
            Set dpiDIB = New pdDIB
            
            dpiDIB.CreateFromExistingDIB resDIB
            
            'Erase and resize the primary DIB
            resDIB.CreateBlank FixDPI(dpiDIB.GetDIBWidth), FixDPI(dpiDIB.GetDIBHeight), dpiDIB.GetDIBColorDepth
            
            'Use GDI+ to resize the cursor from dpiDIB into resDIB
            GDIPlusResizeDIB resDIB, 0, 0, resDIB.GetDIBWidth, resDIB.GetDIBHeight, dpiDIB, 0, 0, dpiDIB.GetDIBWidth, dpiDIB.GetDIBHeight, GP_IM_NearestNeighbor
        
            'Release our temporary DIB
            Set dpiDIB = Nothing
        
        End If
        
        'Generate a blank monochrome mask to pass to the icon creation function.
        ' (This is a stupid Windows workaround for 32bpp cursors.  The cursor creation function always assumes
        '  the presence of a mask bitmap, so we have to submit one even if we want the PNG's alpha channel
        '  used for transparency!)
        Dim monoBmp As Long
        monoBmp = CreateBitmap(resDIB.GetDIBWidth, resDIB.GetDIBHeight, 1, 1, ByVal 0&)
        
        'Create an icon header and point it at our temp mask bitmap and our PNG resource
        Dim icoInfo As ICONINFO
        With icoInfo
            .fIcon = False
            .xHotspot = FixDPI(curHotspotX)
            .yHotspot = FixDPI(curHotspotY)
            .hbmMask = monoBmp
            .hbmColor = resDIB.GetDIBHandle
        End With
        
        'Create the cursor
        CreateCursorFromResource = CreateNewIcon(icoInfo)
        
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
Public Sub InitializeCursors()

    ReDim customCursorHandles(0) As Long

    'Previously, system cursors were cached here.  This is no longer needed per https://github.com/tannerhelland/PhotoDemon/issues/78
    ' I am leaving this sub in case I need to pre-load tool cursors in the future.
    
    'Note that UnloadAllCursors below is still required, as the program may dynamically generate custom cursors while running, and
    ' these cursors will not be automatically deleted by the system.

End Sub

'Unload any custom cursors from memory
Public Sub UnloadAllCursors()
    
    If (numOfCustomCursors = 0) Then Exit Sub
    
    Dim i As Long
    For i = 0 To numOfCustomCursors - 1
        DestroyCursor customCursorHandles(i)
    Next i
    
End Sub

'Use any 32bpp PNG resource as a cursor .  When setting the mouse pointer of VB objects, please use setPNGCursorToObject, below.
Public Sub SetPNGCursorToHwnd(ByVal dstHwnd As Long, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    SetClassLong dstHwnd, GCL_HCURSOR, RequestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Use any 32bpp PNG resource as a cursor.  Use this function preferentially over the previous one, "SetPNGCursorToHwnd", when possible.
' (If a VB object does not have its MousePointer property set to "custom", it will override our attempts to set a custom mouse icon.)
Public Sub SetPNGCursorToObject(ByRef srcObject As Object, ByVal pngTitle As String, Optional ByVal curHotspotX As Long = 0, Optional ByVal curHotspotY As Long = 0)
    srcObject.MousePointer = vbCustom
    SetClassLong srcObject.hWnd, GCL_HCURSOR, RequestCustomCursor(pngTitle, curHotspotX, curHotspotY)
End Sub

'Set a single object to use the hand cursor
Public Sub SetHandCursor(ByRef tControl As Object)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 99
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub SetHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND)
End Sub

Public Sub SetArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'Set a single form to use the arrow cursor
Public Sub SetArrowCursor(ByRef tControl As Object)
    tControl.MousePointer = vbCustom
    SetClassLong tControl.hWnd, GCL_HCURSOR, LoadCursor(0, IDC_ARROW)
End Sub

'If a custom PNG cursor has not been loaded, this function will load the PNG, convert it to cursor format, then store
' the cursor resource for future reference (so the image doesn't have to be loaded again).
Public Function RequestCustomCursor(ByVal CursorName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0) As Long

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
        RequestCustomCursor = customCursorHandles(cursorLocation)
    Else
        Dim tmpHandle As Long
        tmpHandle = CreateCursorFromResource(CursorName, cursorHotspotX, cursorHotspotY)
        
        If (tmpHandle <> 0) Then
            ReDim Preserve customCursorNames(0 To numOfCustomCursors) As String
            ReDim Preserve customCursorHandles(0 To numOfCustomCursors) As Long
            customCursorNames(numOfCustomCursors) = CursorName
            customCursorHandles(numOfCustomCursors) = tmpHandle
            numOfCustomCursors = numOfCustomCursors + 1
        End If
        
        RequestCustomCursor = tmpHandle
    End If

End Function

'Given an image in the .exe's resource section (typically a PNG image), load it to a pdDIB object.
' The calling function is responsible for deleting the DIB once they are done with it.
Public Function LoadResourceToDIB(ByVal resTitle As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Long = 0, Optional ByVal useCustomColor As Long = -1) As Boolean
        
    'As of v7.0, PD now has two places from which to pull resources:
    ' 1) Its own custom resource handler (which is the preferred location)
    ' 2) The old, standard .exe resource section (which is deprecated, and in the process of being removed)
    '
    'We always attempt (1) before falling back to (2).  The goal for 7.0's release is to remove (2) entirely.
        
    'Some functions may call this before GDI+ has loaded; exit if that happens
    If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then
    
        'Attempt the default resource manager first
        Dim intResFound As Boolean: intResFound = False
        If (Not g_Resources Is Nothing) Then
            If g_Resources.AreResourcesAvailable Then
            
                'Attempt to load the requested resource.  (This may fail, as I am still in the process of migrating
                ' all resources to the new format.)
                intResFound = g_Resources.LoadImageResource(resTitle, dstDIB, desiredWidth, desiredHeight, desiredBorders, , useCustomColor)
                LoadResourceToDIB = intResFound
            
            End If
        End If
        
        'If we failed to load the resource data using our internal methods, fall back to the old system
        If (Not intResFound) Then
            
            On Error GoTo Err_ResNotFound
            
            'Start by extracting the resource data (typically a PNG) into a bytestream
            Dim ImageData() As Byte
            ImageData() = LoadResData(resTitle, "CUSTOM")
            
            Dim IStream As IUnknown
            Set IStream = VB_Hacks.GetStreamFromVBArray(VarPtr(ImageData(0)), UBound(ImageData) - LBound(ImageData) + 1)
            
            If (Not IStream Is Nothing) Then
                
                'Use GDI+ to convert the bytestream into a usable image
                ' (Note that GDI+ will have been initialized already, as part of the core PhotoDemon startup routine)
                Dim gdipBitmap As Long
                If (GdipLoadImageFromStream(IStream, gdipBitmap) = 0) Then
                
                    'Retrieve the image's size and pixel format
                    Dim tmpRect As RECTF
                    GdipGetImageBounds gdipBitmap, tmpRect, UnitPixel
                    
                    Dim gdiPixelFormat As Long
                    GdipGetImagePixelFormat gdipBitmap, gdiPixelFormat
                    
                    'Create the DIB anew as necessary
                    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
                    
                    'If the image has an alpha channel, create a 32bpp DIB to receive it
                    If (gdiPixelFormat And PixelFormatAlpha <> 0) Or (gdiPixelFormat And PixelFormatPAlpha <> 0) Then
                        dstDIB.CreateBlank tmpRect.Width, tmpRect.Height, 32
                        dstDIB.SetInitialAlphaPremultiplicationState True
                    Else
                        dstDIB.CreateBlank tmpRect.Width, tmpRect.Height, 24
                    End If
                    
                    'Convert the GDI+ bitmap to a standard Windows hBitmap
                    Dim hBitmap As Long
                    If (GdipCreateHBITMAPFromBitmap(gdipBitmap, hBitmap, vbBlack) = 0) Then
                    
                        'Select the hBitmap into a new DC so we can BitBlt it into the target DIB
                        Dim gdiDC As Long
                        gdiDC = GDI.GetMemoryDC()
                        
                        Dim oldBitmap As Long
                        oldBitmap = SelectObject(gdiDC, hBitmap)
                        
                        'Copy the GDI+ bitmap into the DIB
                        BitBlt dstDIB.GetDIBDC, 0, 0, tmpRect.Width, tmpRect.Height, gdiDC, 0, 0, vbSrcCopy
                        
                        'Release the original DDB and temporary device context
                        SelectObject gdiDC, oldBitmap
                        DeleteObject hBitmap
                        GDI.FreeMemoryDC gdiDC
                        
                        'As an added bonus, free the destination DIB from its DC as well.  (pdDIB objects automatically
                        ' select themselves into a DC, as necessary, so if this DIB isn't needed right away, we can
                        ' spare usage of a DC until it actually needs to be rendered.)
                        dstDIB.FreeFromDC
                        
                        LoadResourceToDIB = True
                        
                    Else
                        LoadResourceToDIB = False
                        Debug.Print "GDI+ failed to create an HBITMAP for requested resource " & resTitle & " stream."
                    End If
                    
                    'Release the GDI+ bitmap
                    GdipDisposeImage gdipBitmap
                        
                Else
                    LoadResourceToDIB = False
                    Debug.Print "GDI+ failed to load requested resource " & resTitle & " from stream."
                End If
            
                'Free the memory stream
                Set IStream = Nothing
                
            Else
                LoadResourceToDIB = False
                Debug.Print "Could not load requested resource " & resTitle & " from file."
            End If
            
        End If
        
    Else
        'Debug.Print "GDI+ unavailable; resources suspended for this session."
        LoadResourceToDIB = False
    End If
    
    Exit Function
    
Err_ResNotFound:

    Debug.Print "Requested resource " & resTitle & " wasn't found."
    LoadResourceToDIB = False
        
End Function

'PD will automatically update its taskbar icon to reflect the current image being edited.  I find this especially helpful
' when multiple PD sessions are operating in parallel.
Public Sub ChangeAppIcons(ByVal hIconSmall As Long, ByVal hIconLarge As Long)
    
    If (Not ALLOW_DYNAMIC_ICONS) Then Exit Sub
    Dim oldHIconL As Long, oldHIconS As Long
    oldHIconS = SendMessageA(FormMain.hWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall)
    oldHIconL = SendMessageA(FormMain.hWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge)
    
    'Generally speaking, you want to destroy the old icons after a change, but we track (and manage)
    ' these values internally, so there's no need to destroy icons at WM_SETICON time.
    'If (oldHIconS <> 0) Then DestroyIcon oldHIconS
    'If (oldHIconL <> 0) Then DestroyIcon oldHIconL
    
End Sub

'When loading a modal dialog, the dialog will not have an icon by default.  We can assign an icon at run-time to ensure that icons
' appear in the Alt+Tab dialog of older OSes.
Public Sub ChangeWindowIcon(ByVal targetHwnd As Long, ByVal hIconSmall As Long, ByVal hIconLarge As Long, Optional ByRef dstSmallIcon As Long = 0, Optional ByRef dstLargeIcon As Long = 0)
    If (targetHwnd <> 0) Then
        dstSmallIcon = SendMessageA(targetHwnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall)
        dstLargeIcon = SendMessageA(targetHwnd, WM_SETICON, ICON_BIG, ByVal hIconLarge)
    End If
End Sub

Public Sub MirrorCurrentIconsToWindow(ByVal targetHwnd As Long, Optional ByVal setLargeIconOnly As Boolean = False, Optional ByRef dstSmallIcon As Long = 0, Optional ByRef dstLargeIcon As Long = 0)
    If (g_OpenImageCount > 0) Then
        ChangeWindowIcon targetHwnd, IIf(setLargeIconOnly, 0&, pdImages(g_CurrentImage).curFormIcon16), pdImages(g_CurrentImage).curFormIcon32, dstSmallIcon, dstLargeIcon
    Else
        ChangeWindowIcon targetHwnd, IIf(setLargeIconOnly, 0&, m_DefaultIconSmall), m_DefaultIconLarge, dstSmallIcon, dstLargeIcon
    End If
End Sub

'When all images are unloaded (or when the program is first loaded), we must reset the program icon to its default values.
Public Sub ResetAppIcons()
    
    If (m_DefaultIconLarge = 0) Then
        m_DefaultIconLarge = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_SHARED)
    End If
    
    If (m_DefaultIconSmall = 0) Then
        m_DefaultIconSmall = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_SHARED)
    End If
    
    ChangeAppIcons m_DefaultIconSmall, m_DefaultIconLarge
    
End Sub

'When PD is first loaded, we associate an icon with the master "ThunderMain" owner window, to ensure proper icons in places
' like Task Manager.
Public Sub SetThunderMainIcon()

    'Start by loading the default icons from the resource file, as necessary
    ResetAppIcons
    
    Dim tmHWnd As Long
    tmHWnd = VB_Hacks.GetThunderMainHWnd()
    SendMessageA tmHWnd, WM_SETICON, ICON_SMALL, ByVal m_DefaultIconLarge
    SendMessageA tmHWnd, WM_SETICON, ICON_BIG, ByVal m_DefaultIconSmall

End Sub
