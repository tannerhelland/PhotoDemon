Attribute VB_Name = "Menu_Icon_Handler"
'***************************************************************************
'Specialized Icon Handler
'Copyright ©2011-2013 by Tanner Helland
'Created: 24/June/12
'Last updated: 12/August/12
'Last update: Added ResetMenuIcons, which redraws menu icons that may have been dropped due to the menu
'             caption changing (necessary for Undo/Redo text updating)
'
'Because VB6 doesn't provide many mechanisms for working with icons, I've had to manually add a number of
' icon-related functions to PhotoDemon.  First is a way to add icons/bitmaps to menus, as originally written
' by Leandro Ascierto.  Menu icons are extracted from a resource file (where they're stored in PNG format) and
' rendered to the menu at run-time.  See the clsMenuImage class for details on how this works.
' (A link to Leandro's original project can also be found there.)
'
'NOTE: Because the Windows XP version of Leandro's code utilizes potentially dirty subclassing,
' PhotoDemon automatically disables menu icons while running in the IDE on Windows XP.  Compile the project to see icons.
' (Windows Vista and 7 use a different mechanism, so menu icons are enabled in the IDE, and menu icons appear on all
' versions of Windows when compiled.)
'
'Also in this module is a heavily modified version of Paul Turcksin's "Icon Handlemaker" project, which I've modified
' to convert bitmaps to icons on the fly (the "CreateCustomFormIcon" sub).  PhotoDemon uses this to dynamically change
' the icon of its MDI child forms.  To see Paul's original project, please visit this PSC link:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=60600&lngWId=1
'
'***************************************************************************

Option Explicit

'API calls for building an icon at run-time
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Types required by the above API calls
Private Type Bitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Type ICONINFO
   fIcon As Boolean
   xHotspot As Long
   yHotspot As Long
   hbmMask As Long
   hbmColor As Long
End Type

'This array will be used to store our dynamically created icon handles so we can delete them on program exit
Dim numOfIcons As Long
Dim iconHandles() As Long

'This constant is used for testing only.  It should always be set to TRUE for production code.
Public Const ALLOW_DYNAMIC_ICONS As Boolean = True

'The types and constants below (commented out) can be used to generate an icon object for use within VB

'Private Type Guid
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(7) As Byte
'   End Type

'Private Type pictDesc
'   cbSizeofStruct As Long
'   picType As Long
'   hImage As Long
'End Type

'Constants required by the icon-related API calls
'Private Const PICTYPE_BITMAP = 1
'Private Const PICTYPE_ICON = 3

'These arrays will track the resource identifiers and consequent numeric identifiers of all loaded icons.  The size of the array
' is arbitrary; just make sure it's larger than the max number of loaded icons.
Private iconNames(0 To 255) As String

'We also need to track how many icons have been loaded; this counter will also be used to reference icons in the database
Dim curIcon As Long

'API call for manually setting a 32-bit icon to a form (as opposed to Form.Icon = ...)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'clsMenuImage does the heavy lifting for inserting icons into menus
Dim cMenuImage As clsMenuImage

'A second class is used to manage the icons for the MRU list.
Dim cMRUIcons As clsMenuImage

'Load all the menu icons from PhotoDemon's embedded resource file
Public Sub LoadMenuIcons()

    Set cMenuImage = New clsMenuImage
    
    With cMenuImage
            
        'Use Leandro's class to check if the current Windows install supports theming.
        g_IsThemingEnabled = .CanWeTheme
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not g_IsVistaOrLater) And (g_IsProgramCompiled = False) Then Exit Sub
        
        .Init FormMain.hWnd, 16, 16
        
    End With
    
    curIcon = 0
        
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

'This new, simpler technique for adding menu icons requires only the menu location (including sub-menus) and the icon's identifer
' in the resource file.  If the icon has already been loaded, it won't be loaded again; instead, the function will check the list
' of loaded icons and automatically fill in the numeric identifier as necessary.
Private Sub AddMenuIcon(ByVal resID As String, ByVal topMenu As Long, ByVal subMenu As Long, Optional ByVal subSubMenu As Long = -1)

    Static i As Long
    Static iconLocation As Long
    Static iconAlreadyLoaded As Boolean
    
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
    AddMenuIcon "UNDO", 1, 0        'Undo
    AddMenuIcon "REDO", 1, 1        'Redo
    AddMenuIcon "REPEAT", 1, 2      'Repeat Last Action
    AddMenuIcon "COPY", 1, 4        'Copy
    AddMenuIcon "PASTE", 1, 5       'Paste
    AddMenuIcon "CLEAR", 1, 6       'Empty Clipboard
    
    'View Menu
    AddMenuIcon "FITWINIMG", 2, 0     'Fit Viewport to Image
    AddMenuIcon "FITONSCREEN", 2, 1   'Fit on Screen
    AddMenuIcon "ZOOMIN", 2, 3        'Zoom In
    AddMenuIcon "ZOOMOUT", 2, 4       'Zoom Out
    AddMenuIcon "ZOOMACTUAL", 2, 10   'Zoom 100%
    AddMenuIcon "LEFTPANSHOW", 2, 16  'Show/Hide the left-hand panel
    
    'Image Menu
    AddMenuIcon "DUPLICATE", 3, 0      'Duplicate
    AddMenuIcon "MODE24", 3, 2         'Image Mode
        '--> Image Mode sub-menu
        AddMenuIcon "MODE24", 3, 2, 0  '24bpp
        AddMenuIcon "MODE32", 3, 2, 1  '32bpp
    AddMenuIcon "RESIZE", 3, 4         'Resize
    AddMenuIcon "CROPSEL", 3, 5        'Crop to Selection
    AddMenuIcon "MIRROR", 3, 7         'Mirror
    AddMenuIcon "FLIP", 3, 8           'Flip
    AddMenuIcon "ROTATECW", 3, 10      'Rotate Clockwise
    AddMenuIcon "ROTATECCW", 3, 11     'Rotate Counter-clockwise
    AddMenuIcon "ROTATE180", 3, 12     'Rotate 180
    'NOTE: the specific menu values will be different if the FreeImage plugin (FreeImage.dll) isn't found.
    If g_ImageFormats.FreeImageEnabled Then
        AddMenuIcon "ROTATEANY", 3, 13 'Rotate Arbitrary
        AddMenuIcon "ISOMETRIC", 3, 15 'Isometric
        AddMenuIcon "TILE", 3, 16      'Tile
    Else
        AddMenuIcon "ISOMETRIC", 3, 14 'Isometric
        AddMenuIcon "TILE", 3, 15      'Tile
    End If
    
    'Color Menu
    AddMenuIcon "BRIGHT", 4, 0      'Brightness/Contrast
    AddMenuIcon "GAMMA", 4, 1      'Gamma Correction
    AddMenuIcon "HSL", 4, 2     'HSL adjustment
    AddMenuIcon "LEVELS", 4, 3      'Levels
    AddMenuIcon "TEMPERATURE", 4, 4     'Temperature
    AddMenuIcon "WHITEBAL", 4, 5      'White Balance
    AddMenuIcon "HISTOGRAM", 4, 7      'Histogram
        '--> Histogram sub-menu
        AddMenuIcon "HISTOGRAM", 4, 7, 0  'Display Histogram
        AddMenuIcon "EQUALIZE", 4, 7, 2   'Equalize
        AddMenuIcon "STRETCH", 4, 7, 3  'Stretch
    AddMenuIcon "COLORSHIFTR", 4, 9      'Color Shift
        '--> Color-Shift sub-menu
        AddMenuIcon "COLORSHIFTR", 4, 9, 0  'Shift Right
        AddMenuIcon "COLORSHIFTL", 4, 9, 1  'Shift Left
    AddMenuIcon "RECHANNELB", 4, 10      'Rechannel
    AddMenuIcon "BLACKWHITE", 4, 12      'Black and White
    AddMenuIcon "COLORIZE", 4, 13     'Colorize
    AddMenuIcon "ENHANCE", 4, 14     'Enhance
        '--> Enhance sub-menu
        AddMenuIcon "ENCONTRAST", 4, 14, 0  'Contrast
        AddMenuIcon "ENHIGHLIGHT", 4, 14, 1   'Highlights
        AddMenuIcon "ENMIDTONE", 4, 14, 2  'Midtones
        AddMenuIcon "ENSHADOW", 4, 14, 3  'Shadows
    AddMenuIcon "FADE", 4, 15     'Fade
        '--> Fade sub-menu
        AddMenuIcon "FADELOW", 4, 15, 0 'Low Fade
        AddMenuIcon "FADE", 4, 15, 1 'Medium Fade
        AddMenuIcon "FADEHIGH", 4, 15, 2 'High Fade
        AddMenuIcon "CUSTOMFADE", 4, 15, 3 'Custom Fade
        AddMenuIcon "UNFADE", 4, 15, 5 'Unfade
    AddMenuIcon "GRAYSCALE", 4, 16      'Grayscale
    AddMenuIcon "INVERT", 4, 17     'Invert
        '--> Invert sub-menu
        AddMenuIcon "INVCMYK", 4, 17, 0  'Invert CMYK
        AddMenuIcon "INVHUE", 4, 17, 1  'Invert Hue
        AddMenuIcon "INVRGB", 4, 17, 2  'Invert RGB
        AddMenuIcon "INVCOMPOUND", 4, 17, 4 'Compound Invert
    AddMenuIcon "POSTERIZE", 4, 18      'Posterize
    AddMenuIcon "SEPIA", 4, 19    'Sepia
    AddMenuIcon "COUNTCOLORS", 4, 21     'Count Colors
    AddMenuIcon "REDUCECOLORS", 4, 22      'Reduce Colors
    
    'Filters Menu
    AddMenuIcon "FADELAST", 5, 0       'Fade Last
    AddMenuIcon "ARTISTIC", 5, 2       'Artistic
        '--> Artistic sub-menu
        AddMenuIcon "ANTIQUE", 5, 2, 0   'Antique (Sepia)
        AddMenuIcon "COMICBOOK", 5, 2, 1   'Comic Book
        AddMenuIcon "PENCIL", 5, 2, 2   'Pencil
        AddMenuIcon "MOSAIC", 5, 2, 3   'Pixelate (Mosaic)
        AddMenuIcon "RELIEF", 5, 2, 4   'Relief
    AddMenuIcon "BLUR", 5, 3      'Blur
        '--> Blur sub-menu
        AddMenuIcon "ANTIALIAS", 5, 3, 0   'Antialias
        AddMenuIcon "SOFTEN", 5, 3, 2   'Soften
        AddMenuIcon "SOFTENMORE", 5, 3, 3   'Soften More
        AddMenuIcon "BLUR2", 5, 3, 4   'Blur
        AddMenuIcon "BLURMORE", 5, 3, 5   'Blur More
        AddMenuIcon "GAUSSBLUR", 5, 3, 6   'Gaussian Blur
        AddMenuIcon "GAUSSBLURMOR", 5, 3, 7   'Gaussian Blur More
        AddMenuIcon "GRIDBLUR", 5, 3, 9   'Grid Blur
    'AddMenuIcon "DISTORT", 5, 4       'Distort
        '--> Distort sub-menu
    AddMenuIcon "EDGES", 5, 5       'Edges
        '--> Edges sub-menu
        AddMenuIcon "EMBOSS", 5, 5, 0  'Emboss / Engrave
        AddMenuIcon "EDGEENHANCE", 5, 5, 1   'Enhance Edges
        AddMenuIcon "EDGES", 5, 5, 2   'Find Edges
    AddMenuIcon "OTHER", 5, 6      'Fun
        '--> Fun sub-menu
        AddMenuIcon "ALIEN", 5, 6, 0  'Alien
        AddMenuIcon "BLACKLIGHT", 5, 6, 1  'Blacklight
        AddMenuIcon "DREAM", 5, 6, 2  'Dream
        AddMenuIcon "RADIOACTIVE", 5, 6, 3  'Radioactive
        AddMenuIcon "SYNTHESIZE", 5, 6, 4  'Synthesize
        AddMenuIcon "HEATMAP", 5, 6, 5  'Thermograph
        AddMenuIcon "VIBRATE", 5, 6, 6  'Vibrate
    AddMenuIcon "NATURAL", 5, 7       'Natural
        '--> Natural sub-menu
        AddMenuIcon "ATMOSPHERE", 5, 7, 0  'Atmosphere
        AddMenuIcon "BURN", 5, 7, 1  'Burn
        AddMenuIcon "FOG", 5, 7, 2  'Fog
        AddMenuIcon "FREEZE", 5, 7, 3  'Freeze
        AddMenuIcon "LAVA", 5, 7, 4  'Lava
        'AddMenuIcon "OCEAN", 5, 7, 5  'Ocean
        AddMenuIcon "RAINBOW", 5, 7, 5  'Rainbow
        AddMenuIcon "STEEL", 5, 7, 6  'Steel
        AddMenuIcon "WATER", 5, 7, 7  'Water
    AddMenuIcon "NOISE", 5, 8       'Noise
        '--> Noise sub-menu
        AddMenuIcon "ADDNOISE", 5, 8, 0  'Add Noise
        AddMenuIcon "DESPECKLE", 5, 8, 2  'Despeckle
        AddMenuIcon "REMOVEORPHAN", 5, 8, 3  'Remove Orphan
    AddMenuIcon "RANK", 5, 9       'Rank
        '--> Rank sub-menu
        AddMenuIcon "DILATE", 5, 9, 0  'Dilate
        AddMenuIcon "ERODE", 5, 9, 1  'Erode
        AddMenuIcon "EXTREME", 5, 9, 2  'Extreme
        AddMenuIcon "CUSTRANK", 5, 9, 4  'Custom Rank
    AddMenuIcon "SHARPEN", 5, 10       'Sharpen
        '--> Sharpen sub-menu
        AddMenuIcon "UNSHARP", 5, 10, 0 'Unsharp
        AddMenuIcon "SHARPEN", 5, 10, 2  'Sharpen
        AddMenuIcon "SHARPENMORE", 5, 10, 3 'Sharpen More
    AddMenuIcon "STYLIZE", 5, 11      'Stylize
        '--> Stylize sub-menu
        AddMenuIcon "DIFFUSE", 5, 11, 0  'Diffuse
        AddMenuIcon "SOLARIZE", 5, 11, 1 'Solarize
        AddMenuIcon "TWINS", 5, 11, 2 'Twins
    AddMenuIcon "CUSTFILTER", 5, 13      'Custom Filter
    
    'Tools Menu
    AddMenuIcon "RECORD", 6, 0       'Macros
        '--> Macro sub-menu
        AddMenuIcon "OPENMACRO", 6, 0, 0    'Open Macro
        AddMenuIcon "RECORD", 6, 0, 2   'Start Recording
        AddMenuIcon "RECORDSTOP", 6, 0, 3    'Stop Recording
    AddMenuIcon "PREFERENCES", 6, 2      'Options (Preferences)
    AddMenuIcon "PLUGIN", 6, 3     'Plugin Manager
    
    'Window Menu
    AddMenuIcon "NEXTIMAGE", 7, 0    'Next image
    AddMenuIcon "PREVIMAGE", 7, 1    'Previous image
    AddMenuIcon "ARNGICONS", 7, 3     'Arrange Icons
    AddMenuIcon "CASCADE", 7, 4     'Cascade
    AddMenuIcon "TILEVER", 7, 5     'Tile Horizontally
    AddMenuIcon "TILEHOR", 7, 6     'Tile Vertically
    AddMenuIcon "MINALL", 7, 8     'Minimize All
    AddMenuIcon "RESTOREALL", 7, 9     'Restore All
    
    'Help Menu
    AddMenuIcon "FAVORITE", 8, 0     'Donate
    AddMenuIcon "UPDATES", 8, 2     'Check for updates
    AddMenuIcon "FEEDBACK", 8, 3     'Submit Feedback
    AddMenuIcon "BUG", 8, 4     'Submit Bug
    AddMenuIcon "PDWEBSITE", 8, 6     'Visit the PhotoDemon website
    AddMenuIcon "DOWNLOADSRC", 8, 7    'Download source code
    AddMenuIcon "LICENSE", 8, 8    'License
    AddMenuIcon "ABOUT", 8, 10    'About PD
    
End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' Note that to keep the code small, all changeable icons are refreshed whenever this is called.
Public Sub ResetMenuIcons()

    'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
    If (Not g_IsVistaOrLater) And (g_IsProgramCompiled = False) Then Exit Sub

    'The position of menus changes if the MDI child is maximized.  When maximized, the form menu is given index 0, shifting
    ' everything to the right by one.
    
    'Thus, we must check for that and redraw the Undo/Redo menus accordingly
    Dim posModifier As Long
    posModifier = 0

    If NumOfWindows > 0 Then
        If FormMain.ActiveForm.WindowState = vbMaximized Then posModifier = 1
    End If
        
    'Redraw the Undo/Redo menus
    With cMenuImage
        AddMenuIcon "UNDO", 1 + posModifier, 0     'Undo
        AddMenuIcon "REDO", 1 + posModifier, 1     'Redo
    End With
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    Dim numOfMRUFiles As Long
    numOfMRUFiles = MRU_ReturnCount()
    AddMenuIcon "CLEARRECENT", 0 + posModifier, 1, numOfMRUFiles + 1
    
    'Change the Show/Hide left icon panel to match
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideLeftPanel", False) Then
        AddMenuIcon "LEFTPANSHOW", 2 + posModifier, 16    'Show the panel
    Else
        AddMenuIcon "LEFTPANHIDE", 2 + posModifier, 16     'Hide the panel
    End If
        
    'If the OS is Vista or later, render MRU icons to the Open Recent menu
    If g_IsVistaOrLater Then
    
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

'As a courtesy to the user, the Image -> Mode icon is dynamically changed to match the current image's mode
Public Sub updateModeIcon(ByVal nMode As Boolean)
    
    If nMode Then
        'Update the parent menu
        AddMenuIcon "MODE32", 3, 2
        
        'Update the children menus as well
        AddMenuIcon "MODE24", 3, 2, 0   '24bpp
        AddMenuIcon "MODE32CHK", 3, 2, 1    '32bpp
        
    Else
        AddMenuIcon "MODE24", 3, 2
        AddMenuIcon "MODE24CHK", 3, 2, 0
        AddMenuIcon "MODE32", 3, 2, 1
    End If
    
End Sub

'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
'Again, thanks to Paul Turcksin for the original draft of this code.
Public Sub CreateCustomFormIcon(ByRef imgForm As FormImage)

    If Not ALLOW_DYNAMIC_ICONS Then Exit Sub

    'Generating an icon requires many variables; see below for specific comments on each one
    Dim BitmapData As Bitmap
    Dim iWidth As Long
    Dim iHeight As Long
    Dim srcDC As Long
    Dim oldSrcObj As Long
    Dim MonoDC As Long
    Dim MonoBmp As Long
    Dim oldMonoObj As Long
    Dim InvertDC As Long
    Dim InvertBmp As Long
    Dim oldInvertObj As Long
    Dim cBkColor As Long
    Dim maskClr As Long
    Dim icoInfo As ICONINFO
    Dim generatedIcon As Long
   
    'The icon can be drawn at any size, but 16x16 is how it will (typically) end up on the form.  Since we are now rendering
    ' a dynamically generated icon to the task bar as well, we opt for 32x32, and from that we generate an additional 16x16 version.
    Dim icoSize As Long
    
    'If we are rendering a dynamic taskbar icon, we will perform two reductions - first to 32x32, second to 16x16
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) Then icoSize = 32 Else icoSize = 16

    'Determine aspect ratio
    Dim aspectRatio As Single
    aspectRatio = CSng(pdImages(imgForm.Tag).Width) / CSng(pdImages(imgForm.Tag).Height)
    
    'The target icon's width and height, x and y positioning
    Dim tIcoWidth As Single, tIcoHeight As Single, TX As Single, TY As Single
    
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
    
    'I have two systems in place for rendering the dynamic form icons.  One relies on FreeImage, and generates a very high-quality,
    ' 32bpp with full alpha icon.  This is obviously the preferred method.  If FreeImage cannot be found, a StretchBlt-based
    ' technique is used.  Alpha is not taken into account by that method (obviously), and the icon quality is much worse.
    If g_ImageFormats.FreeImageEnabled Then
    
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
                MonoDC = CreateCompatibleDC(0&)
                MonoBmp = CreateCompatibleBitmap(MonoDC, icoSize, icoSize)
                oldMonoObj = SelectObject(MonoDC, MonoBmp)
            
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
                DeleteObject icoInfo.hbmMask
                DeleteObject icoInfo.hbmColor
                DeleteDC MonoDC
                
                'Finally, resize the 32x32 icon to 16x16 so it will work as the current form icon as well
                icoSize = 16
                finalDIB = FreeImage_RescaleByPixel(finalDIB, 16, 16, True, FILTER_BILINEAR)
                
            End If
            
            'Generate a blank monochrome mask to pass to the icon creation function.
            MonoDC = CreateCompatibleDC(0&)
            MonoBmp = CreateCompatibleBitmap(MonoDC, icoSize, icoSize)
            oldMonoObj = SelectObject(MonoDC, MonoBmp)
            
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
                        
        End If
        
    'If FreeImage isn't enabled, fall back to StretchBlt
    Else
    
        'Clear out the current picture box
        imgForm.picIcon.Picture = LoadPicture("")
    
        'Resize the picture box that will receive the first draft of the icon
        imgForm.picIcon.Width = icoSize
        imgForm.picIcon.Height = icoSize
    
        'Because we'll be shrinking the image dramatically, set StretchBlt to use resampling
        SetStretchBltMode imgForm.picIcon.hDC, STRETCHBLT_HALFTONE
    
        'Render the bitmap that will ultimately be converted into an icon
        StretchBlt imgForm.picIcon.hDC, CLng(TX), CLng(TY), CLng(tIcoWidth), CLng(tIcoHeight), pdImages(imgForm.Tag).mainLayer.getLayerDC, 0, 0, pdImages(imgForm.Tag).Width, pdImages(imgForm.Tag).Height, vbSrcCopy
        imgForm.picIcon.Picture = imgForm.picIcon.Image
        
        'Now that we have a first draft to work from, start preparing the data types required by the icon API calls
        GetObject imgForm.picIcon.Picture.Handle, Len(BitmapData), BitmapData

        With BitmapData
            iWidth = .bmWidth
            iHeight = .bmHeight
        End With
       
        'Create a copy of the original image; this will be used to generate a mask (necessary if the image isn't square-shaped)
        srcDC = CreateCompatibleDC(0&)
        oldSrcObj = SelectObject(srcDC, imgForm.picIcon.Picture.Handle)
       
        'If the image isn't square-shaped, the backcolor of the first draft image will need to be made transparent
        If tIcoWidth < icoSize Or tIcoHeight < icoSize Then
            maskClr = imgForm.picIcon.backColor
        Else
            maskClr = 0
        End If
       
        'Generate two masks. First, a monochrome mask.
        MonoDC = CreateCompatibleDC(0&)
        MonoBmp = CreateCompatibleBitmap(MonoDC, iWidth, iHeight)
        oldMonoObj = SelectObject(MonoDC, MonoBmp)
        cBkColor = GetBkColor(srcDC)
        SetBkColor srcDC, maskClr
        BitBlt MonoDC, 0, 0, iWidth, iHeight, srcDC, 0, 0, vbSrcCopy
        SetBkColor srcDC, cBkColor
    
        'Second, an AND mask
        InvertDC = CreateCompatibleDC(0&)
        InvertBmp = CreateCompatibleBitmap(imgForm.hDC, iWidth, iHeight)
        oldInvertObj = SelectObject(InvertDC, InvertBmp)
        BitBlt InvertDC, 0, 0, iWidth, iHeight, srcDC, 0, 0, vbSrcCopy
        SetBkColor InvertDC, vbBlack
        SetTextColor InvertDC, vbWhite
        BitBlt InvertDC, 0, 0, iWidth, iHeight, MonoDC, 0, 0, vbSrcAnd
  
        'We no longer need our copy of the original image, so free up that memory
        SelectObject srcDC, oldSrcObj
        DeleteDC srcDC
        
        'We can also free up the temporary DCs used to generate our two masks
        SelectObject MonoDC, oldMonoObj
        SelectObject InvertDC, oldInvertObj
        
        With icoInfo
            .fIcon = True
            .xHotspot = icoSize
            .yHotspot = icoSize
            .hbmMask = MonoBmp
            .hbmColor = InvertBmp
        End With
        
    End If
    
    'Render the icon to a handle
    generatedIcon = CreateIconIndirect(icoInfo)
        
    If g_ImageFormats.FreeImageEnabled Then
    
        'If we are dynamically updating the taskbar icon to match the current image, we need to assign the 16x16 icon now
        If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) Then
            setNewAppIcon generatedIcon
            pdImages(imgForm.Tag).curFormIcon16 = generatedIcon
        End If
    
        FreeImage_UnloadEx finalDIB
        FreeLibrary hLib
        DeleteObject icoInfo.hbmMask
        DeleteObject icoInfo.hbmColor
        DeleteDC MonoDC
    Else
        'Clear out our temporary masks (whose info are now embedded in the icon itself)
        DeleteObject icoInfo.hbmMask
        DeleteObject icoInfo.hbmColor
        DeleteDC MonoDC
        DeleteDC InvertDC
    End If
   
    'Use the API to assign this new icon to the specified MDI child form
    SendMessageLong imgForm.hWnd, &H80, 0, generatedIcon
        
    'Store this icon in our running list, so we can destroy it when the program is closed
    addIconToList generatedIcon

    'When an MDI child form is maximized, the icon is not updated properly. This requires further investigation to solve.
    'If imgForm.WindowState = vbMaximized Then DoEvents
   
    'The chunk of code below will generate an actual icon object for use within VB. I don't use this mechanism because
    ' VB will internally convert the icon to 256-colors before assigning it to the form <sigh>.  Rather than do that,
    ' I use an alternate API call above to assign the new icon in its transparent, full color glory.
    
    'Dim iGuid As Guid
    'With iGuid
    ' .Data1 = &H20400
    ' .Data4(0) = &HC0
    ' .Data4(7) = &H46
    'End With
    
    'Dim pDesc As pictDesc
    'With pDesc
    ' .cbSizeofStruct = Len(pDesc)
    ' .picType = PICTYPE_ICON
    ' .hImage = generatedIcon
    'End With
    
    'Dim icoObject As IPicture
    'OleCreatePictureIndirect pDesc, iGuid, 1, icoObject
    
    'imgForm.Icon = icoObject
   
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
