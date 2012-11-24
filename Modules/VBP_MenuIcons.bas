Attribute VB_Name = "Menu_Icon_Handler"
'***************************************************************************
'Specialized Icon Handler
'Copyright ©2011-2012 by Tanner Helland
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

'This array will be used to store our icon handles so we can delete them on program exit
Dim numOfIcons As Long
Dim iconHandles() As Long

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
    
        'Leandro's class automatically detects the current Windows version.  We're only concerned with Vista or later, which lets us
        ' know that certain features are guaranteed to be available.
        isVistaOrLater = .IsWindowVistaOrLater
        
        'Also, use it to check if the current Windows install supports theming.
        isThemingEnabled = .CanWeTheme
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not isVistaOrLater) And (IsProgramCompiled = False) Then Exit Sub
        
        .Init FormMain.hWnd, 16, 16
        
        .AddImageFromStream LoadResData("OPENIMG", "CUSTOM")     '0
        .AddImageFromStream LoadResData("OPENREC", "CUSTOM")     '1
        .AddImageFromStream LoadResData("IMPORT", "CUSTOM")      '2
        .AddImageFromStream LoadResData("SAVE", "CUSTOM")        '3
        .AddImageFromStream LoadResData("SAVEAS", "CUSTOM")      '4
        .AddImageFromStream LoadResData("CLOSE", "CUSTOM")       '5
        .AddImageFromStream LoadResData("BCONVERT", "CUSTOM")    '6
        .AddImageFromStream LoadResData("PRINT", "CUSTOM")       '7
        .AddImageFromStream LoadResData("SCANNER", "CUSTOM")     '8
        .AddImageFromStream LoadResData("DOWNLOAD", "CUSTOM")    '9
        .AddImageFromStream LoadResData("SCREENCAP", "CUSTOM")   '10
        .AddImageFromStream LoadResData("FRXIMPORT", "CUSTOM")   '11
        .AddImageFromStream LoadResData("UNDO", "CUSTOM")        '12
        .AddImageFromStream LoadResData("REDO", "CUSTOM")        '13
        .AddImageFromStream LoadResData("REPEAT", "CUSTOM")      '14
        .AddImageFromStream LoadResData("COPY", "CUSTOM")        '15
        .AddImageFromStream LoadResData("PASTE", "CUSTOM")       '16
        .AddImageFromStream LoadResData("CLEAR", "CUSTOM")       '17
        .AddImageFromStream LoadResData("PREFERENCES", "CUSTOM") '18
        .AddImageFromStream LoadResData("RESIZE", "CUSTOM")      '19
        .AddImageFromStream LoadResData("ROTATECW", "CUSTOM")    '20
        .AddImageFromStream LoadResData("ROTATECCW", "CUSTOM")   '21
        .AddImageFromStream LoadResData("ROTATE180", "CUSTOM")   '22
        .AddImageFromStream LoadResData("FLIP", "CUSTOM")        '23
        .AddImageFromStream LoadResData("MIRROR", "CUSTOM")      '24
        .AddImageFromStream LoadResData("PDWEBSITE", "CUSTOM")   '25
        .AddImageFromStream LoadResData("FEEDBACK", "CUSTOM")    '26
        .AddImageFromStream LoadResData("ABOUT", "CUSTOM")       '27
        .AddImageFromStream LoadResData("FITWINIMG", "CUSTOM")   '28
        .AddImageFromStream LoadResData("FITONSCREEN", "CUSTOM") '29
        .AddImageFromStream LoadResData("TILEHOR", "CUSTOM")     '30
        .AddImageFromStream LoadResData("TILEVER", "CUSTOM")     '31
        .AddImageFromStream LoadResData("CASCADE", "CUSTOM")     '32
        .AddImageFromStream LoadResData("ARNGICONS", "CUSTOM")   '33
        .AddImageFromStream LoadResData("MINALL", "CUSTOM")      '34
        .AddImageFromStream LoadResData("RESTOREALL", "CUSTOM")  '35
        .AddImageFromStream LoadResData("OPENMACRO", "CUSTOM")   '36
        .AddImageFromStream LoadResData("RECORD", "CUSTOM")      '37
        .AddImageFromStream LoadResData("RECORDSTOP", "CUSTOM")  '38
        .AddImageFromStream LoadResData("BUG", "CUSTOM")         '39
        .AddImageFromStream LoadResData("FAVORITE", "CUSTOM")    '40
        .AddImageFromStream LoadResData("UPDATES", "CUSTOM")     '41
        .AddImageFromStream LoadResData("DUPLICATE", "CUSTOM")   '42
        .AddImageFromStream LoadResData("EXIT", "CUSTOM")        '43
        .AddImageFromStream LoadResData("CLEARRECENT", "CUSTOM") '44
        .AddImageFromStream LoadResData("SCANNERSEL", "CUSTOM")  '45
        .AddImageFromStream LoadResData("BRIGHT", "CUSTOM")      '46
        .AddImageFromStream LoadResData("GAMMA", "CUSTOM")       '47
        .AddImageFromStream LoadResData("LEVELS", "CUSTOM")      '48
        .AddImageFromStream LoadResData("WHITEBAL", "CUSTOM")    '49
        .AddImageFromStream LoadResData("HISTOGRAM", "CUSTOM")   '50
        .AddImageFromStream LoadResData("EQUALIZE", "CUSTOM")    '51
        .AddImageFromStream LoadResData("STRETCH", "CUSTOM")     '52
        .AddImageFromStream LoadResData("COLORSHIFTR", "CUSTOM") '53
        .AddImageFromStream LoadResData("COLORSHIFTL", "CUSTOM") '54
        .AddImageFromStream LoadResData("RECHANNELR", "CUSTOM")  '55
        .AddImageFromStream LoadResData("RECHANNELG", "CUSTOM")  '56
        .AddImageFromStream LoadResData("RECHANNELB", "CUSTOM")  '57
        .AddImageFromStream LoadResData("BLACKWHITE", "CUSTOM")  '58
        .AddImageFromStream LoadResData("COLORIZE", "CUSTOM")    '59
        .AddImageFromStream LoadResData("ENHANCE", "CUSTOM")     '60
        .AddImageFromStream LoadResData("ENCONTRAST", "CUSTOM")  '61
        .AddImageFromStream LoadResData("ENHIGHLIGHT", "CUSTOM") '62
        .AddImageFromStream LoadResData("ENMIDTONE", "CUSTOM")   '63
        .AddImageFromStream LoadResData("ENSHADOW", "CUSTOM")    '64
        .AddImageFromStream LoadResData("GRAYSCALE", "CUSTOM")   '65
        .AddImageFromStream LoadResData("INVERT", "CUSTOM")      '66
        .AddImageFromStream LoadResData("INVCMYK", "CUSTOM")     '67
        .AddImageFromStream LoadResData("INVHUE", "CUSTOM")      '68
        .AddImageFromStream LoadResData("INVRGB", "CUSTOM")      '69
        .AddImageFromStream LoadResData("POSTERIZE", "CUSTOM")   '70
        .AddImageFromStream LoadResData("REDUCECOLORS", "CUSTOM") '71
        .AddImageFromStream LoadResData("COUNTCOLORS", "CUSTOM") '72
        .AddImageFromStream LoadResData("FADELAST", "CUSTOM")    '73
        .AddImageFromStream LoadResData("ARTISTIC", "CUSTOM")    '74
        .AddImageFromStream LoadResData("BLUR", "CUSTOM")        '75
        .AddImageFromStream LoadResData("DIFFUSE", "CUSTOM")     '76
        .AddImageFromStream LoadResData("EDGES", "CUSTOM")       '77
        .AddImageFromStream LoadResData("NATURAL", "CUSTOM")     '78
        .AddImageFromStream LoadResData("NOISE", "CUSTOM")       '79
        .AddImageFromStream LoadResData("OTHER", "CUSTOM")       '80
        .AddImageFromStream LoadResData("RANK", "CUSTOM")        '81
        .AddImageFromStream LoadResData("SHARPEN", "CUSTOM")     '82
        .AddImageFromStream LoadResData("ANTIQUE", "CUSTOM")     '83
        .AddImageFromStream LoadResData("COMICBOOK", "CUSTOM")   '84
        .AddImageFromStream LoadResData("MOSAIC", "CUSTOM")      '85
        .AddImageFromStream LoadResData("PENCIL", "CUSTOM")      '86
        .AddImageFromStream LoadResData("RELIEF", "CUSTOM")      '87
        .AddImageFromStream LoadResData("ANTIALIAS", "CUSTOM")   '88
        .AddImageFromStream LoadResData("SOFTEN", "CUSTOM")      '89
        .AddImageFromStream LoadResData("SOFTENMORE", "CUSTOM")  '90
        .AddImageFromStream LoadResData("BLUR2", "CUSTOM")       '91
        .AddImageFromStream LoadResData("BLURMORE", "CUSTOM")    '92
        .AddImageFromStream LoadResData("GAUSSBLUR", "CUSTOM")   '93
        .AddImageFromStream LoadResData("GAUSSBLURMOR", "CUSTOM") '94
        .AddImageFromStream LoadResData("GRIDBLUR", "CUSTOM")    '95
        .AddImageFromStream LoadResData("DIFFUSEMORE", "CUSTOM") '96
        .AddImageFromStream LoadResData("DIFFUSECUST", "CUSTOM") '97
        .AddImageFromStream LoadResData("CUSTFILTER", "CUSTOM")  '98
        .AddImageFromStream LoadResData("EDGEENHANCE", "CUSTOM") '99
        .AddImageFromStream LoadResData("EMBOSS", "CUSTOM")      '100
        .AddImageFromStream LoadResData("INVCOMPOUND", "CUSTOM") '101
        .AddImageFromStream LoadResData("FADE", "CUSTOM")        '102
        .AddImageFromStream LoadResData("FADELOW", "CUSTOM")     '103
        .AddImageFromStream LoadResData("FADEHIGH", "CUSTOM")    '104
        .AddImageFromStream LoadResData("CUSTOMFADE", "CUSTOM")  '105
        .AddImageFromStream LoadResData("UNFADE", "CUSTOM")      '106
        .AddImageFromStream LoadResData("ATMOSPHERE", "CUSTOM")  '107
        .AddImageFromStream LoadResData("BURN", "CUSTOM")        '108
        .AddImageFromStream LoadResData("FOG", "CUSTOM")         '109
        .AddImageFromStream LoadResData("FREEZE", "CUSTOM")      '110
        .AddImageFromStream LoadResData("LAVA", "CUSTOM")        '111
        .AddImageFromStream LoadResData("OCEAN", "CUSTOM")       '112
        .AddImageFromStream LoadResData("RAINBOW", "CUSTOM")     '113
        .AddImageFromStream LoadResData("STEEL", "CUSTOM")       '114
        .AddImageFromStream LoadResData("WATER", "CUSTOM")       '115
        .AddImageFromStream LoadResData("ADDNOISE", "CUSTOM")    '116
        .AddImageFromStream LoadResData("DESPECKLE", "CUSTOM")   '117
        .AddImageFromStream LoadResData("REMOVEORPHAN", "CUSTOM") '118
        .AddImageFromStream LoadResData("DILATE", "CUSTOM")      '119
        .AddImageFromStream LoadResData("ERODE", "CUSTOM")       '120
        .AddImageFromStream LoadResData("EXTREME", "CUSTOM")     '121
        .AddImageFromStream LoadResData("CUSTRANK", "CUSTOM")    '122
        .AddImageFromStream LoadResData("SHARPENMORE", "CUSTOM") '123
        .AddImageFromStream LoadResData("UNSHARP", "CUSTOM")     '124
        .AddImageFromStream LoadResData("ISOMETRIC", "CUSTOM")   '125
        .AddImageFromStream LoadResData("ALIEN", "CUSTOM")       '126
        .AddImageFromStream LoadResData("BLACKLIGHT", "CUSTOM")  '127
        .AddImageFromStream LoadResData("DREAM", "CUSTOM")       '128
        .AddImageFromStream LoadResData("RADIOACTIVE", "CUSTOM") '129
        .AddImageFromStream LoadResData("SOLARIZE", "CUSTOM")    '130
        .AddImageFromStream LoadResData("SYNTHESIZE", "CUSTOM")  '131
        .AddImageFromStream LoadResData("TILE", "CUSTOM")        '132
        .AddImageFromStream LoadResData("TWINS", "CUSTOM")       '133
        .AddImageFromStream LoadResData("VIBRATE", "CUSTOM")     '134
        .AddImageFromStream LoadResData("TEMPERATURE", "CUSTOM") '135
        .AddImageFromStream LoadResData("ZOOMIN", "CUSTOM")      '136
        .AddImageFromStream LoadResData("ZOOMOUT", "CUSTOM")     '137
        .AddImageFromStream LoadResData("ZOOMACTUAL", "CUSTOM")  '138
        .AddImageFromStream LoadResData("NEXTIMAGE", "CUSTOM")   '139
        .AddImageFromStream LoadResData("PREVIMAGE", "CUSTOM")   '140
        .AddImageFromStream LoadResData("DOWNLOADSRC", "CUSTOM") '141
        .AddImageFromStream LoadResData("LICENSE", "CUSTOM")     '142
        .AddImageFromStream LoadResData("SEPIA", "CUSTOM")       '143
        .AddImageFromStream LoadResData("CROPSEL", "CUSTOM")     '144
        .AddImageFromStream LoadResData("HSL", "CUSTOM")         '145
        .AddImageFromStream LoadResData("ROTATEANY", "CUSTOM")   '146
        .AddImageFromStream LoadResData("HEATMAP", "CUSTOM")     '147
        .AddImageFromStream LoadResData("STYLIZE", "CUSTOM")     '148
        .AddImageFromStream LoadResData("MODE24", "CUSTOM")      '149
        .AddImageFromStream LoadResData("MODE32", "CUSTOM")      '150
        .AddImageFromStream LoadResData("MODE24CHK", "CUSTOM")   '151
        .AddImageFromStream LoadResData("MODE32CHK", "CUSTOM")   '152
        
        'File Menu
        .PutImageToVBMenu 0, 0, 0       'Open Image
        .PutImageToVBMenu 1, 1, 0       'Open recent
        .PutImageToVBMenu 2, 2, 0       'Import
        .PutImageToVBMenu 3, 4, 0       'Save
        .PutImageToVBMenu 4, 5, 0       'Save As...
        .PutImageToVBMenu 5, 7, 0       'Close...
        .PutImageToVBMenu 6, 9, 0       'Batch conversion
        .PutImageToVBMenu 7, 11, 0      'Print
        .PutImageToVBMenu 43, 13, 0     'Exit
        
        '--> Import Sub-Menu
        'NOTE: the specific menu values will be different if the scanner plugin (eztw32.dll) isn't found.
        If ScanEnabled = True Then
            .PutImageToVBMenu 8, 0, 0, 2       'Scan Image
            .PutImageToVBMenu 45, 1, 0, 2      'Select Scanner
            .PutImageToVBMenu 9, 3, 0, 2       'Download Image
            .PutImageToVBMenu 10, 5, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 7, 0, 2      'Import from FRX
        Else
            .PutImageToVBMenu 9, 0, 0, 2       'Download Image
            .PutImageToVBMenu 10, 2, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 4, 0, 2      'Import from FRX
        End If
        
        'Edit Menu
        .PutImageToVBMenu 12, 0, 1      'Undo
        .PutImageToVBMenu 13, 1, 1      'Redo
        .PutImageToVBMenu 14, 2, 1      'Repeat Last Action
        .PutImageToVBMenu 15, 4, 1      'Copy
        .PutImageToVBMenu 16, 5, 1      'Paste
        .PutImageToVBMenu 17, 6, 1      'Empty Clipboard
        .PutImageToVBMenu 18, 8, 1      'Program Preferences
        
        'View Menu
        .PutImageToVBMenu 29, 0, 2     'Fit on Screen
        .PutImageToVBMenu 28, 1, 2     'Fit Viewport to Image
        .PutImageToVBMenu 136, 3, 2     'Zoom In
        .PutImageToVBMenu 137, 4, 2     'Zoom Out
        .PutImageToVBMenu 138, 10, 2     'Zoom 100%
        
        'Image Menu
        .PutImageToVBMenu 42, 0, 3      'Duplicate
        .PutImageToVBMenu 149, 2, 3     'Image Mode
            '--> Image Mode sub-menu
            .PutImageToVBMenu 149, 0, 3, 2   '24bpp
            .PutImageToVBMenu 150, 1, 3, 2   '32bpp
        .PutImageToVBMenu 19, 4, 3      'Resize
        .PutImageToVBMenu 144, 5, 3     'Crop to Selection
        .PutImageToVBMenu 24, 7, 3      'Mirror
        .PutImageToVBMenu 23, 8, 3      'Flip
        .PutImageToVBMenu 20, 10, 3      'Rotate Clockwise
        .PutImageToVBMenu 21, 11, 3      'Rotate Counter-clockwise
        .PutImageToVBMenu 22, 12, 3      'Rotate 180
        'NOTE: the specific menu values will be different if the FreeImage plugin (FreeImage.dll) isn't found.
        If imageFormats.FreeImageEnabled Then
            .PutImageToVBMenu 146, 13, 3     'Rotate Arbitrary
            .PutImageToVBMenu 125, 15, 3     'Isometric
            .PutImageToVBMenu 132, 16, 3     'Tile
        Else
            .PutImageToVBMenu 125, 14, 3     'Isometric
            .PutImageToVBMenu 132, 15, 3     'Tile
        End If
        
        'Color Menu
        .PutImageToVBMenu 46, 0, 4      'Brightness/Contrast
        .PutImageToVBMenu 47, 1, 4      'Gamma Correction
        .PutImageToVBMenu 145, 2, 4     'HSL adjustment
        .PutImageToVBMenu 48, 3, 4      'Levels
        .PutImageToVBMenu 135, 4, 4     'Temperature
        .PutImageToVBMenu 49, 5, 4      'White Balance
        .PutImageToVBMenu 50, 7, 4      'Histogram
            '--> Histogram sub-menu
            .PutImageToVBMenu 50, 0, 4, 7   'Display Histogram
            .PutImageToVBMenu 51, 2, 4, 7   'Equalize
            .PutImageToVBMenu 52, 3, 4, 7   'Stretch
        .PutImageToVBMenu 53, 9, 4      'Color Shift
            '--> Color-Shift sub-menu
            .PutImageToVBMenu 53, 0, 4, 9   'Shift Right
            .PutImageToVBMenu 54, 1, 4, 9   'Shift Left
        .PutImageToVBMenu 57, 10, 4      'Rechannel
            '--> Rechannel sub-menu
            '.PutImageToVBMenu 55, 0, 4, 9   'Red
            '.PutImageToVBMenu 56, 1, 4, 9   'Green
            '.PutImageToVBMenu 57, 2, 4, 9   'Blue
        .PutImageToVBMenu 58, 12, 4      'Black and White
        .PutImageToVBMenu 59, 13, 4      'Colorize
        .PutImageToVBMenu 60, 14, 4      'Enhance
            '--> Enhance sub-menu
            .PutImageToVBMenu 61, 0, 4, 14   'Contrast
            .PutImageToVBMenu 62, 1, 4, 14   'Highlights
            .PutImageToVBMenu 63, 2, 4, 14   'Midtones
            .PutImageToVBMenu 64, 3, 4, 14   'Shadows
        .PutImageToVBMenu 102, 15, 4     'Fade
            '--> Fade sub-menu
            .PutImageToVBMenu 103, 0, 4, 15  'Low Fade
            .PutImageToVBMenu 102, 1, 4, 15  'Medium Fade
            .PutImageToVBMenu 104, 2, 4, 15  'High Fade
            .PutImageToVBMenu 105, 3, 4, 15  'Custom Fade
            .PutImageToVBMenu 106, 5, 4, 15  'Unfade
        .PutImageToVBMenu 65, 16, 4      'Grayscale
        .PutImageToVBMenu 66, 17, 4      'Invert
            '--> Invert sub-menu
            .PutImageToVBMenu 67, 0, 4, 17   'Invert CMYK
            .PutImageToVBMenu 68, 1, 4, 17   'Invert Hue
            .PutImageToVBMenu 69, 2, 4, 17   'Invert RGB
            .PutImageToVBMenu 101, 4, 4, 17  'Compound Invert
        .PutImageToVBMenu 70, 18, 4      'Posterize
        .PutImageToVBMenu 143, 19, 4     'Sepia
        .PutImageToVBMenu 72, 21, 4      'Count Colors
        .PutImageToVBMenu 71, 22, 4      'Reduce Colors
        
        'Filters Menu
        .PutImageToVBMenu 73, 0, 5       'Fade Last
        .PutImageToVBMenu 74, 2, 5       'Artistic
            '--> Artistic sub-menu
            .PutImageToVBMenu 83, 0, 5, 2   'Antique (Sepia)
            .PutImageToVBMenu 84, 1, 5, 2   'Comic Book
            .PutImageToVBMenu 85, 2, 5, 2   'Mosaic
            .PutImageToVBMenu 86, 3, 5, 2   'Pencil
            .PutImageToVBMenu 87, 4, 5, 2   'Relief
        .PutImageToVBMenu 75, 3, 5       'Blur
            '--> Blur sub-menu
            .PutImageToVBMenu 88, 0, 5, 3   'Antialias
            .PutImageToVBMenu 89, 2, 5, 3   'Soften
            .PutImageToVBMenu 90, 3, 5, 3   'Soften More
            .PutImageToVBMenu 91, 4, 5, 3   'Blur
            .PutImageToVBMenu 92, 5, 5, 3   'Blur More
            .PutImageToVBMenu 93, 6, 5, 3   'Gaussian Blur
            .PutImageToVBMenu 94, 7, 5, 3   'Gaussian Blur More
            .PutImageToVBMenu 95, 9, 5, 3   'Grid Blur
        .PutImageToVBMenu 77, 4, 5       'Edges
            '--> Edges sub-menu
            .PutImageToVBMenu 100, 0, 5, 4  'Emboss / Engrave
            .PutImageToVBMenu 99, 1, 5, 4   'Enhance Edges
            .PutImageToVBMenu 77, 2, 5, 4   'Find Edges
        .PutImageToVBMenu 80, 5, 5       'Fun
            '--> Fun sub-menu
            .PutImageToVBMenu 126, 0, 5, 5  'Alien
            .PutImageToVBMenu 127, 1, 5, 5  'Blacklight
            .PutImageToVBMenu 128, 2, 5, 5  'Dream
            .PutImageToVBMenu 129, 3, 5, 5  'Radioactive
            .PutImageToVBMenu 131, 4, 5, 5  'Synthesize
            .PutImageToVBMenu 147, 5, 5, 5  'Thermograph
            .PutImageToVBMenu 134, 6, 5, 5  'Vibrate
        .PutImageToVBMenu 78, 6, 5       'Natural
            '--> Natural sub-menu
            .PutImageToVBMenu 107, 0, 5, 6  'Atmosphere
            .PutImageToVBMenu 108, 1, 5, 6  'Burn
            .PutImageToVBMenu 109, 2, 5, 6  'Fog
            .PutImageToVBMenu 110, 3, 5, 6  'Freeze
            .PutImageToVBMenu 111, 4, 5, 6  'Lava
            .PutImageToVBMenu 112, 5, 5, 6  'Ocean
            .PutImageToVBMenu 113, 6, 5, 6  'Rainbow
            .PutImageToVBMenu 114, 7, 5, 6  'Steel
            .PutImageToVBMenu 115, 8, 5, 6  'Water
        .PutImageToVBMenu 79, 7, 5       'Noise
            '--> Noise sub-menu
            .PutImageToVBMenu 116, 0, 5, 7  'Add Noise
            .PutImageToVBMenu 117, 2, 5, 7  'Despeckle
            .PutImageToVBMenu 118, 3, 5, 7  'Remove Orphan
        .PutImageToVBMenu 81, 8, 5       'Rank
            '--> Rank sub-menu
            .PutImageToVBMenu 119, 0, 5, 8  'Dilate
            .PutImageToVBMenu 120, 1, 5, 8  'Erode
            .PutImageToVBMenu 121, 2, 5, 8  'Extreme
            .PutImageToVBMenu 122, 4, 5, 8  'Custom Rank
        .PutImageToVBMenu 82, 9, 5       'Sharpen
            '--> Sharpen sub-menu
            .PutImageToVBMenu 124, 0, 5, 9  'Unsharp
            .PutImageToVBMenu 82, 2, 5, 9   'Sharpen
            .PutImageToVBMenu 123, 3, 5, 9  'Sharpen More
        .PutImageToVBMenu 148, 10, 5      'Stylize
            '--> Stylize sub-menu
            .PutImageToVBMenu 76, 0, 5, 10  'Diffuse
            .PutImageToVBMenu 96, 1, 5, 10  'Diffuse More
            .PutImageToVBMenu 97, 2, 5, 10  'Diffuse (Custom)
            .PutImageToVBMenu 130, 4, 5, 10 'Solarize
            .PutImageToVBMenu 133, 5, 5, 10 'Twins
        .PutImageToVBMenu 98, 12, 5      'Custom Filter
        
        'Macro Menu
        .PutImageToVBMenu 36, 0, 6     'Open Macro
        .PutImageToVBMenu 37, 2, 6     'Start Recording
        .PutImageToVBMenu 38, 3, 6     'Stop Recording
        
        'Window Menu
        .PutImageToVBMenu 139, 0, 7    'Next image
        .PutImageToVBMenu 140, 1, 7    'Previous image
        .PutImageToVBMenu 33, 3, 7     'Arrange Icons
        .PutImageToVBMenu 32, 4, 7     'Cascade
        .PutImageToVBMenu 31, 5, 7     'Tile Horizontally
        .PutImageToVBMenu 30, 6, 7     'Tile Vertically
        .PutImageToVBMenu 34, 8, 7     'Minimize All
        .PutImageToVBMenu 35, 9, 7     'Restore All
        
        'Help Menu
        .PutImageToVBMenu 40, 0, 8     'Donate
        .PutImageToVBMenu 41, 2, 8     'Check for updates
        .PutImageToVBMenu 26, 3, 8     'Submit Feedback
        .PutImageToVBMenu 39, 4, 8     'Submit Bug
        .PutImageToVBMenu 25, 6, 8     'Visit the PhotoDemon website
        .PutImageToVBMenu 141, 7, 8    'Download source code
        .PutImageToVBMenu 142, 8, 8    'License
        .PutImageToVBMenu 27, 10, 8    'About PD
    
    End With
    
    'Finally, calculate where to place the "Clear MRU" menu item
    Dim numOfMRUFiles As Long
    numOfMRUFiles = MRU_ReturnCount()
    cMenuImage.PutImageToVBMenu 44, numOfMRUFiles + 1, 0, 1
    
    'And initialize the MRU icon handler.  (Unfortunately, MRU icons must be disabled on XP.  We can't
    ' double-subclass the main form, and using a single menu icon class isn't possible at present.)
    If isVistaOrLater Then
        Set cMRUIcons = New clsMenuImage
        cMRUIcons.Init FormMain.hWnd, 64, 64
    End If
    
End Sub

'When menu captions are changed, the associated images are lost.  This forces a reset.
' At present, it only address the Undo and Redo menu items.
Public Sub ResetMenuIcons()

    'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
    If (Not isVistaOrLater) And (IsProgramCompiled = False) Then Exit Sub

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
        .PutImageToVBMenu 12, 0, 1 + posModifier    'Undo
        .PutImageToVBMenu 13, 1, 1 + posModifier    'Redo
    End With
    
    'Dynamically calculate the position of the Clear Recent Files menu item and update its icon
    Dim numOfMRUFiles As Long
    numOfMRUFiles = MRU_ReturnCount()
    cMenuImage.PutImageToVBMenu 44, numOfMRUFiles + 1, 0 + posModifier, 1
        
    'If the OS is Vista or later, render MRU icons to the Open Recent menu
    If isVistaOrLater Then
    
        cMRUIcons.Clear
        Dim tmpFilename As String
    
        'Loop through the MRU list, and attempt to load thumbnail images for each entry
        Dim i As Long
        For i = 0 To numOfMRUFiles
        
            'Start by seeing if an image exists for this MRU entry
            tmpFilename = getMRUThumbnailPath(i)
        
            'If the file exists, add it to the MRU icon handler
            If FileExist(tmpFilename) Then
            
                cMRUIcons.AddImageFromFile tmpFilename
                cMRUIcons.PutImageToVBMenu i, i, 0 + posModifier, 1
            
            End If
        
        Next i
        
    End If
        
End Sub

'As a courtesy to the user, the Image -> Mode icon is dynamically changed to match the current image's mode
Public Sub updateModeIcon(ByVal nMode As Boolean)
    If nMode Then
        'Update the parent menu
        cMenuImage.PutImageToVBMenu 150, 2, 3
        
        'Update the children menus as well
        cMenuImage.PutImageToVBMenu 149, 0, 3, 2   '24bpp
        cMenuImage.PutImageToVBMenu 152, 1, 3, 2   '32bpp
        
    Else
        cMenuImage.PutImageToVBMenu 149, 2, 3
        cMenuImage.PutImageToVBMenu 151, 0, 3, 2
        cMenuImage.PutImageToVBMenu 150, 1, 3, 2
    End If
End Sub

'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
'Again, thanks to Paul Turcksin for the original draft of this code.
Public Sub CreateCustomFormIcon(ByRef imgForm As FormImage)

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
   
    'The icon can be drawn at any size, but 16x16 is how it will (typically) end up on the form. I use 32x32 here
    ' in order to get slightly higher quality stretching during the resampling phase.
    Dim icoSize As Long
    icoSize = 32

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

    'Populate the icon header
    With icoInfo
      .fIcon = True
      .xHotspot = icoSize
      .yHotspot = icoSize
      .hbmMask = MonoBmp
      .hbmColor = InvertBmp
    End With
      
    'Render the icon to a handle
    Dim generatedIcon As Long
    generatedIcon = CreateIconIndirect(icoInfo)
    
    'Clear out our temporary masks (whose info are now embedded in the icon itself)
    DeleteObject icoInfo.hbmMask
    DeleteObject icoInfo.hbmColor
    DeleteDC MonoDC
    DeleteDC InvertDC
   
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

    Dim i As Long
    For i = 0 To numOfIcons - 1
        DestroyIcon iconHandles(i)
    Next i
    
    Erase iconHandles

End Sub
