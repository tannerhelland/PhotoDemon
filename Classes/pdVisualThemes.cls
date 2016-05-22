VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdVisualThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Visual Theming class
'Copyright 2013-2016 by Tanner Helland
'Created: 23/October/13
'Last updated: 22/January/16
'Last update: implement all the little XML bits that make custom theming files possible
'
'There isn't much here yet, but stay tuned - more is coming!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'***************************************************************************
'FYI FEB 2015:
' A bunch of the declarations in this class will be disappearing in the coming weeks.  PD's new XML theme format
' is coming along nicely, and most of PD's custom controls have already been migrated to the new system.
' As such, please ignore the dust until this class reaches its final form.
'***************************************************************************

'I have no idea what colors will actually be used in the final theme format; these are just placeholders for now,
' to enable centralized coloration of a few UCs.  The end goal is to provide the smallest possible subset of colors
' necessary to render PD's various UI elements.
Private Const NUM_THEME_COLORS As Long = 23

Public Enum PD_THEME_COLOR
    
    'Accent color variations
    PDTC_ACCENT_DEFAULT
    PDTC_ACCENT_HIGHLIGHT
    PDTC_ACCENT_ULTRALIGHT
    PDTC_ACCENT_SHADOW
    
    'Non-interactive variations (used to render control outlines, etc)
    PDTC_GRAY_DEFAULT
    PDTC_GRAY_HIGHLIGHT
    PDTC_GRAY_ULTRALIGHT
    PDTC_GRAY_SHADOW
    
    'Background variations
    PDTC_BACKGROUND_DEFAULT
    PDTC_BACKGROUND_CANVAS
    PDTC_BACKGROUND_COMMANDBAR
    
    'Text variations
    PDTC_TEXT_DEFAULT
    PDTC_TEXT_EDITBOX
    PDTC_TEXT_HYPERLINK
    PDTC_TEXT_TITLE
    PDTC_TEXT_INVERT
    
    'Misc requirements
    PDTC_DISABLED
    PDTC_OK_GREEN
    PDTC_CANCEL_RED
    PDTC_FOCUS_RECT
    
    PDTC_CHANNEL_RED
    PDTC_CHANNEL_GREEN
    PDTC_CHANNEL_BLUE
    
End Enum

#If False Then
    Private Const PDTC_ACCENT_DEFAULT = 0, PDTC_ACCENT_HIGHLIGHT = 0, PDTC_ACCENT_ULTRALIGHT = 0, PDTC_ACCENT_SHADOW = 0
    Private Const PDTC_GRAY_DEFAULT = 0, PDTC_GRAY_HIGHLIGHT = 0, PDTC_GRAY_ULTRALIGHT = 0, PDTC_GRAY_SHADOW = 0
    Private Const PDTC_BACKGROUND_DEFAULT = 0, PDTC_BACKGROUND_CANVAS = 0, PDTC_BACKGROUND_COMMANDBAR = 0
    Private Const PDTC_TEXT_DEFAULT = 0, PDTC_TEXT_EDITBOX = 0, PDTC_TEXT_HYPERLINK = 0, PDTC_TEXT_TITLE = 0, PDTC_TEXT_INVERT = 0
    Private Const PDTC_DISABLED = 0, PDTC_OK_GREEN = 0, PDTC_CANCEL_RED = 0, PDTC_FOCUS_RECT = 0
    Private Const PDTC_CHANNEL_RED = 0, PDTC_CHANNEL_GREEN = 0, PDTC_CHANNEL_BLUE = 0
#End If

'As a broad advisement, themes are classified as:
' - THEME_LIGHT (dark accents and text on light backgrounds)
' - THEME_DARK (light accents and text on dark backgrounds)
' - THEME_HIGH_CONTRAST (ultra-contrasted elements, for accessibility)
' This is important when finding replacement colors for missing color entries, as we may have to fall back on PD's default
' values for some entries, and using LIGHT default entries on a DARK theme would be bad.
Public Enum PD_THEME_CONTRAST
    THEME_LIGHT = 0
    THEME_DARK = 1
    THEME_HIGH_CONTRAST = 2
End Enum

#If False Then
    Private Const THEME_LIGHT = 0, THEME_DARK = 1, THEME_HIGH_CONTRAST = 2
#End If

Private curThemeContrast As PD_THEME_CONTRAST

'When a theme is loaded, this class will cache the relevant color values inside this structure.  That makes retrieval
' extremely fast and easy.  Also, it makes it possible to create themes at run-time with minimal effort.
Private curThemeColors() As Long

'XML object for parsing theme files.
Private m_XML As pdXML

'When colors are retrieved from the theme file, we cache them locally.  This spares us time on subsequent color requests,
' especially for generic colors (e.g. "Background") which are retrieved by multiple controls.
Private m_NumColorsCached As Long
Private m_ColorCache() As PDCachedColor
Private Const DEFAULT_COLOR_CACHE_SIZE As Long = 16

'For the most part, we leave individual controls to manage their own color lists.  This provides excellent flexibility
' with UI rendering.  However, there are some colors that appear so frequently throughout PD that it makes more sense
' to cache them here, so one-off functions don't need to deal with messy color maintenance.
Public Enum PD_UI_COLOR_LIST
    [_First] = 0
    UI_CanvasElement = 0
    UI_LineEdge = 1
    UI_LineCenter = 2
    UI_ErrorRed = 3
    UI_TextReadOnly = 4
    UI_TextClickableSelected = 5
    UI_TextClickableUnselected = 6
    [_Last] = 6
    [_Count] = 7
End Enum

#If False Then
    Private Const UI_CanvasElement = 0, UI_LineEdge = 1, UI_LineCenter = 2, UI_ErrorRed = 3, UI_TextReadOnly = 4, UI_TextClickableSelected = 5, UI_TextClickableUnselected = 6
#End If

'Color retrieval and storage of program-wide UI colors is handled by a dedicated class
Private m_UniversalColors As pdThemeColors

'Despite an identical name, this function does *not* use the Win32 API to retrieve a color from the system theme.
' Instead, it retrieves a color from PD's current theme color cache.
Friend Function GetThemeColor(ByVal colorType As PD_THEME_COLOR) As Long
    
    'Returning a theme color is fairly trivial.  Start by seeing if the requested color is available.
    GetThemeColor = curThemeColors(colorType)
    
    'If the retrieved value is -1, that means that the current theme does not contain a value for this color.
    ' Return PD's default value for this color.
    If GetThemeColor = -1 Then
    
        'Default to the relevant PD default theme value.
        GetThemeColor = GetDefaultPDColor(colorType)
        
    End If
        
End Function

'Reset the theme color cache to all -1 values.
Private Sub ResetThemeCache()

    Dim i As Long
    For i = 0 To NUM_THEME_COLORS - 1
        curThemeColors(i) = -1
    Next i

End Sub

'Fill the curThemeColors array with PD's default color scheme
Public Sub LoadDefaultPDTheme()

    'Default to the light theme
    ' (TODO: read the user's preference for theme, which should probably be just a path to the theme file?)
    curThemeContrast = THEME_LIGHT
    
    'Retrieve the preferred theme file from the user preferences file.  (NOTE: this step will fail inside the IDE.)
    Dim themeFilename As String
    If g_IsProgramRunning Then
        themeFilename = "Default_Light.xml"
    Else
        'TODO: pull the theme directly from the resource file, or perhaps trigger some kind of failsafe IDE mode...?
    End If
    
    'Load the preferred XML file, and if it fails, fall back to PD's default theme
    Dim themeLoadedCorrectly As Boolean
    themeLoadedCorrectly = Me.LoadThemeFile(themeFilename)
    
    #If DEBUGMODE = 1 Then
        If themeLoadedCorrectly Then
            If g_IsProgramRunning Then pdDebug.LogAction "Successfully loaded theme file: " & themeFilename
        Else
            If g_IsProgramRunning Then pdDebug.LogAction "WARNING!  Failed to load theme file: " & themeFilename
        End If
    #End If
    
    'Theme colors are loaded on-demand, so we have no further work to do here
    
    'OLD CODE FOLLOWS:
    
    
    'Loop through the curThemeColors array, filling each entry with the PD default
    Dim i As Long
    For i = 0 To NUM_THEME_COLORS - 1
        curThemeColors(i) = GetDefaultPDColor(i)
    Next i

End Sub

'Load a given theme file.  Note that the filename SHOULD NOT INCLUDE THE FULL PATH - just the filename.  PD will
' automatically search the /App and /Data folders as necessary to find the file.
'
'Also, while this function does return success/failure status, if the load operation fails, PD will automatically
' fall back to its default theme to prevent the program from exploding.
Public Function LoadThemeFile(ByVal themeFilename As String, Optional ByVal overrideColorDefinitionFilename As String = vbNullString) As Boolean
    
    'Like most things in PD, themes can exist in two places:
    ' 1) The "untouchable" /App folder, which contains PD's core data
    ' 2) The /Data folder, which contains user-specific data (and can be deleted willy-nilly)
    
    'Attempt to resolve the passed themeFilename to one of these locations, giving preference to the /App folder.
    ' (TODO 6.8: make a decision on how much control we expose over theme editing; maybe the /Data folder is unnecessary)
    If g_IsProgramRunning And Not g_ProgramShuttingDown Then
    
        Dim fullThemePath As String
        fullThemePath = g_UserPreferences.GetThemePath & themeFilename
        
        If m_XML.LoadXMLFile(fullThemePath) Then
            
            'Perform minor validation on the file
            LoadThemeFile = m_XML.isPDDataType("Visual theme") And m_XML.validateLoadedXMLData("Colors")
            
        Else
            #If DEBUGMODE = 1 Then
                If g_IsProgramRunning Then
                    pdDebug.LogAction "WARNING! Failed to load requested theme: " & fullThemePath
                    pdDebug.LogAction "WARNING! Falling back to default PD theme..."
                End If
            #End If
            
            fullThemePath = g_UserPreferences.GetThemePath & "Default_Light.xml"
            LoadThemeFile = m_XML.LoadXMLFile(fullThemePath)
            LoadThemeFile = LoadThemeFile And m_XML.isPDDataType("Visual theme") And m_XML.validateLoadedXMLData("Colors")
            
        End If
        
        'Regardless of success or failure, reset our internal color cache(s)
        ResetColorCache
        
        'Theme files generally consist of two parts: a theme XML file, and a color definition file.  This system allows
        ' a single theme file to be re-used against multiple color definition files, making it easy to support various
        ' color schemes with minimal work.
        
        'Anyway, if the theme file loaded correctly, we need to also load its color definition file (if any).
        If LoadThemeFile Then
        
            'Color description files are listed under the DefinitionFile tag.  (This tag is optional, so we can
            ' assume all definitions are embedded in the file if the DefinitionFile tag doesn't exist.)
            If m_XML.DoesTagExist("DefinitionFile") Or (Len(overrideColorDefinitionFilename) <> 0) Then
            
                'Load and validate the specified definition file
                Dim tmpXML As pdXML
                Set tmpXML = New pdXML
                
                Dim fullDefinitionPath As String
                If Len(overrideColorDefinitionFilename) <> 0 Then
                    fullDefinitionPath = g_UserPreferences.GetThemePath & overrideColorDefinitionFilename
                Else
                    fullDefinitionPath = g_UserPreferences.GetThemePath & m_XML.getUniqueTag_String("DefinitionFile")
                End If
                
                If tmpXML.LoadXMLFile(fullDefinitionPath) Then
                
                    If tmpXML.isPDDataType("Color definitions") And tmpXML.validateLoadedXMLData("Definitions") Then
                    
                        'Retrieve the definition list
                        Dim colorDefinitionList As String
                        colorDefinitionList = tmpXML.getUniqueTag_String("Definitions")
                        
                        'Plug it straight into the Definitions section of the current XML file.
                        colorDefinitionList = colorDefinitionList & m_XML.getUniqueTag_String("Definitions")
                        
                        If Not m_XML.updateTag("Definitions", colorDefinitionList) Then
                            #If DEBUGMODE = 1 Then
                                If g_IsProgramRunning Then pdDebug.LogAction "WARNING!  The color definition file listed in " & fullThemePath & " couldn't be dynamically inserted into its parent theme."
                            #End If
                        End If
                    
                    Else
                        #If DEBUGMODE = 1 Then
                            If g_IsProgramRunning Then pdDebug.LogAction "WARNING!  The color definition file listed in " & fullThemePath & " failed to validate."
                        #End If
                    End If
                
                Else
                    #If DEBUGMODE = 1 Then
                        If g_IsProgramRunning Then pdDebug.LogAction "WARNING!  The color definition file listed in " & fullThemePath & " failed to load."
                    #End If
                End If
            
            End If
        
        End If
        
        'With all color definitions imported, we can now cache a few program-wide UI colors
        CacheUniversalColors
        
        'If the user's choice of theme didn't load correctly, or the default theme failed to load, run some heuristics
        ' on the theme folder.
        If Not LoadThemeFile Then
            #If DEBUGMODE = 1 Then
                If g_IsProgramRunning Then pdDebug.LogAction "WARNING!  PD's default theme failed to load!  Catastrophic failure imminent!"
            #End If
            ' (TODO: this entire step, including pulling themes from the .exe's resource section as necessary)
        End If
        
    End If
    
End Function

'Retrieve a default PD theme color.  These values are hard-coded, and will be superceded by any values supplied
' by a user-selected theme.
Private Function GetDefaultPDColor(ByVal colorType As PD_THEME_COLOR) As Long

    'LIGHT THEME DEFAULTS
    If curThemeContrast = THEME_LIGHT Then
    
        'FOR REFERENCE: here are some of the colors currently in use in PD, prior to fully implementing theming:

        ' Most label controls: #404040 | RGB(64, 64, 64)
        ' Slider track border: #B1BAC2 | RGB(177, 186, 194)
        ' Slider fill and knob: #3CAFE6 | RGB(60, 175, 230)
        ' Radio button border: #7e8c92 | RGB(126, 140, 146)
        ' Radio button fill (if selected): #3296dc | RGB(50, 150, 220)
        ' Green for "valid input": #00c0000 | RGB(0, 192, 0)
        ' Red for "bad input": #c000000 | RGB(192, 0, 0)
        ' Hyperlink blue: #3170c0 | RGB(49, 112, 192)
        
        'FOR REFERENCE: this is a list of theme colors being used *in the present build*:
        
        ' Radio button UC:
            ' PDTC_BACKGROUND_DEFAULT (background)
            
            ' PDTC_ACCENT_NONINTERACTIVE, PDTCV_NORMAL (radio border, enabled)
            ' PDTC_ACCENT_INTERACTIVE, PDTCV_NORMAL (radio fill, enabled)
            ' PDTC_ACCENT_INTERACTIVE, PDTCV_DISABLED (radio fill + border, disabled)
            
            ' PDTC_TEXT_DEFAULT, PDTCV_HIGHLIGHT (text, mouse_over)
            ' PDTC_TEXT_DEFAULT, PDTCV_NORMAL (text, not mouse_over)
            ' PDTC_TEXT_DEFAULT, PDTCV_DISABLED (text, disabled)
        
        Select Case colorType
        
            'Accent color variations
            
            'Accent lines, etc
            Case PDTC_ACCENT_DEFAULT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("3296dc")
            
            'Slider track, large contiguous blocks of color
            Case PDTC_ACCENT_HIGHLIGHT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#3cafe6")
            
            'Background on toolbox images (which can't afford to be as dark as the usual highlight value)
            Case PDTC_ACCENT_ULTRALIGHT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("d2f0fa")
            
            'Small or heavily aliased items (e.g. check box check, radio button dot)
            Case PDTC_ACCENT_SHADOW
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#3270dc")
            
            
            'Non-interactive variations (used to render control outlines, etc)
            Case PDTC_GRAY_DEFAULT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#929292")
                
            Case PDTC_GRAY_HIGHLIGHT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#bababa")
            
            Case PDTC_GRAY_ULTRALIGHT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#e2e2e2")
            
            Case PDTC_GRAY_SHADOW
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#5b5b5b")
            
            'Background variations
            Case PDTC_BACKGROUND_DEFAULT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#ffffff")
                
            Case PDTC_BACKGROUND_CANVAS
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#a0a0a0")
                
            Case PDTC_BACKGROUND_COMMANDBAR
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#dcdce1")
            
            'Text variations
            Case PDTC_TEXT_DEFAULT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#404040")
                
            Case PDTC_TEXT_EDITBOX
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#202020")
                
            Case PDTC_TEXT_HYPERLINK
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#3170c0")
                
            Case PDTC_TEXT_TITLE
            
            Case PDTC_TEXT_INVERT
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#ffffff")
            
            'Misc colors
            Case PDTC_DISABLED
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#b1bac2")
            
            Case PDTC_OK_GREEN
            Case PDTC_CANCEL_RED
                GetDefaultPDColor = Colors.GetRGBLongFromHex("#e81828")
            
            Case PDTC_FOCUS_RECT
            
            Case PDTC_CHANNEL_RED
                    GetDefaultPDColor = Colors.GetRGBLongFromHex("#dc2824")
                    
            Case PDTC_CHANNEL_GREEN
                    GetDefaultPDColor = Colors.GetRGBLongFromHex("#3ccf31")
                    
            Case PDTC_CHANNEL_BLUE
                    GetDefaultPDColor = Colors.GetRGBLongFromHex("#317bcf")
            
        End Select
    
    'DARK THEME DEFAULTS
    ElseIf curThemeContrast = THEME_DARK Then
    
    'HIGH CONTRAST DEFAULTS
    Else
    
    End If

End Function

'Call this function to verify that an object exists inside the current theme file.  If it doesn't, you should not
' proceed with color loading.
Public Function VerifyThemeObject(ByRef objectName As String) As Boolean
    VerifyThemeObject = m_XML.DoesTagExist(objectName)
End Function

'Look up a unique theme color in the current theme.  Object name is required, and this class will automatically fall back
' to the Default namespace as necessary.  Also, colors described by definition will automatically be tracked back to their
' source.  (Note, however, that this function has no way to deal with circular references, so please avoid that.)
' RETURNS: a color hexadecimal value if successful; a null-string otherwise.
Public Function LookUpColor(ByVal objectName As String, ByRef colorName As String) As String

    'First things first: see if the object name exists in the theme file.  If it doesn't, we need to fall back to the
    ' "default" namespace.
    Const DEFAULT_NAMESPACE As String = "Default"
    Dim objectNameExists As Boolean
    objectNameExists = m_XML.DoesTagExist(objectName)
    If Not objectNameExists Then
        objectName = DEFAULT_NAMESPACE
        objectNameExists = m_XML.DoesTagExist(objectName)
    End If
    
    'If the color exists in either the Default or object-specific namespace, we can proceed with parsing.
    If objectNameExists Then
        
        'Inside the current object's color definition block, retrieve the specified color
        Dim colorDescription As String, finalColor As String
        colorDescription = m_XML.GetNonUniqueTag_String(colorName, objectName)
        
        'If we retrieved any valid string, attempt to resolve it to an actual color value.  (At this point, the color
        ' may just be a variable instead of an actual hex value.)
        If Len(colorDescription) <> 0 Then
            finalColor = ResolveColor(colorDescription)
        
        'If we used a custom object name, but no color is defined for that value, try a new retrieval from
        ' the "Default" namespace.  (Empty colors are still valid, as long as their Default variant is defined.)
        Else
            If StrComp(objectName, DEFAULT_NAMESPACE, vbBinaryCompare) <> 0 Then
                objectName = DEFAULT_NAMESPACE
                If m_XML.DoesTagExist(objectName) Then
                    colorDescription = m_XML.GetNonUniqueTag_String(colorName, objectName)
                    If Len(colorDescription) <> 0 Then finalColor = ResolveColor(colorDescription)
                End If
            End If
        End If
        
        LookUpColor = finalColor
        
    Else
        LookUpColor = vbNullString
    End If

End Function

'Given the raw value retrieved by LookUpColor(), above, retrieve that color's ultimate representation (e.g. not a
' named color variable, but an actual color literal, like #ff0000).
Private Function ResolveColor(ByVal initialColorValue As String) As String

    Do
        
        'First, reject any empty strings (to prevent subsequent parse errors)
        If Len(initialColorValue) = 0 Then
            ResolveColor = vbNullString
            Exit Function
        End If
        
        'Next, see if the current color value appears to be some kind of valid color representation
        If Colors.IsStringAColor(initialColorValue) Then
            ResolveColor = initialColorValue
            Exit Function
        
        'This is not a valid color representation, so assume it's a custom color descriptor (or invalid, I suppose)
        Else
            
            'Attempt to retrieve a new value from the theme's color definition section, then run our validation
            ' checks a second time.  (We'll repeat this until we fail to retrieve a new definition, or we identify
            ' a string that can be parsed into an actual color.)
            initialColorValue = m_XML.getUniqueTag_String(initialColorValue, vbNullString, , "Definitions")
            
        End If
    
    Loop

End Function

'Whenever a new theme is loaded, we must wipe the entire color cache.
Private Sub ResetColorCache()
    m_NumColorsCached = 0
    ReDim m_ColorCache(0 To DEFAULT_COLOR_CACHE_SIZE - 1) As PDCachedColor
End Sub

'After the external pdThemeColors class has properly resolved a base color (and all its variants) to final RGB longs,
' it will cache the newly created variable via this function.  This allows subsequent color requests to bypass the
' XML data entirely.
Friend Sub AddColorToCache(ByRef objectName As String, ByRef colorName As String, ByRef srcColorEntry As PDThemeColor)
    
    m_ColorCache(m_NumColorsCached).OrigObjectName = objectName
    m_ColorCache(m_NumColorsCached).OrigColorName = colorName
    m_ColorCache(m_NumColorsCached).OrigColorValues = srcColorEntry
    
    m_NumColorsCached = m_NumColorsCached + 1
    If m_NumColorsCached > UBound(m_ColorCache) Then ReDim Preserve m_ColorCache(0 To m_NumColorsCached * 2 - 1) As PDCachedColor
    
End Sub

'Look up a color in the color cache.  If it exists, the function returns TRUE, and the destination PDThemeColor struct
' is filled with the matching cache values.
Friend Function RetrieveColorFromCache(ByRef objectName As String, ByRef colorName As String, ByRef dstColorEntry As PDThemeColor) As Boolean

    RetrieveColorFromCache = False
    
    Dim i As Long
    For i = 0 To m_NumColorsCached - 1
        If StrComp(objectName, m_ColorCache(i).OrigObjectName, vbBinaryCompare) = 0 Then
            If StrComp(colorName, m_ColorCache(i).OrigColorName, vbBinaryCompare) = 0 Then
                RetrieveColorFromCache = True
                dstColorEntry = m_ColorCache(i).OrigColorValues
                Exit For
            End If
        End If
    Next i
    
End Function

'For the most part, PD lets individual control instances manage their own color lists.  This provides high levels of
' flexibility with rendering, as different controls may favor different techniques.  However, some colors are so
' ubiquitous throughout PD that it's easier to cache their results locally, then let outside functions retrieve colors
' with minimal effort on this part.
'
'Obviously, this cache must be reset any time a new theme file is loaded.  As there is no easy way for external functions
' to be notified of such a change, you should *not* reuse colors retrieved from this cache.  They need to be retrieved
' anew on every use.
Private Sub CacheUniversalColors()

    Dim colorCount As PD_UI_COLOR_LIST: colorCount = [_Count]
    m_UniversalColors.InitializeColorList "UIElements", colorCount
    
    With m_UniversalColors
        .LoadThemeColor UI_CanvasElement, "UniversalCanvasElement", IDE_GRAY
        .LoadThemeColor UI_LineEdge, "UniversalLineEdge", IDE_BLUE
        .LoadThemeColor UI_LineCenter, "UniversalLineCenter", IDE_GRAY
        .LoadThemeColor UI_ErrorRed, "UniversalErrorRed", RGB(255, 0, 0)
        .LoadThemeColor UI_TextReadOnly, "UniversalTextReadOnly", IDE_BLACK
        .LoadThemeColor UI_TextClickableSelected, "UniversalTextClickableSelected", IDE_WHITE
        .LoadThemeColor UI_TextClickableUnselected, "UniversalTextClickableUnselected", IDE_GRAY
    End With
    
End Sub

'External functions can use this to retrieve a color from the local m_UniversalColors cache.  If an object requires
' a bunch of object-specific colors, they will get better performance by managing their own color cache.
Friend Function GetGenericUIColor(ByVal colorID As PD_UI_COLOR_LIST, Optional ByVal enabledState As Boolean = True, Optional ByVal activeState As Boolean = False, Optional ByVal hoverState As Boolean = False) As Long
    GetGenericUIColor = m_UniversalColors.RetrieveColor(colorID, enabledState, activeState, hoverState)
End Function

Private Sub Class_Initialize()
    
    Set m_XML = New pdXML
    Set m_UniversalColors = New pdThemeColors
    
    'OLD CODE FOLLOWS
    
    'Prepare our color cache
    ReDim curThemeColors(0 To NUM_THEME_COLORS - 1) As Long
    
    'Reset the color cache.  This is important, because 0 is a valid color, so we can't use it to indicate a
    ' missing theme entry.  Instead, we use -1 to represent a missing entry.  The color retrieval function
    ' will automatically cascade to the nearest relevant value if an entry can't be found.
    ResetThemeCache
    
End Sub
