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
'Last updated: 02/June/16
'Last update: finally purge all the old hard-coded theme color functions
'
'As of release 7.0, PhotoDemon supports the notion of "visual themes".  These themes are XML files that modify
' the program's appearance.
'
'To ensure that all UI elements are themed correctly, colors are never hard-coded.  Instead, they are retrieved
' from this class via one of several means (typically, control-specific color caches or universal color caches).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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
    UI_Accent = 0
    UI_AccentDark = 1
    UI_AccentLight = 2
    UI_AccentSemiDark = 3
    UI_AccentSemiLight = 4
    UI_AccentUltraDark = 5
    UI_AccentUltraLight = 6
    UI_Background = 7
    UI_CanvasElement = 8
    UI_ErrorRed = 9
    UI_GrayDefault = 10
    UI_GrayDisabled = 11
    UI_GrayDark = 12
    UI_GrayLight = 13
    UI_GrayNeutral = 14
    UI_GrayUltraLight = 15
    UI_LineEdge = 16
    UI_LineCenter = 17
    UI_TextReadOnly = 18
    UI_TextClickableSelected = 19
    UI_TextClickableUnselected = 20
    UI_ChannelRed = 21
    UI_ChannelGreen = 22
    UI_ChannelBlue = 23
    [_Last] = 23
    [_Count] = 24
End Enum

#If False Then
    Private Const UI_Accent = 0, UI_AccentDark = 1, UI_AccentLight = 2, UI_AccentSemiDark = 3, UI_AccentSemiLight = 4, UI_AccentUltraDark = 5, UI_AccentUltraLight = 6, UI_Background = 7, UI_CanvasElement = 8, UI_ErrorRed = 9
    Private Const UI_GrayDefault = 10, UI_GrayDisabled = 11, UI_GrayDark = 12, UI_GrayLight = 13, UI_GrayNeutral = 14, UI_GrayUltraLight = 15, UI_LineEdge = 16, UI_LineCenter = 17, UI_TextReadOnly = 18, UI_TextClickableSelected = 19
    Private Const UI_TextClickableUnselected = 20, UI_ChannelRed = 21, UI_ChannelGreen = 22, UI_ChannelBlue = 23
#End If

'Color retrieval and storage of program-wide UI colors is handled by a dedicated class
Private m_UniversalColors As pdThemeColors

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
        .LoadThemeColor UI_Accent, "UniversalAccent", IDE_BLUE
        .LoadThemeColor UI_AccentDark, "UniversalAccentDark", IDE_BLUE
        .LoadThemeColor UI_AccentSemiDark, "UniversalAccentSemidark", IDE_BLUE
        .LoadThemeColor UI_AccentUltraDark, "UniversalAccentUltradark", IDE_BLUE
        .LoadThemeColor UI_AccentLight, "UniversalAccentLight", IDE_BLUE
        .LoadThemeColor UI_AccentSemiLight, "UniversalAccentSemilight", IDE_BLUE
        .LoadThemeColor UI_AccentUltraLight, "UniversalAccentUltralight", IDE_BLUE
        .LoadThemeColor UI_Background, "UniversalBackground", IDE_WHITE
        .LoadThemeColor UI_CanvasElement, "UniversalCanvasElement", IDE_GRAY
        .LoadThemeColor UI_ErrorRed, "UniversalErrorRed", RGB(255, 0, 0)
        .LoadThemeColor UI_GrayDefault, "UniversalGrayDefault", IDE_GRAY
        .LoadThemeColor UI_GrayDisabled, "UniversalGrayDisabled", IDE_GRAY
        .LoadThemeColor UI_GrayDark, "UniversalGrayDark", IDE_GRAY
        .LoadThemeColor UI_GrayLight, "UniversalGrayLight", IDE_GRAY
        .LoadThemeColor UI_GrayNeutral, "UniversalGrayNeutral", IDE_GRAY
        .LoadThemeColor UI_GrayUltraLight, "UniversalGrayUltralight", IDE_GRAY
        .LoadThemeColor UI_LineEdge, "UniversalLineEdge", IDE_BLUE
        .LoadThemeColor UI_LineCenter, "UniversalLineCenter", IDE_GRAY
        .LoadThemeColor UI_TextReadOnly, "UniversalTextReadOnly", IDE_BLACK
        .LoadThemeColor UI_TextClickableSelected, "UniversalTextClickableSelected", IDE_WHITE
        .LoadThemeColor UI_TextClickableUnselected, "UniversalTextClickableUnselected", IDE_GRAY
        .LoadThemeColor UI_ChannelRed, "UniversalChannelRed", RGB(255, 0, 0)
        .LoadThemeColor UI_ChannelGreen, "UniversalChannelGreen", RGB(0, 255, 0)
        .LoadThemeColor UI_ChannelBlue, "UniversalChannelBlue", RGB(0, 0, 255)
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
End Sub
