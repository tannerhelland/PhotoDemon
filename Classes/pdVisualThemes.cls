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
'Copyright 2013-2017 by Tanner Helland
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
' - LIGHT (dark accents and text on light backgrounds)
' - DARK (light accents and text on dark backgrounds)
' - HIGH CONTRAST (ultra-contrasted elements, for accessibility)
'Among other things, these values are used to correctly render things like monochrome icons.
Public Enum PD_THEME_CLASS
    PDTC_Light = 0
    PDTC_Dark = 1
    PDTC_HighContrast = 2
End Enum

#If False Then
    Private Const PDTC_Light = 0, PDTC_Dark = 1, PDTC_HighContrast = 2
#End If

Private m_ThemeClass As PD_THEME_CLASS

Public Enum PD_THEME_ACCENT
    PDTA_Undefined = -1
    PDTA_Blue = 0
    PDTA_Green = 1
    PDTA_Purple = 2
End Enum

#If False Then
    Private Const PDTA_Undefined = -1, PDTA_Blue = 0, PDTA_Green = 1, PDTA_Purple = 2
#End If

Private m_ThemeAccent As PD_THEME_ACCENT

'Extra API functions for painting form backgrounds
Private Const WM_PAINT As Long = &HF
Private Const WM_ERASEBKGND As Long = &H14
Private Declare Function ValidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal targetHwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal targetHwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHwnd As Long) As Long

'As part of the painting process, we're gonna be generating a looot of paint messages.  To avoid churn, we'll declare
' a single paint struct and update rect up front.
Private m_PaintStruct As PAINTSTRUCT
Private m_UpdateRect As RECT

'XML object for parsing theme files.
Private m_XML As pdXML
Private Const DEFAULT_NAMESPACE As String = "Default"

'When colors are retrieved from the theme file, we cache them locally.  This spares us time on subsequent color requests,
' especially for generic colors (e.g. "Background") which are retrieved by multiple controls.
Private m_NumColorsCached As Long
Private m_ColorCache() As PDCachedColor
Private Const DEFAULT_COLOR_CACHE_SIZE As Long = 16

'For the most part, we leave individual controls to manage their own color lists.  This provides excellent flexibility
' with UI rendering.  However, there are some colors that appear so frequently throughout PD that it makes more sense
' to cache them here, so one-off functions aren't burdened with themed color maintenance.
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
    UI_ChannelRed = 9
    UI_ChannelGreen = 10
    UI_ChannelBlue = 11
    UI_ErrorRed = 12
    UI_GrayDefault = 13
    UI_GrayDisabled = 14
    UI_GrayDark = 15
    UI_GrayLight = 16
    UI_GrayNeutral = 17
    UI_GrayUltraLight = 18
    UI_IconMonochrome = 19
    UI_IconMonochromeMenu = 20
    UI_ImageDisabled = 21
    UI_LineEdge = 22
    UI_LineCenter = 23
    UI_TextReadOnly = 24
    UI_TextClickable = 25
    UI_TextClickableSelected = 26
    UI_TextClickableUnselected = 27
    [_Last] = 27
    [_Count] = 28
End Enum

#If False Then
    Private Const UI_Accent = 0, UI_AccentDark = 1, UI_AccentLight = 2, UI_AccentSemiDark = 3, UI_AccentSemiLight = 4, UI_AccentUltraDark = 5, UI_AccentUltraLight = 6, UI_Background = 7, UI_CanvasElement = 8, UI_ChannelRed = 9
    Private Const UI_ChannelGreen = 10, UI_ChannelBlue = 11, UI_ErrorRed = 12, UI_GrayDefault = 13, UI_GrayDisabled = 14, UI_GrayDark = 15, UI_GrayLight = 16, UI_GrayNeutral = 17, UI_GrayUltraLight = 18, UI_IconMonochrome = 19
    Private Const UI_IconMonochromeMenu = 20, UI_ImageDisabled = 21, UI_LineEdge = 22, UI_LineCenter = 23, UI_TextReadOnly = 24, UI_TextClickable = 25, UI_TextClickableSelected = 26, UI_TextClickableUnselected = 27
#End If

'Color retrieval and storage of program-wide UI colors is handled by a dedicated class
Private m_UniversalColors As pdThemeColors

'As a convenience, we manually paint form backgrounds (as VB won't paint them before first display, if the background
' color changes in Form_Load - as it does for us)
Private m_Subclass As cSelfSubHookCallback

'Each theme is assigned a custom ID when it is loaded.  This ID can be used by external objects to see if they
' need to re-theme themselves.  (If the program's theme hasn't changed since the last time they re-themed,
' they can safely ignore UpdateAgainstCurrentTheme requests.)
'
'Note that UI elements should not generally query this value directly; instead, they should use the Interface
' module to do it, because it appends a language ID to this value.
Private m_CurrentThemeID As String

Friend Function GetCurrentThemeID() As String
    GetCurrentThemeID = m_CurrentThemeID
End Function

'Pull the current default PD theme from the user preferences file, and attempt to load it.  If the theme can't be loaded,
' we'll fall back to PD's default light-on-dark theme.
Friend Sub LoadDefaultPDTheme()

    'By default, assume a light-on-dark theme.  (This will be overridden by a successful theme load.)
    m_ThemeClass = PDTC_Light
    
    'Retrieve the preferred theme file from the user preferences file.  (NOTE: this step will fail inside the IDE.)
    Dim themeName As String, themeFilename As String, accentName As String, accentFilename As String
    If g_IsProgramRunning Then
        themeName = g_UserPreferences.GetPref_String("Themes", "CurrentTheme", "Light")
        themeFilename = "Default_" & themeName & ".xml"
        accentName = g_UserPreferences.GetPref_String("Themes", "CurrentAccent", "Blue")
        accentFilename = "Colors_" & accentName & ".xml"
    Else
        'FYI: inside the designer, PD will silently fall back on hard-coded IDE colors
    End If
    
    'Load the preferred XML file, and if it fails, fall back to PD's default theme
    Dim themeLoadedCorrectly As Boolean: themeLoadedCorrectly = False
    If (Len(themeFilename) <> 0) Then
        If (Len(accentFilename) <> 0) Then
            themeLoadedCorrectly = Me.LoadThemeFile(themeFilename, accentFilename)
        Else
            themeLoadedCorrectly = Me.LoadThemeFile(themeFilename)
        End If
    End If
    
    If themeLoadedCorrectly Then
        If g_IsProgramRunning Then RaiseThemingError "successfully loaded theme file: " & themeFilename, True
    Else
        If g_IsProgramRunning Then
            
            RaiseThemingError "failed to load theme file: " & themeFilename & ".  Attempting theme recovery now...", True
            AttemptThemeRecovery
            RaiseThemingError "Best attempt at theme recovery was made.  Trying to load default theme now...", True
            
            m_ThemeClass = PDTC_Light
            themeName = "Light"
            themeFilename = "Default_Light.xml"
            accentName = "Blue"
            accentFilename = "Colors_Blue.xml"
            themeLoadedCorrectly = Me.LoadThemeFile(themeFilename, accentFilename)
            
            'If this attempt fails, there's nothing left to try.  Pray to the programming gods that the
            ' program will even load.  (It probably won't.)
            If themeLoadedCorrectly Then
                RaiseThemingError "Dodged a bullet there - theme recovery successful.  Proceeding normally.", True
            Else
                RaiseThemingError "Theme recovery failed.  This session is probably doomed.", True
            End If
            
        End If
    End If
    
    'Notify the UI manager that a new theme is active.  Individual controls need to check this to know if
    ' they require a redraw to match any new theme settings.
    Interface.GenerateInterfaceID
    
    'Theme colors are loaded on-demand, so we have no further work to do here
    
End Sub

'Failed to load the user's requested theme at startup?  Call this function to attempt a recovery.
' (Some users use stupid .zip file software that doesn't preserve folders, so themes may be spread across
'  the root folder.  This sub will try to fix that.)
Private Sub AttemptThemeRecovery()

    'At present, there are six key theme files PD expects to exist (in the /App/PhotoDemon/Themes folder).
    ' These files come in three types:
    ' 1) Core theme files that describe most of the colors used by the program
    ' 2) Theme "accent" files that describe a (much smaller) set of accent colors.  These can be plugged
    '    into any base theme for a unique coloring scheme.
    ' 3) A theme resource file, which contains all of the icons associated with a given theme.  (At present,
    '    there is only one of these, built by me.)
    
    'If these files go missing, PD is liable to break.  Horribly.
    ' In the future, if we can't recover the files in question, I intend to pull a failsafe copy from PD's
    ' resource file, which should be enough for the user to at least use the program as-is.  For now,
    ' however, there is no failsafe if the theme folder breaks, and PD is unlikely to be usable.
    
    'NOTE!  YOU MUST UPDATE THIS LIST if new theme files become required for basic program functionality.
    DoesThemeFileExist "Colors_Blue.xml"
    DoesThemeFileExist "Colors_Green.xml"
    DoesThemeFileExist "Colors_Purple.xml"
    DoesThemeFileExist "Default_Dark.xml"
    DoesThemeFileExist "Default_Light.xml"
    
End Sub

'This function performs several tasks:
' 1) If the requested theme file exists in the target folder, great; it returns TRUE and exits.
' 2) If the requested theme file does NOT exist in the target folder, it scans the program folder to see if it
'     can find it there.
' 3) If it finds a missing theme in the program folder, it will automatically move the file to the theme folder,
'     where it belongs.
' 4) If the move is successful, it will return TRUE and exit.
Private Function DoesThemeFileExist(ByVal themeFile As String) As Boolean
    
    Dim themeFilename As String
    themeFilename = g_UserPreferences.GetThemePath(False) & themeFile
    
    'pdFSO is used for all file interactions
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'See if the file exists.  If it does, great!  We can exit immediately.
    If cFile.FileExist(themeFilename) Then
        DoesThemeFileExist = True
    
    'The theme file is missing.  Let's see if we can find it.
    Else
    
        pdDebug.LogAction "WARNING!  Theme file <" & themeFile & "> is missing.  Scanning alternate folders..."
        
        'See if the file exists in the base PD folder.  This can happen if a user unknowingly extracts
        ' the PD .zip without folders preserved.
        If cFile.FileExist(g_UserPreferences.GetProgramPath & themeFile) Then
            
            pdDebug.LogAction "UPDATE!  Theme file found in the base PD folder.  Attempting to relocate..."
            
            'Move the file to the proper folder
            If cFile.CopyFile(g_UserPreferences.GetProgramPath & themeFile, g_UserPreferences.GetThemePath & themeFile) Then
                
                pdDebug.LogAction "UPDATE!  Theme file relocated successfully."
                
                'Kill the old file instance
                cFile.KillFile g_UserPreferences.GetProgramPath & themeFile
                
                'Return success!
                DoesThemeFileExist = True
            
            'The file couldn't be moved.  There's probably write issues with the folder structure, in which case
            ' this program session is pretty much doomed.  Exit now.
            Else
                pdDebug.LogAction "WARNING!  Theme file <" & themeFile & "> could not be relocated.  Initialization abandoned."
                DoesThemeFileExist = False
            End If
        
        'If the theme file doesn't exist in the base folder either, we're SOL.  Exit now.
        ' (In the future, we'll want to try and recover a failsafe copy from the .exe resource file.)
        ' TODO 7.0!
        Else
            pdDebug.LogAction "WARNING!  Theme file <" & themeFile & "> wasn't found in alternate locations.  Initialization abandoned."
            DoesThemeFileExist = False
        End If
    
    End If
    
End Function

Friend Function SetNewTheme(ByVal themeClass As PD_THEME_CLASS, Optional ByVal accentColor As PD_THEME_ACCENT = PDTA_Undefined, Optional ByVal syncMenusToMatch As Boolean = False) As Boolean
    
    Dim themeName As String
    Select Case themeClass
        Case PDTC_Light
            themeName = "Light"
        Case PDTC_Dark
            themeName = "Dark"
        Case PDTC_HighContrast
            themeName = "HighContrast"
    End Select
    
    Dim accentName As String
    Select Case accentColor
        Case PDTA_Blue
            accentName = "Blue"
        Case PDTA_Green
            accentName = "Green"
        Case PDTA_Purple
            accentName = "Purple"
        Case Else
            accentName = vbNullString
    End Select
    
    m_ThemeClass = themeClass
    m_ThemeAccent = accentColor
    
    If (Not g_UserPreferences Is Nothing) Then
        If (Len(themeName) <> 0) Then g_UserPreferences.WritePreference "Themes", "CurrentTheme", themeName
        If (Len(accentName) <> 0) Then g_UserPreferences.WritePreference "Themes", "CurrentAccent", accentName
        SetNewTheme = True
    Else
        SetNewTheme = False
    End If
    
    If syncMenusToMatch Then SynchronizeThemeMenus
    
End Function

'After theme changes are made, call this sub to synchronize the main window's theme menu to the current theme's properties
Friend Sub SynchronizeThemeMenus()
    
    If g_IsProgramRunning Then
    
        Dim themeName As String, accentName As String
        themeName = g_UserPreferences.GetPref_String("Themes", "CurrentTheme", "Light")
        accentName = g_UserPreferences.GetPref_String("Themes", "CurrentAccent", "Blue")
        
        If (StrComp(LCase$(themeName), "light", vbBinaryCompare) = 0) Then
            FormMain.MnuDevelopers(2).Checked = True
            FormMain.MnuDevelopers(3).Checked = False
        Else
            FormMain.MnuDevelopers(2).Checked = False
            FormMain.MnuDevelopers(3).Checked = True
        End If
        
        Select Case accentName
        
            Case "Blue"
                FormMain.MnuDevelopers(5).Checked = True
                FormMain.MnuDevelopers(6).Checked = False
                FormMain.MnuDevelopers(7).Checked = False
            
            Case "Green"
                FormMain.MnuDevelopers(5).Checked = False
                FormMain.MnuDevelopers(6).Checked = True
                FormMain.MnuDevelopers(7).Checked = False
            
            Case Else
                FormMain.MnuDevelopers(5).Checked = False
                FormMain.MnuDevelopers(6).Checked = False
                FormMain.MnuDevelopers(7).Checked = True
        
        End Select
        
    End If

End Sub

Friend Function GetCurrentThemeClass() As PD_THEME_CLASS
    GetCurrentThemeClass = m_ThemeClass
End Function

Friend Function GetCurrentThemeAccent() As PD_THEME_ACCENT
    GetCurrentThemeAccent = m_ThemeAccent
End Function

'Load a given theme file.  Note that the filename SHOULD NOT INCLUDE THE FULL PATH - just the filename.  PD will
' automatically search the /App and /Data folders as necessary to find the file.
'
'Also, while this function does return success/failure status, if the load operation fails, PD will automatically
' fall back to its default theme to prevent the program from exploding.
Friend Function LoadThemeFile(ByVal themeFilename As String, Optional ByVal overrideColorDefinitionFilename As String = vbNullString) As Boolean
    
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
            LoadThemeFile = m_XML.IsPDDataType("Visual theme") And m_XML.ValidateLoadedXMLData("Colors")
            
        Else
            If g_IsProgramRunning Then
                RaiseThemingError "failed to load requested theme: " & fullThemePath
                RaiseThemingError "falling back to default PD theme..."
            End If
            
            fullThemePath = g_UserPreferences.GetThemePath & "Default_Light.xml"
            LoadThemeFile = m_XML.LoadXMLFile(fullThemePath)
            LoadThemeFile = LoadThemeFile And m_XML.IsPDDataType("Visual theme") And m_XML.ValidateLoadedXMLData("Colors")
            
        End If
        
        'Regardless of success or failure, reset our internal color cache(s)
        ResetColorCache
        
        'Resolve the theme class type (light, dark, or high-contrast)
        Select Case m_XML.GetUniqueTag_String("ThemeClass", "Light")
            
            Case "Light"
                m_ThemeClass = PDTC_Light
            
            Case "Dark"
                m_ThemeClass = PDTC_Dark
            
            Case "HighContrast"
                m_ThemeClass = PDTC_HighContrast
            
            Case Else
                RaiseThemingError "theme file doesn't define a theme class; assuming light theme"
                m_ThemeClass = PDTC_Light
            
        End Select
        
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
                    fullDefinitionPath = g_UserPreferences.GetThemePath & m_XML.GetUniqueTag_String("DefinitionFile")
                End If
                
                If tmpXML.LoadXMLFile(fullDefinitionPath) Then
                
                    If tmpXML.IsPDDataType("Color definitions") And tmpXML.ValidateLoadedXMLData("Definitions") Then
                    
                        'Retrieve the definition list
                        Dim colorDefinitionList As String
                        colorDefinitionList = tmpXML.GetUniqueTag_String("Definitions")
                        
                        'Plug it straight into the Definitions section of the current XML file.
                        colorDefinitionList = colorDefinitionList & m_XML.GetUniqueTag_String("Definitions")
                        
                        If Not m_XML.UpdateTag("Definitions", colorDefinitionList) Then
                            RaiseThemingError "color definition file listed in " & fullThemePath & " couldn't be dynamically inserted into parent theme"
                        End If
                    
                    Else
                        RaiseThemingError "color definition file listed in " & fullThemePath & " failed to validate"
                    End If
                
                Else
                    RaiseThemingError "color definition file listed in " & fullThemePath & " failed to load"
                End If
            
            End If
        
        End If
        
        'With all color definitions imported, we can now cache a few program-wide UI colors
        CacheUniversalColors
        
        'If the user's choice of theme didn't load correctly, or the default theme failed to load, run some heuristics
        ' on the theme folder.
        If LoadThemeFile Then
            If (Len(overrideColorDefinitionFilename) <> 0) Then
                m_CurrentThemeID = themeFilename & "-" & overrideColorDefinitionFilename
            Else
                m_CurrentThemeID = themeFilename
            End If
        Else
            RaiseThemingError "Default theme failed to load!  Catastrophic failure imminent!"
            ' (TODO: this entire step, including pulling themes from the .exe's resource section as necessary)
        End If
        
    End If
    
End Function

'Call this function to verify that an object exists inside the current theme file.  If it doesn't, you should not
' proceed with color loading.
Friend Function VerifyThemeObject(ByRef objectName As String) As Boolean
    VerifyThemeObject = m_XML.DoesTagExist(objectName)
End Function

'Look up a unique theme color in the current theme.  Object name is required, and this class will automatically fall back
' to the Default namespace as necessary.  Also, colors described by definition will automatically be tracked back to their
' source.  (Note, however, that this function has no way to deal with circular references, so please avoid that.)
' RETURNS: a color hexadecimal value if successful; a null-string otherwise.
Friend Function LookUpColor(ByVal objectName As String, ByRef colorName As String) As String

    'First things first: see if the object name exists in the theme file.  If it doesn't, we need to fall back to the
    ' "default" namespace.
    Dim objectNameExists As Boolean, objPosition As Long
    objectNameExists = m_XML.DoesTagExist(objectName, , , objPosition)
    If (Not objectNameExists) Then
        objectName = DEFAULT_NAMESPACE
        objectNameExists = m_XML.DoesTagExist(objectName, , , objPosition)
    End If
    
    'If the color exists in either the Default or object-specific namespace, we can proceed with parsing.
    If objectNameExists Then
        
        'Inside the current object's color definition block, retrieve the specified color
        Dim colorDescription As String, finalColor As String
        colorDescription = m_XML.GetNonUniqueTag_String(colorName, objectName, , , , objPosition)
        
        'If we retrieved any valid string, attempt to resolve it to an actual color value.  (At this point, the color
        ' may just be a variable instead of an actual hex value.)
        If (Len(colorDescription) <> 0) Then
            finalColor = ResolveColor(colorDescription)
        
        'If we used a custom object name, but no color is defined for that value, try a new retrieval from
        ' the "Default" namespace.  (Empty colors are still valid, as long as their Default variant is defined.)
        Else
            If (StrComp(objectName, DEFAULT_NAMESPACE, vbBinaryCompare) <> 0) Then
                objectName = DEFAULT_NAMESPACE
                If m_XML.DoesTagExist(objectName, , , objPosition) Then
                    colorDescription = m_XML.GetNonUniqueTag_String(colorName, objectName, , , , objPosition)
                    If (Len(colorDescription) <> 0) Then finalColor = ResolveColor(colorDescription)
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
        If (Len(initialColorValue) = 0) Then
            ResolveColor = vbNullString
            Exit Function
        End If
        
        'Next, see if the current color value appears to be some kind of valid color representation
        If Colors.IsStringAColor(initialColorValue, , False) Then
            ResolveColor = initialColorValue
            Exit Function
        
        'This is not a valid color representation, so assume it's a custom color descriptor (or invalid, I suppose)
        Else
            
            'Attempt to retrieve a new value from the theme's color definition section, then run our validation
            ' checks a second time.  (We'll repeat this until we fail to retrieve a new definition, or we identify
            ' a string that can be parsed into an actual color.)
            initialColorValue = m_XML.GetUniqueTag_String(initialColorValue, vbNullString, , "Definitions")
            
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
    
    With m_ColorCache(m_NumColorsCached)
        .OrigObjectName = objectName
        .OrigColorName = colorName
        .OrigColorValues = srcColorEntry
    End With
    
    m_NumColorsCached = m_NumColorsCached + 1
    If (m_NumColorsCached > UBound(m_ColorCache)) Then ReDim Preserve m_ColorCache(0 To m_NumColorsCached * 2 - 1) As PDCachedColor
    
End Sub

'Look up a color in the color cache.  If it exists, the function returns TRUE, and the destination PDThemeColor struct
' is filled with the matching cache values.
Friend Function RetrieveColorFromCache(ByRef objectName As String, ByRef colorName As String, ByRef dstColorEntry As PDThemeColor) As Boolean

    RetrieveColorFromCache = False
    
    Dim i As Long
    For i = 0 To m_NumColorsCached - 1
        If (StrComp(objectName, m_ColorCache(i).OrigObjectName, vbBinaryCompare) = 0) Then
            If (StrComp(colorName, m_ColorCache(i).OrigColorName, vbBinaryCompare) = 0) Then
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
        .LoadThemeColor UI_ChannelRed, "UniversalChannelRed", RGB(255, 0, 0)
        .LoadThemeColor UI_ChannelGreen, "UniversalChannelGreen", RGB(0, 255, 0)
        .LoadThemeColor UI_ChannelBlue, "UniversalChannelBlue", RGB(0, 0, 255)
        .LoadThemeColor UI_ErrorRed, "UniversalErrorRed", RGB(255, 0, 0)
        .LoadThemeColor UI_GrayDefault, "UniversalGrayDefault", IDE_GRAY
        .LoadThemeColor UI_GrayDisabled, "UniversalGrayDisabled", IDE_GRAY
        .LoadThemeColor UI_GrayDark, "UniversalGrayDark", IDE_GRAY
        .LoadThemeColor UI_GrayLight, "UniversalGrayLight", IDE_GRAY
        .LoadThemeColor UI_GrayNeutral, "UniversalGrayNeutral", IDE_GRAY
        .LoadThemeColor UI_GrayUltraLight, "UniversalGrayUltralight", IDE_GRAY
        .LoadThemeColor UI_IconMonochrome, "UniversalIconMonochrome", IDE_BLACK
        .LoadThemeColor UI_IconMonochromeMenu, "UniversalIconMenuMonochrome", IDE_BLACK
        .LoadThemeColor UI_ImageDisabled, "UniversalImageDisabled", IDE_GRAY
        .LoadThemeColor UI_LineEdge, "UniversalLineEdge", IDE_BLUE
        .LoadThemeColor UI_LineCenter, "UniversalLineCenter", IDE_GRAY
        .LoadThemeColor UI_TextClickable, "UniversalTextClickable", IDE_BLUE
        .LoadThemeColor UI_TextClickableSelected, "UniversalTextClickableSelected", IDE_WHITE
        .LoadThemeColor UI_TextClickableUnselected, "UniversalTextClickableUnselected", IDE_GRAY
        .LoadThemeColor UI_TextReadOnly, "UniversalTextReadOnly", IDE_BLACK
    End With
    
End Sub

'External functions can use this to retrieve a color from the local m_UniversalColors cache.  If an object requires
' a bunch of object-specific colors, they will get better performance by managing their own color cache.
Friend Function GetGenericUIColor(ByVal colorID As PD_UI_COLOR_LIST, Optional ByVal enabledState As Boolean = True, Optional ByVal activeState As Boolean = False, Optional ByVal hoverState As Boolean = False) As Long
    GetGenericUIColor = m_UniversalColors.RetrieveColor(colorID, enabledState, activeState, hoverState)
End Function

Friend Sub AddWindowPainter(ByVal srcHwnd As Long)
    If g_IsProgramRunning Then
        If (m_Subclass Is Nothing) Then Set m_Subclass = New cSelfSubHookCallback
        m_Subclass.ssc_Subclass srcHwnd, 0, 1, Me
        m_Subclass.ssc_AddMsg srcHwnd, MSG_BEFORE, WM_PAINT, WM_ERASEBKGND
    End If
End Sub

Friend Sub RemoveWindowPainter(ByVal srcHwnd As Long)
    If g_IsProgramRunning Then m_Subclass.ssc_UnSubclass srcHwnd
End Sub

Private Sub RaiseThemingError(ByVal msgError As String, Optional ByVal msgIsNonErrorFeedback As Boolean = False)
    #If DEBUGMODE = 1 Then
        If msgIsNonErrorFeedback Then
            pdDebug.LogAction "pdVisualThemes reported: " & msgError
        Else
            pdDebug.LogAction "WARNING!  pdVisualThemes error: " & msgError
        End If
    #End If
End Sub

Private Sub Class_Initialize()
    Set m_XML = New pdXML
    m_XML.SetTextCompareMode vbBinaryCompare
    Set m_UniversalColors = New pdThemeColors
End Sub

Private Sub Class_Terminate()
    If (Not m_Subclass Is Nothing) Then Set m_Subclass = Nothing
End Sub

'All messages subclassed by m_Subclass are handled here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    Dim tmpDC As Long
    
    If (uMsg = WM_PAINT) Then
        
        'Ignore paint requests for empty regions
        If GetUpdateRect(hWnd, m_UpdateRect, 0) <> 0 Then
        
            tmpDC = BeginPaint(hWnd, m_PaintStruct)
            
            With m_UpdateRect
                GDI.FillRectToDC tmpDC, .Left, .Top, .Right - .Left, .Bottom - .Top, g_Themer.GetGenericUIColor(UI_Background)
            End With
            
            'End painting (note: BeginPaint automatically validated the window's contents, so we don't need to do any
            ' additional validation here)
            EndPaint hWnd, m_PaintStruct
            
        End If
        
        'Mark the message as handled and exit
        bHandled = True
        lReturn = 0
                
    ElseIf (uMsg = WM_ERASEBKGND) Then
        
        tmpDC = wParam
        GetClientRect hWnd, m_UpdateRect
        With m_UpdateRect
            GDI.FillRectToDC tmpDC, .Left, .Top, .Right - .Left, .Bottom - .Top, g_Themer.GetGenericUIColor(UI_Background)
        End With
        
        bHandled = True
        lReturn = 1
        
    End If

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


