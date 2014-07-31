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
'Copyright ©2013-2014 by Tanner Helland
'Created: 23/October/13
'Last updated: 26/July/14
'Last update: start centralizing color definitions
'
'There isn't much here yet, but stay tuned - more is coming!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'In order to make themed controls transparent, we must subclass their parent containers, and manually process the
' WM_PRINTCLIENT message using the API decs below.  Many thanks to pro VB programmer LaVolpe for this workaround.
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_PAINT As Long = &HF&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'I have no idea what colors will actually be used in the final theme format; these are just placeholders for now,
' to enable centralized coloration of a few UCs.  The end goal is to provide the smallest possible subset of colors
' necessary to render PD's various UI elements.
Private Const numThemeColors As Long = 14
Private Const numThemeColorVariants As Long = 5

Public Enum PD_THEME_COLOR
    PDTC_ACCENT_DEFAULT = 0
    PDTC_ACCENT_INTERACTIVE = 1
    PDTC_ACCENT_INVALID = 2
    PDTC_ACCENT_NONINTERACTIVE = 3
    PDTC_ACCENT_OK = 4
    
    PDTC_BACKGROUND_CANVAS = 5
    PDTC_BACKGROUND_COMMANDBAR = 6
    PDTC_BACKGROUND_DEFAULT = 7
    PDTC_BACKGROUND_DIALOG = 8
    
    PDTC_TEXT_DEFAULT = 9
    PDTC_TEXT_HYPERLINK = 10
    PDTC_TEXT_INTERACTIVE = 11
    PDTC_TEXT_LABEL = 12
    PDTC_TEXT_TITLE = 13
End Enum

#If False Then
    Private Const PDTC_ACCENT_DEFAULT = 0, PDTC_ACCENT_INTERACTIVE = 1, PDTC_ACCENT_INVALID = 2, PDTC_ACCENT_NONINTERACTIVE = 3, PDTC_ACCENT_OK = 4
    Private Const PDTC_BACKGROUND_CANVAS = 5, PDTC_BACKGROUND_COMMANDBAR = 6, PDTC_BACKGROUND_DEFAULT = 7
    Private Const PDTC_TEXT_DEFAULT = 9, PDTC_TEXT_HYPERLINK = 10, PDTC_TEXT_INTERACTIVE = 11, PDTC_TEXT_LABEL = 12, PDTC_TEXT_TITLE = 13
#End If

Public Enum PD_THEME_COLOR_VARIANT
    PDTCV_NORMAL = 0
    PDTCV_HIGHLIGHT = 1
    PDTCV_INVERT = 2
    PDTCV_SHADOW = 3
    PDTCV_DISABLED = 4
End Enum

#If False Then
    Private Const PDTCV_NORMAL = 0, PDTCV_HIGHLIGHT = 1, PDTCV_INVERT = 2, PDTCV_SHADOW = 3, PDTCV_DISABLED = 4
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

'cSubclass handles all subclassing duties for us
Private cSubclass As cSelfSubHookCallback

'Despite an identical name, this function does *not* use the Win32 API to retrieve a color from the system theme.
' Instead, it retrieves a color from PD's current theme color cache.
Friend Function getThemeColor(ByVal colorType As PD_THEME_COLOR, Optional ByVal colorTypeVariant As PD_THEME_COLOR_VARIANT = PDTCV_NORMAL) As Long
    
    'Back up the original color type and variant, as we may make use of it if the requested color does not exist,
    ' and we have to retrieve a match from PD's default theme color set.
    Dim origColorType As PD_THEME_COLOR, origColorTypeVariant As PD_THEME_COLOR_VARIANT
    origColorType = colorType
    origColorTypeVariant = colorTypeVariant
    
    'Returning a theme color is fairly trivial.  Start by seeing if the requested variant is available.
    getThemeColor = curThemeColors(colorType, colorTypeVariant)
    
    Dim valueDoesNotExist As Boolean
    valueDoesNotExist = False
    
    'If the retrieved value is -1, that means that the current theme does not supply a value for this type/variant combination.
    ' (This isn't a bad thing; in fact, themes are not expected to provide variants for certain color types.)
    If getThemeColor = -1 Then valueDoesNotExist = True
    
    'Because we use a cascading replacement system for color variations, we may need to cascade through several steps
    ' in order to find a valid replacement color.  As such, we use a loop instead of an If/Then structure.
    Do While valueDoesNotExist
    
        'If this is a default color type, we've got a problem.  Those values are supposed to always exist in a theme!
        ' As a last resort, replace the color with PD's default color, at the originally requested variant.
        If ((colorType = PDTC_ACCENT_DEFAULT) Or (colorType = PDTC_BACKGROUND_DEFAULT) Or (colorType = PDTC_TEXT_DEFAULT)) Then
        
            'Start by replacing the color type with the original color type.  (It may have changed during a cascading search.)
            colorType = origColorType
            colorTypeVariant = origColorTypeVariant
            
            'HIGHLIGHT and SHADOW variants are not crucial for most tasks.  Try to replace them with the NORMAL variant.
            If (colorTypeVariant = PDTCV_HIGHLIGHT) Or (colorTypeVariant = PDTCV_SHADOW) Then
                getThemeColor = curThemeColors(colorType, PDTCV_NORMAL)
                If getThemeColor >= 0 Then Exit Function
            End If
            
            'If we made it all the way here, a color of the NORMAL, INVERT, or DISABLED types has been requested, but it
            ' could not be located.  Default to the relevant PD default theme value.
            getThemeColor = getDefaultPDColor(colorType, colorTypeVariant)
            Exit Function
            
        'For non-normal variations, try to replace the request with the same variant of a related type.
        ' Generally speaking, we try the following formula:
        ' 1) If it's an obscure color type, try replacing it with a more common type, but same variant.
        ' 2) If it's a common type, but a weird variant, try replacing it with a more normal variant of that type.
        Else
            
            Select Case colorType
            
                'ACCENT variants
                
                'Interactive accents (generally blue) can be replaced by the default accent color
                Case PDTC_ACCENT_INTERACTIVE
                    colorType = PDTC_ACCENT_DEFAULT
                    
                '"INVALID" accents (generally red) can be replaced by the non-interactive accent color
                Case PDTC_ACCENT_INVALID
                    colorType = PDTC_ACCENT_NONINTERACTIVE
                    
                '"NONINTERACTIVE" accents (generally a more subdued version of the interactive accent) can be
                ' replaced by the interactive variant, as necessary.
                Case PDTC_ACCENT_NONINTERACTIVE
                    colorType = PDTC_ACCENT_INTERACTIVE
                    
                '"OK" accents (generally green) can be replaced by the interactive accent color
                Case PDTC_ACCENT_OK
                    colorType = PDTC_ACCENT_INTERACTIVE
                    
                'BACKGROUND variants
                
                'Various background types will all revert to the default background color if they can't be found
                Case PDTC_BACKGROUND_CANVAS, PDTC_BACKGROUND_COMMANDBAR, PDTC_BACKGROUND_DIALOG
                    colorType = PDTC_BACKGROUND_DEFAULT
                    
                'TEXT variants
                
                '"LABELS" revert to the default color type
                Case PDTC_TEXT_LABEL
                    colorType = PDTC_TEXT_DEFAULT
                    
                '"HYPERLINKS" revert to the default interactive color
                Case PDTC_TEXT_HYPERLINK
                    colorType = PDTC_TEXT_INTERACTIVE
                    
                '"INTERACTIVES" revert to the default color type
                Case PDTC_TEXT_INTERACTIVE
                    colorType = PDTC_TEXT_DEFAULT
                    
                '"TITLES" revert to the default "LABEL" color
                Case PDTC_TEXT_TITLE
                    colorType = PDTC_TEXT_LABEL
                    
            End Select
            
            'Retrieve the substitute color type
            getThemeColor = curThemeColors(colorType, colorTypeVariant)
            
        End If
        
        'If the retrieved color still does not exist, we'll need to cascade further
        If getThemeColor >= 0 Then valueDoesNotExist = False Else valueDoesNotExist = True
    
    Loop
    
End Function

'When a container object needs to render contained controls transparently, register it using this function
Public Sub requestContainerSubclass(ByVal containerHwnd As Long)

    If (cSubclass Is Nothing) Then
        Set cSubclass = New cSelfSubHookCallback
    End If
    
    cSubclass.ssc_Subclass containerHwnd, , , Me
    cSubclass.ssc_AddMsg containerHwnd, MSG_BEFORE, WM_PRINTCLIENT

End Sub

'While not strictly required, it is good behavior to unsubclass container objects when they are unloaded
Public Sub releaseContainerSubclass(ByVal containerHwnd As Long)
    cSubclass.ssc_UnSubclass containerHwnd
End Sub

Private Sub Class_Initialize()

    'Prepare our color cache
    ReDim curThemeColors(0 To numThemeColors - 1, 0 To numThemeColorVariants - 1) As Long
    
    'Reset the color cache.  This is important, because 0 is a valid color, so we can't use it to indicate a
    ' missing theme entry.  Instead, we use -1 to represent a missing entry.  The color retrieval function
    ' will automatically cascade to the nearest relevant value if an entry can't be found.
    resetThemeCache
    
    'Populate the color cache with PD's default theme values.  If something goes wrong when attempting to load the
    ' user's theme of choice, this at least ensures that we have a valid fallback.
    loadDefaultPDTheme

End Sub

'Reset the theme color cache to all -1 values.
Private Sub resetThemeCache()

    Dim i As Long, j As Long
    For i = 0 To numThemeColors - 1
    For j = 0 To numThemeColorVariants - 1
        curThemeColors(i, j) = -1
    Next j
    Next i

End Sub

'Fill the curThemeColors array with PD's default color scheme
Private Sub loadDefaultPDTheme()

    'Default to the light theme
    curThemeContrast = THEME_LIGHT
    
    'Loop through the curThemeColors array, filling each entry with the PD default
    Dim i As Long, j As Long
    For i = 0 To numThemeColors - 1
    For j = 0 To numThemeColorVariants - 1
        curThemeColors(i, j) = getDefaultPDColor(i, j)
    Next j
    Next i

End Sub

'Retrieve a default PD theme color.  These values are hard-coded, and will be superceded by any values supplied
' by a user-selected theme.
Private Function getDefaultPDColor(ByVal colorType As PD_THEME_COLOR, Optional ByVal colorTypeVariant As PD_THEME_COLOR_VARIANT = PDTCV_NORMAL) As Long

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
        
            'Accent colors
            Case PDTC_ACCENT_DEFAULT
                
            Case PDTC_ACCENT_INTERACTIVE
            
                If (colorTypeVariant = PDTCV_NORMAL) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#3296dc")
                    
                ElseIf (colorTypeVariant = PDTCV_HIGHLIGHT) Then
                    'getDefaultPDColor = Color_Functions.getRGBLongFromHex("#3296dc")
                    
                ElseIf (colorTypeVariant = PDTCV_SHADOW) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#3296dc")
                    
                ElseIf (colorTypeVariant = PDTCV_INVERT) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
                
                ElseIf (colorTypeVariant = PDTCV_DISABLED) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#b1bac2")
                
                End If
            
            Case PDTC_ACCENT_INVALID
            
            Case PDTC_ACCENT_NONINTERACTIVE
            
                If (colorTypeVariant = PDTCV_NORMAL) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#7e8c92")
                    
                ElseIf (colorTypeVariant = PDTCV_HIGHLIGHT) Or (colorTypeVariant = PDTCV_SHADOW) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#7e8c92")
                    
                ElseIf (colorTypeVariant = PDTCV_INVERT) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
                
                ElseIf (colorTypeVariant = PDTCV_DISABLED) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#b1bac2")
                
                End If
            
            Case PDTC_ACCENT_OK
            
            'Background colors
            Case PDTC_BACKGROUND_CANVAS
            
            Case PDTC_BACKGROUND_COMMANDBAR
            
            Case PDTC_BACKGROUND_DEFAULT
                getDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
            
            Case PDTC_BACKGROUND_DIALOG
            
            'Text colors
            Case PDTC_TEXT_DEFAULT
                
                If (colorTypeVariant = PDTCV_NORMAL) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#404040")
                    
                ElseIf (colorTypeVariant = PDTCV_HIGHLIGHT) Or (colorTypeVariant = PDTCV_SHADOW) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#3170c0")
                
                ElseIf (colorTypeVariant = PDTCV_INVERT) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
                
                ElseIf (colorTypeVariant = PDTCV_DISABLED) Then
                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#b1bac2")
                
                End If
            
            Case PDTC_TEXT_HYPERLINK
            
            Case PDTC_TEXT_INTERACTIVE
            
'                If (colorTypeVariant = PDTCV_NORMAL) Or (colorTypeVariant = PDTCV_HIGHLIGHT) Or (colorTypeVariant = PDTCV_SHADOW) Then
'                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#3170c0")
'
'                ElseIf (colorTypeVariant = PDTCV_INVERT) Then
'                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#FFFFFF")
'
'                ElseIf (colorTypeVariant = PDTCV_DISABLED) Then
'                    getDefaultPDColor = Color_Functions.getRGBLongFromHex("#808080")
'
'                End If
            
            Case PDTC_TEXT_LABEL
            
            Case PDTC_TEXT_TITLE
        
        End Select
    
    'DARK THEME DEFAULTS
    ElseIf curThemeContrast = THEME_DARK Then
    
    'HIGH CONTRAST DEFAULTS
    Else
    
    End If

End Function

'All events subclassed by this class are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
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

    If uMsg = WM_PRINTCLIENT Then
        SendMessage lng_hWnd, WM_PAINT, wParam, ByVal 0&
        bHandled = True
    End If

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
End Sub
