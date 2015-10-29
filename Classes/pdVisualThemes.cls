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
'Copyright 2013-2015 by Tanner Helland
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

'cSubclass handles all subclassing duties for us
Private cSubclass As cSelfSubHookCallback

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

'When a container object needs to render contained controls transparently, register it using this function
Public Sub RequestContainerSubclass(ByVal containerHwnd As Long)

    If (cSubclass Is Nothing) Then
        Set cSubclass = New cSelfSubHookCallback
    End If
    
    cSubclass.ssc_Subclass containerHwnd, , , Me
    cSubclass.ssc_AddMsg containerHwnd, MSG_BEFORE, WM_PRINTCLIENT

End Sub

'While not strictly required, it is good behavior to unsubclass container objects when they are unloaded
Public Sub ReleaseContainerSubclass(ByVal containerHwnd As Long)
    cSubclass.ssc_UnSubclass containerHwnd
End Sub

'Reset the theme color cache to all -1 values.
Private Sub ResetThemeCache()

    Dim i As Long
    For i = 0 To NUM_THEME_COLORS - 1
        curThemeColors(i) = -1
    Next i

End Sub

'Fill the curThemeColors array with PD's default color scheme
Private Sub LoadDefaultPDTheme()

    'Default to the light theme
    curThemeContrast = THEME_LIGHT
    
    'Loop through the curThemeColors array, filling each entry with the PD default
    Dim i As Long
    For i = 0 To NUM_THEME_COLORS - 1
        curThemeColors(i) = GetDefaultPDColor(i)
    Next i

End Sub

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
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("3296dc")
            
            'Slider track, large contiguous blocks of color
            Case PDTC_ACCENT_HIGHLIGHT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#3cafe6")
            
            'Background on toolbox images (which can't afford to be as dark as the usual highlight value)
            Case PDTC_ACCENT_ULTRALIGHT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("d2f0fa")
            
            'Small or heavily aliased items (e.g. check box check, radio button dot)
            Case PDTC_ACCENT_SHADOW
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#3270dc")
            
            
            'Non-interactive variations (used to render control outlines, etc)
            Case PDTC_GRAY_DEFAULT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#929292")
                
            Case PDTC_GRAY_HIGHLIGHT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#bababa")
            
            Case PDTC_GRAY_ULTRALIGHT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#e2e2e2")
            
            Case PDTC_GRAY_SHADOW
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#5b5b5b")
            
            'Background variations
            Case PDTC_BACKGROUND_DEFAULT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
                
            Case PDTC_BACKGROUND_CANVAS
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#a0a0a0")
                
            Case PDTC_BACKGROUND_COMMANDBAR
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#dcdce1")
            
            'Text variations
            Case PDTC_TEXT_DEFAULT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#404040")
                
            Case PDTC_TEXT_EDITBOX
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#202020")
                
            Case PDTC_TEXT_HYPERLINK
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#3170c0")
                
            Case PDTC_TEXT_TITLE
            
            Case PDTC_TEXT_INVERT
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#ffffff")
            
            'Misc colors
            Case PDTC_DISABLED
                GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#b1bac2")
            
            Case PDTC_OK_GREEN
            Case PDTC_CANCEL_RED
            Case PDTC_FOCUS_RECT
            
            Case PDTC_CHANNEL_RED
                    GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#dc2824")
                    
            Case PDTC_CHANNEL_GREEN
                    GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#3ccf31")
                    
            Case PDTC_CHANNEL_BLUE
                    GetDefaultPDColor = Color_Functions.getRGBLongFromHex("#317bcf")
            
        End Select
    
    'DARK THEME DEFAULTS
    ElseIf curThemeContrast = THEME_DARK Then
    
    'HIGH CONTRAST DEFAULTS
    Else
    
    End If

End Function

Private Sub Class_Initialize()

    'Prepare our color cache
    ReDim curThemeColors(0 To NUM_THEME_COLORS - 1) As Long
    
    'Reset the color cache.  This is important, because 0 is a valid color, so we can't use it to indicate a
    ' missing theme entry.  Instead, we use -1 to represent a missing entry.  The color retrieval function
    ' will automatically cascade to the nearest relevant value if an entry can't be found.
    ResetThemeCache
    
    'Populate the color cache with PD's default theme values.  If something goes wrong when attempting to load the
    ' user's theme of choice, this at least ensures that we have a valid fallback.
    LoadDefaultPDTheme
    
End Sub

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
