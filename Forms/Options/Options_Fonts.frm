VERSION 5.00
Begin VB.Form options_Fonts 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "options_Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Fonts panel
'Copyright 2025-2025 by Tanner Helland
'Created: 04/April/25
'Last updated: 04/April/25
'Last update: initial build
'
'This form contains a single subpanel worth of program options.  At run-time, it is dynamically
' made a child of FormOptions.  It will only be loaded if/when the user interacts with this category.
'
'All Tools > Options child panels must some mandatory public functions, including ones for loading
' and saving user preferences, as well as validating any UI elements where the user can enter
' custom values.  (A reset-style function is *not* required; this is automatically handled by
' FormOptions.)
'
'This form, like all Tools > Options panels, interacts heavily with the UserPrefs module.
' (That module is responsible for all low-level preference reading/writing.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub Form_Load()
'
'    'Interface prefs
'    btsTitleText.AddItem "compact (filename only)", 0
'    btsTitleText.AddItem "verbose (filename and path)", 1
'    btsTitleText.AssignTooltip "The title bar of the main PhotoDemon window displays information about the currently loaded image.  Use this preference to control how much information is displayed."
'
'    lblCanvasColor.Caption = g_Language.TranslateMessage("canvas background color: ")
'    csCanvasColor.SetLeft lblCanvasColor.GetLeft + lblCanvasColor.GetWidth + Interface.FixDPI(8)
'    csCanvasColor.SetWidth (btsTitleText.GetLeft + btsTitleText.GetWidth) - (csCanvasColor.GetLeft)
'
'    lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent files to remember: ")
'    tudRecentFiles.SetLeft lblRecentFileCount.GetLeft + lblRecentFileCount.GetWidth + Interface.FixDPI(8)
'
'    btsMRUStyle.AddItem "compact (filename only)", 0
'    btsMRUStyle.AddItem "verbose (filename and path)", 1
'    btsMRUStyle.AssignTooltip "The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations."
'
'    chkZoomMouse.Value = UserPrefs.GetPref_Boolean("Interface", "wheel-zoom", False)
'
'    m_userInitiatedAlphaSelection = False
'    cboAlphaCheck.Clear
'    cboAlphaCheck.AddItem "highlights", 0
'    cboAlphaCheck.AddItem "midtones", 1
'    cboAlphaCheck.AddItem "shadows", 2, True
'    cboAlphaCheck.AddItem "red", 3
'    cboAlphaCheck.AddItem "orange", 4
'    cboAlphaCheck.AddItem "green", 5
'    cboAlphaCheck.AddItem "blue", 6
'    cboAlphaCheck.AddItem "purple", 7, True
'    cboAlphaCheck.AddItem "custom", 8
'    cboAlphaCheck.AssignTooltip "To help identify transparent pixels, a special grid appears ""behind"" them.  This setting modifies the grid's appearance."
'    m_userInitiatedAlphaSelection = True
'
'    cboAlphaCheckSize.Clear
'    cboAlphaCheckSize.AddItem "small", 0
'    cboAlphaCheckSize.AddItem "medium", 1
'    cboAlphaCheckSize.AddItem "large", 2
'    cboAlphaCheckSize.AssignTooltip "To help identify transparent pixels, a special grid appears ""behind"" them.  This setting modifies the grid's appearance."
'
End Sub

Public Sub LoadUserPreferences()
'
'    'Interface preferences
'    btsTitleText.ListIndex = UserPrefs.GetPref_Long("Interface", "Window Caption Length", 0)
'    csCanvasColor.Color = UserPrefs.GetCanvasColor()
'    tudRecentFiles.Value = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
'    btsMRUStyle.ListIndex = UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0)
'    spnSnapDistance(0).Value = UserPrefs.GetPref_Long("Interface", "snap-distance", 8&)
'    spnSnapDistance(1).Value = UserPrefs.GetPref_Float("Interface", "snap-degrees", 7.5)
'    m_userInitiatedAlphaSelection = False
'    cboAlphaCheck.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Mode", 0)
'    csAlphaOne.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check One", RGB(255, 255, 255))
'    csAlphaTwo.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check Two", RGB(204, 204, 204))
'    m_userInitiatedAlphaSelection = True
'    cboAlphaCheckSize.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Size", 1)
'    UpdateAlphaGridVisibility
'
End Sub

Public Sub SaveUserPreferences()
'
'    'Interface preferences
'    UserPrefs.SetPref_Long "Interface", "Window Caption Length", btsTitleText.ListIndex
'    UserPrefs.SetPref_String "Interface", "Canvas Color", Colors.GetHexStringFromRGB(csCanvasColor.Color)
'    UserPrefs.SetCanvasColor csCanvasColor.Color
'
'    'Changes to the recent files list (including count and how it's displayed) may require us to
'    ' trigger a full rebuild of the menu
'    Dim mruNeedsToBeRebuilt As Boolean
'    mruNeedsToBeRebuilt = (btsMRUStyle.ListIndex <> UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0))
'    UserPrefs.SetPref_Long "Interface", "MRU Caption Length", btsMRUStyle.ListIndex
'
'    Dim newMaxRecentFiles As Long
'    If tudRecentFiles.IsValid Then newMaxRecentFiles = tudRecentFiles.Value Else newMaxRecentFiles = 10
'    If (Not mruNeedsToBeRebuilt) Then mruNeedsToBeRebuilt = (newMaxRecentFiles <> UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10))
'    UserPrefs.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value
'
'    'If any MRUs need to be rebuilt, do so now
'    If mruNeedsToBeRebuilt Then
'        g_RecentFiles.NotifyMaxLimitChanged
'        g_RecentMacros.MRU_NotifyNewMaxLimit
'    End If
'
'    UserPrefs.SetPref_Boolean "Interface", "wheel-zoom", chkZoomMouse.Value
'    UserPrefs.SetZoomWithWheel chkZoomMouse.Value
'
'    UserPrefs.SetPref_Long "Interface", "snap-distance", spnSnapDistance(0).Value
'    UserPrefs.SetPref_Long "Interface", "snap-degrees", spnSnapDistance(1).Value
'    Snap.SetSnap_Distance spnSnapDistance(0).Value
'    Snap.SetSnap_Degrees spnSnapDistance(1).Value
'
'    UserPrefs.SetPref_Long "Transparency", "Alpha Check Mode", CLng(cboAlphaCheck.ListIndex)
'    UserPrefs.SetPref_Long "Transparency", "Alpha Check One", CLng(csAlphaOne.Color)
'    UserPrefs.SetPref_Long "Transparency", "Alpha Check Two", CLng(csAlphaTwo.Color)
'    UserPrefs.SetPref_Long "Transparency", "Alpha Check Size", cboAlphaCheckSize.ListIndex
'    Drawing.CreateAlphaCheckerboardDIB g_CheckerboardPattern
'
End Sub

'Upon calling, validate all input.  Return FALSE if validation on 1+ controls fails.
Public Function ValidateAllInput() As Boolean
    
    ValidateAllInput = True
    
    Dim eControl As Object
    For Each eControl In Me.Controls
        
        'Most UI elements on this dialog are idiot-proof, but spin controls (including those embedded
        ' in slider controls) are an exception.
        If (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSpinner) Then
            
            'Finally, ask the control to validate itself
            If (Not eControl.IsValid) Then
                ValidateAllInput = False
                Exit For
            End If
            
        End If
    Next eControl
    
End Function

'This function is called at least once, immediately following Form_Load(),
' but it can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    Interface.ApplyThemeAndTranslations Me
End Sub
