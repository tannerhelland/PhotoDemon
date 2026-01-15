VERSION 5.00
Begin VB.Form FormOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Options"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
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
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   6870
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonStripVertical btsvCategory 
      Height          =   6675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   11774
   End
End
Attribute VB_Name = "FormOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options Handler
'Copyright 2002-2026 by Tanner Helland
'Created: 8/November/02
'Last updated: 29/March/25
'Last update: total overhaul of this dialog; individual panels are now split into standalone forms,
'             and each is loaded at run-time if/when the user interacts with it
'
'Dialog for interfacing with the user's desired program options.
'
'This form interacts heavily with the UserPrefs module.  (That module is also responsible for all
' low-level reading/writing of preferences; this is just the UI for interacting with it.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'In 2025, I removed all options panels from *this* form and separated them out into standalone forms.
' This greatly improves load-times for this dialog, cleans up the code a *lot* (because specialized
' preference loading/saving behavior and/or UI considerations are now split across many forms).
'
'Because of this new organization, panels need to be dynamically loaded and positioned at run-time.
Private Enum PD_OptionPanel
    OP_None = -1
    OP_Interface = 0
    OP_Menus = 1
    OP_Performance = 2
    OP_Loading = 3
    OP_Saving = 4
    OP_Metadata = 5
    OP_ColorManagement = 6
    OP_Fonts = 7
    OP_InputDevices = 8
    OP_Updates = 9
    OP_Advanced = 10
End Enum

#If False Then
    Private Const OP_None = -1, OP_Interface = 0, OP_Menus = 1, OP_Performance = 2, OP_Loading = 3, OP_Saving = 4, OP_Metadata = 5
    Private Const OP_ColorManagement = 6, OP_Fonts = 7, OP_InputDevices = 8, OP_Updates = 9, OP_Advanced = 10
#End If

Private Const MAX_NUM_OPTION_PANELS As Long = OP_Advanced + 1

'Because options panels are not loaded unless the user interacts with them, it's likely that only *some*
' panels will be touched in a given session.  To improve performance, we only save options from a given
' panel *if* that panel was touched this session.
Private Type OptionPanelTracker
    PanelHWnd As Long
    PanelWasLoaded As Boolean
End Type

Private m_numOptionPanels As Long, m_Panels() As OptionPanelTracker

'Current and previously active panels are mirrored here
Private m_ActivePanel As PD_OptionPanel
Private m_ActivePanelHWnd As Long

'Position and dimensions of the panel; populated only after the first child form has been loaded and displayed
Private m_PanelRect As RectL_WH

'When the preferences category is changed, only display the controls in that category
Private Sub btsvCategory_Click(ByVal buttonIndex As Long)
    UpdateActivePanel buttonIndex
End Sub

Private Sub UpdateActivePanel(ByVal idxPanel As Long)
 
    'Hide currently active panel (if any), load new panel
    m_ActivePanel = idxPanel
    
    'If our panel tracker doesn't exist, create it now
    If (m_numOptionPanels = 0) Then
        m_numOptionPanels = MAX_NUM_OPTION_PANELS
        ReDim m_Panels(0 To m_numOptionPanels - 1) As OptionPanelTracker
    End If
    
    'Next, we need to display the correct preferences panel.
    If (Not m_Panels(m_ActivePanel).PanelWasLoaded) Then
    
        Select Case m_ActivePanel
            
            'Move/size tool
            Case OP_Interface
                Load options_Interface
                options_Interface.LoadUserPreferences
                options_Interface.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Interface.hWnd
            
            Case OP_Menus
                Load options_Menus
                options_Menus.LoadUserPreferences
                options_Menus.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Menus.hWnd
            
            Case OP_Performance
                Load options_Performance
                options_Performance.LoadUserPreferences
                options_Performance.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Performance.hWnd
                
            Case OP_Loading
                Load options_Loading
                options_Loading.LoadUserPreferences
                options_Loading.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Loading.hWnd
                
            Case OP_Saving
                Load options_Saving
                options_Saving.LoadUserPreferences
                options_Saving.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Saving.hWnd
            
            Case OP_Metadata
                Load options_Metadata
                options_Metadata.LoadUserPreferences
                options_Metadata.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Metadata.hWnd
            
            Case OP_ColorManagement
                Load options_ColorManagement
                options_ColorManagement.LoadUserPreferences
                options_ColorManagement.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_ColorManagement.hWnd
            
            Case OP_Fonts
                Load options_Fonts
                options_Fonts.LoadUserPreferences
                options_Fonts.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Fonts.hWnd
                
            Case OP_InputDevices
                Load options_Input
                options_Input.LoadUserPreferences
                options_Input.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Input.hWnd
                
            Case OP_Updates
                Load options_Updates
                options_Updates.LoadUserPreferences
                options_Updates.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Updates.hWnd
                
            Case OP_Advanced
                Load options_Advanced
                options_Advanced.LoadUserPreferences
                options_Advanced.UpdateAgainstCurrentTheme
                m_Panels(m_ActivePanel).PanelHWnd = options_Advanced.hWnd
                
            Case Else
                m_ActivePanel = OP_None
                
        End Select
        
        If (m_ActivePanel <> OP_None) Then m_Panels(m_ActivePanel).PanelWasLoaded = True
        
    End If
        
    'Next, we want to display the current options panel, while hiding all inactive ones.
    ' (This must be handled carefully, or we risk accidentally enabling unloaded panels,
    '  which we don't want as option panels are quite resource-heavy.)
    If (m_ActivePanelHWnd <> 0&) Then g_WindowManager.DeactivateToolPanel m_ActivePanelHWnd
    m_ActivePanelHWnd = 0&
    
    'To prevent flicker, we handle this in two passes.
    
    'First, activate the new window.
    If (m_numOptionPanels <> 0) Then
        
        Dim i As Long
        For i = 0 To m_numOptionPanels - 1
            
            If (i = m_ActivePanel) Then
                
                'Position the panel slightly to the right of the vertical options list, and at the same
                ' top position as said list.
                With m_PanelRect
                    .Left = Me.btsvCategory.GetLeft + Me.btsvCategory.GetWidth + Interface.FixDPI(12)
                    .Top = Me.btsvCategory.GetTop
                    .Width = g_WindowManager.GetClientWidth(Me.hWnd) - .Left
                    .Height = (cmdBarMini.GetTop - .Top) - 1
                End With
                
                'Use the window manager's child panel manager to handle this for us.
                ' (It automatically tracks window bits and restores them when panels are deactivated.)
                g_WindowManager.ActivateToolPanel m_Panels(i).PanelHWnd, Me.hWnd, 0, m_PanelRect.Left, m_PanelRect.Top, m_PanelRect.Width, m_PanelRect.Height
                m_ActivePanelHWnd = m_Panels(i).PanelHWnd
                Exit For
                
            End If
            
        Next i
        
        'Then, forcibly hide all other panels
        For i = 0 To m_numOptionPanels - 1
            If (i <> m_ActivePanel) Then
                If (m_Panels(i).PanelHWnd <> 0) Then g_WindowManager.SetVisibilityByHWnd m_Panels(i).PanelHWnd, False
            End If
        Next i
        
    End If
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Start by auto-validating any controls that accept user input.
    ' (It's up to child forms to implement this independently.)
    Dim validateCheck As Boolean
    validateCheck = True
    
    Dim i As Long
    For i = 0 To MAX_NUM_OPTION_PANELS - 1
        If m_Panels(i).PanelWasLoaded Then
            
            Select Case i
                Case OP_Interface
                    validateCheck = validateCheck And options_Interface.ValidateAllInput()
                Case OP_Menus
                    validateCheck = validateCheck And options_Menus.ValidateAllInput()
                Case OP_Performance
                    validateCheck = validateCheck And options_Performance.ValidateAllInput()
                Case OP_Loading
                    validateCheck = validateCheck And options_Loading.ValidateAllInput()
                Case OP_Saving
                    validateCheck = validateCheck And options_Saving.ValidateAllInput()
                Case OP_Metadata
                    validateCheck = validateCheck And options_Metadata.ValidateAllInput()
                Case OP_ColorManagement
                    validateCheck = validateCheck And options_ColorManagement.ValidateAllInput()
                Case OP_Fonts
                    validateCheck = validateCheck And options_Fonts.ValidateAllInput()
                Case OP_InputDevices
                    validateCheck = validateCheck And options_Input.ValidateAllInput()
                Case OP_Updates
                    validateCheck = validateCheck And options_Updates.ValidateAllInput()
                Case OP_Advanced
                    validateCheck = validateCheck And options_Advanced.ValidateAllInput()
            End Select
            
            If (Not validateCheck) Then
                Me.btsvCategory.ListIndex = i
                Exit For
            End If
            
        End If
    Next i
    
    If (Not validateCheck) Then
        cmdBarMini.DoNotUnloadForm
        Exit Sub
    End If
    
    'If we're still here, all panels (loaded this session) passed user input validation.
    
    Message "Saving preferences..."
    Me.Visible = False
    
    'After updates on 22 Oct 2014, the preference saving sequence should happen in a flash, but just in case,
    ' we'll supply a bit of processing feedback.
    FormMain.Enabled = False
    ProgressBars.SetProgBarMax MAX_NUM_OPTION_PANELS
    ProgressBars.SetProgBarVal 1
    
    'First, make note of the active panel, so we can default to that if the user returns to this dialog
    UserPrefs.SetPref_Long "Core", "Last Preferences Page", btsvCategory.ListIndex
    
    'Write preferences out to file in category order.  (The preference XML file is order-agnostic, but I try to
    ' maintain the order used in the Preferences dialog itself to make changes easier.)
    For i = 0 To MAX_NUM_OPTION_PANELS - 1
        SetProgBarVal i + 1
        If m_Panels(i).PanelWasLoaded Then
            
            Select Case i
                Case OP_Interface
                    options_Interface.SaveUserPreferences
                Case OP_Menus
                    options_Menus.SaveUserPreferences
                Case OP_Performance
                    options_Performance.SaveUserPreferences
                Case OP_Loading
                    options_Loading.SaveUserPreferences
                Case OP_Saving
                    options_Saving.SaveUserPreferences
                Case OP_Metadata
                    options_Metadata.SaveUserPreferences
                Case OP_ColorManagement
                    options_ColorManagement.SaveUserPreferences
                Case OP_Fonts
                    options_Fonts.SaveUserPreferences
                Case OP_InputDevices
                    options_Input.SaveUserPreferences
                Case OP_Updates
                    options_Updates.SaveUserPreferences
                Case OP_Advanced
                    options_Advanced.SaveUserPreferences
            End Select
            
        End If
    Next i
    
    'Forcibly write a copy of the preference data out to file
    UserPrefs.ForceWriteToFile
    
    'All user preferences have now been written out to file
    
    'Because some preferences affect the program's interface, redraw the active image.
    FormMain.Enabled = True
    FormMain.UpdateMainLayout
    FormMain.MainCanvas(0).UpdateAgainstCurrentTheme FormMain.hWnd, True
    If PDImages.IsImageActive Then
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage, FormMain.MainCanvas(0)
        Interface.SyncInterfaceToCurrentImage
    End If
    FormMain.ChangeSessionListenerState UserPrefs.GetPref_Boolean("Loading", "Single Instance", False), True
    
    'TODO: color management changes need to be propagated here; otherwise, they won't trigger until the program is restarted.
    
    SetProgBarVal 0
    ReleaseProgressBar
    
    Message "Preferences updated."
        
End Sub

Private Sub Form_Activate()
    
    'Because the window manager synchronizes visibility state between parent and child window
    ' (as it should), loading a child form while the parent form is still invisible means the
    ' child window doesn't appear.
    
    'So we need to manually show it at first appearance (or if the parent window gets hidden then re-shown).
    If (m_ActivePanelHWnd <> 0) And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetVisibilityByHWnd m_ActivePanelHWnd, True
    
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the preferences file
Private Sub Form_Load()
    
    m_ActivePanel = OP_None
    
    'Prep the category button strip
    With btsvCategory
        
        'Start by adding captions for each button.  This will also update the control's layout to match.
        .AddItem "Interface", OP_Interface
        .AddItem "Menus", OP_Menus
        .AddItem "Performance", OP_Performance
        .AddItem "Loading", OP_Loading
        .AddItem "Saving", OP_Saving
        .AddItem "Metadata", OP_Metadata
        .AddItem "Color management", OP_ColorManagement
        .AddItem "Fonts", OP_Fonts
        .AddItem "Input devices", OP_InputDevices
        .AddItem "Updates", OP_Updates
        .AddItem "Advanced", OP_Advanced
        
        'Next, add images to each button
        Dim prefButtonSize As Long
        prefButtonSize = Interface.FixDPI(22)
        .AssignImageToItem OP_Interface, "pref_interface", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Menus, "menu", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Performance, "pref_performance", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Loading, "pref_loading", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Saving, "pref_saving", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Metadata, "image_metadata", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_ColorManagement, "pref_colormanagement", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Fonts, "font", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_InputDevices, "keyboard", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Updates, "pref_updates", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem OP_Advanced, "pref_advanced", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        
    End With
    
    'Finally, activate the last preferences panel that the user looked at
    Dim activePanel As Long
    activePanel = UserPrefs.GetPref_Long("Core", "Last Preferences Page", 0)
    
    'Failsafe only against bad inputs or future changes
    If (activePanel >= Me.btsvCategory.ListCount) Or (activePanel < 0) Then activePanel = 0
    
    'Set the left-side button strip index to match, and just in case, manually ensure that panel is active
    btsvCategory.ListIndex = activePanel
    UpdateActivePanel activePanel
    
    'Apply translations and visual themes to *this* form (UpdateActivePanel(), above, already handled
    ' that for the child panel form).
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Interface.ReleaseFormTheming Me
    
    'The active panel (if one exists) has had its window bits manually modified so that we can
    ' embed it inside the parent options window.  Make certain those window bits are reset before
    ' we attempt to unload the panel using built-in VB keywords (because VB will crash if it
    ' encounters unexpected window bits, especially WS_CHILD).
    If (Not g_WindowManager Is Nothing) And (m_ActivePanelHWnd <> 0) Then g_WindowManager.DeactivateToolPanel m_ActivePanelHWnd
    m_ActivePanelHWnd = 0
    
    'Make sure our internal panel collection actually exists before attempting to iterate it
    If (m_numOptionPanels = 0) Then Exit Sub
    
    'Free any panels that were loaded this session
    Dim i As PD_OptionPanel
    For i = 0 To MAX_NUM_OPTION_PANELS - 1
        
        'If we loaded this panel during this session, unload it manually now
        If m_Panels(i).PanelWasLoaded Then
            
            Select Case i
                
                Case OP_Interface
                    Unload options_Interface
                    Set options_Interface = Nothing
                    
                Case OP_Menus
                    Unload options_Menus
                    Set options_Menus = Nothing
                    
                Case OP_Performance
                    Unload options_Performance
                    Set options_Performance = Nothing
                    
                Case OP_Loading
                    Unload options_Loading
                    Set options_Loading = Nothing
                    
                Case OP_Saving
                    Unload options_Saving
                    Set options_Saving = Nothing
                
                Case OP_Metadata
                    Unload options_Metadata
                    Set options_Metadata = Nothing
                
                Case OP_ColorManagement
                    Unload options_ColorManagement
                    Set options_ColorManagement = Nothing
                
                Case OP_Fonts
                    Unload options_Fonts
                    Set options_Fonts = Nothing
                
                Case OP_InputDevices
                    Unload options_Input
                    Set options_Input = Nothing
                    
                Case OP_Updates
                    Unload options_Updates
                    Set options_Updates = Nothing
                    
                Case OP_Advanced
                    Unload options_Advanced
                    Set options_Advanced = Nothing
                    
            End Select
            
            m_Panels(i).PanelWasLoaded = False
            
        End If
        
    Next i
    
End Sub

'Called by the _Advanced options panel if the user wants the nuclear "start over" option.
' (This rebuilds a new user preferences file from scratch.  Note that this doesn't reset *everything*
'  related to PD - individual forms still have their individual preset files, for example - but it
'  resets all core app behavior.)
Public Sub ResetAllPreferences()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = PDMsgBox("All settings will be restored to their default values.  This action cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbExclamation Or vbYesNo, "Reset PhotoDemon")
    
    'If the user gives final permission, rewrite the preferences file from scratch and repopulate this form
    If (confirmReset = vbYes) Then
        
        'This resets the preferences file *permanently*.
        UserPrefs.ResetPreferences
        
        'We now need to manually reload the contents of any loaded panels.
        Dim i As Long
        For i = 0 To MAX_NUM_OPTION_PANELS - 1
            
            If m_Panels(i).PanelWasLoaded Then
                
                Select Case i
                    Case OP_Interface
                        options_Interface.LoadUserPreferences
                    Case OP_Menus
                        options_Menus.LoadUserPreferences
                    Case OP_Performance
                        options_Performance.LoadUserPreferences
                    Case OP_Loading
                        options_Loading.LoadUserPreferences
                    Case OP_Saving
                        options_Saving.LoadUserPreferences
                    Case OP_Metadata
                        options_Metadata.LoadUserPreferences
                    Case OP_ColorManagement
                        options_ColorManagement.LoadUserPreferences
                    Case OP_Fonts
                        options_Fonts.LoadUserPreferences
                    Case OP_InputDevices
                        options_Input.LoadUserPreferences
                    Case OP_Updates
                        options_Updates.LoadUserPreferences
                    Case OP_Advanced
                        options_Advanced.LoadUserPreferences
                End Select
                
            End If
            
        Next i
        
        'Restore the currently active language to the preferences file -  this prevents the language
        ' from resetting to English (a behavior that occurs as an accidental byproduct of resetting everything).
        g_Language.WriteLanguagePreferencesToFile
        
    End If
    
End Sub

'For reasons I don't fully understand, raising a modal dialog from one of our child forms
' breaks interactions until we reset window bits.  Child options forms must call this function after a
' modal dialog is raised, which will restore bits and get things working again.
Public Sub RestoreActivePanelBehavior()
    g_WindowManager.ActivateToolPanel m_ActivePanelHWnd, Me.hWnd, m_ActivePanelHWnd, m_PanelRect.Left, m_PanelRect.Top, m_PanelRect.Width, m_PanelRect.Height
End Sub
