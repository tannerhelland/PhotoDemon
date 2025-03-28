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
'Program Preferences Handler
'Copyright 2002-2025 by Tanner Helland
'Created: 8/November/02
'Last updated: 16/January/25
'Last update: new user preference for snap distance (degrees) for angle snapping
'
'Dialog for interfacing with the user's desired program preferences.  Handles reading/writing from/to the persistent
' XML file that actually stores all preferences.
'
'Note that this form interacts heavily with the pdPreferences class.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When the preferences category is changed, only display the controls in that category
Private Sub btsvCategory_Click(ByVal buttonIndex As Long)
    
    'TODO: hide currently active form (if any), load new form
    
    'Dim catID As Long
    'For catID = 0 To btsvCategory.ListCount - 1
    '    picContainer(catID).Visible = (catID = buttonIndex)
    'Next catID
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Start by auto-validating any controls that accept user input
    Dim validateCheck As Boolean
    validateCheck = True
    
    Dim eControl As Object
    For Each eControl In FormOptions.Controls
        
        'Obviously, we can only validate our own custom objects that have built-in auto-validate functions.
        If (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSpinner) Then
            
            'Finally, ask the control to validate itself
            If (Not eControl.IsValid) Then
                validateCheck = False
                Exit For
            End If
            
        End If
    Next eControl
    
    If (Not validateCheck) Then
        cmdBarMini.DoNotUnloadForm
        Exit Sub
    End If
    
    Message "Saving preferences..."
    Me.Visible = False
    
    'After updates on 22 Oct 2014, the preference saving sequence should happen in a flash, but just in case,
    ' we'll supply a bit of processing feedback.
    FormMain.Enabled = False
    ProgressBars.SetProgBarMax 8
    ProgressBars.SetProgBarVal 1
    
    'First, make note of the active panel, so we can default to that if the user returns to this dialog
    UserPrefs.SetPref_Long "Core", "Last Preferences Page", btsvCategory.ListIndex
    
    'Write preferences out to file in category order.  (The preference XML file is order-agnostic, but I try to
    ' maintain the order used in the Preferences dialog itself to make changes easier.)
    
    '***************************************************************************
    
    'Interface preferences
    SetProgBarVal 1
    
    'Loading preferences
    SetProgBarVal 2
    
    'Saving preferences
    SetProgBarVal 3
    
    'Performance preferences.  (Note that many of these are specially cached, for obvious perf reasons.)
    SetProgBarVal 4
    
    'Color-management preferences
    SetProgBarVal 5
    
    'Update preferences
    SetProgBarVal 6
    
    'Advanced preferences
    SetProgBarVal 7
    
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

'RESET will regenerate the preferences file from scratch.  This can be an effective way to
' "reset" a copy of the program.
Private Sub cmdReset_Click()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = PDMsgBox("All settings will be restored to their default values.  This action cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbExclamation Or vbYesNo, "Reset PhotoDemon")

    'If the user gives final permission, rewrite the preferences file from scratch and repopulate this form
    If (confirmReset = vbYes) Then
    
        UserPrefs.ResetPreferences
        LoadAllPreferences
        
        'Restore the currently active language to the preferences file; this prevents the language from resetting to English
        ' (a behavior that isn't made clear by this action).
        g_Language.WriteLanguagePreferencesToFile
        
    End If

End Sub

'Load all relevant values from the user's preferences file, and populate corresponding UI elements
' with those settings
Private Sub LoadAllPreferences()
    
    'Preferences can be loaded in any order (without consequence), but due to the size of PD's
    ' settings list, I try to keep them ordered by category.
    
    
    'TODO
    
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the preferences file
Private Sub Form_Load()
    
    Dim i As Long
    
    'Prep the category button strip
    With btsvCategory
        
        'Start by adding captions for each button.  This will also update the control's layout to match.
        .AddItem "Interface", 0
        .AddItem "Loading", 1
        .AddItem "Saving", 2
        .AddItem "Performance", 3
        .AddItem "Color management", 4
        .AddItem "Updates", 5
        .AddItem "Advanced", 6
        
        'Next, add images to each button
        Dim prefButtonSize As Long
        prefButtonSize = Interface.FixDPI(32)
        .AssignImageToItem 0, "pref_interface", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 1, "pref_loading", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 2, "pref_saving", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 3, "pref_performance", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 4, "pref_colormanagement", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 5, "pref_updates", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        .AssignImageToItem 6, "pref_advanced", Nothing, prefButtonSize, prefButtonSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_CatmullRom, rf_Automatic)
        
    End With
    
    'Hide all category panels (the proper one will be activated after prefs are loaded)
    'For i = 0 To picContainer.Count - 1
    '    picContainer(i).Visible = False
    'Next i
    
    'With all controls initialized, we can now assign them their corresponding values from the preferences file
    If PDMain.IsProgramRunning() Then LoadAllPreferences
    
    'Finally, activate the last preferences panel that the user looked at
    Dim activePanel As Long
    activePanel = UserPrefs.GetPref_Long("Core", "Last Preferences Page", 0)
    'If (activePanel > picContainer.UBound) Then activePanel = picContainer.UBound
    'picContainer(activePanel).Visible = True
    btsvCategory.ListIndex = activePanel
    
    'Apply translations and visual themes
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
