VERSION 5.00
Begin VB.Form options_Menus 
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
   Begin PhotoDemon.pdButtonStrip btsMnemonics 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1720
      Caption         =   "display access keys (mnemonics)"
   End
   Begin PhotoDemon.pdLabel lblRecentFileCount 
      Height          =   240
      Left            =   135
      Top             =   1620
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   423
      Caption         =   "maximum number of recent files to remember: "
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   13
      Left            =   0
      Top             =   1215
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "recent files"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdButtonStrip btsMRUStyle 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2010
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1508
      Caption         =   "recent file menu text:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSpinner tudRecentFiles 
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   1590
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      DefaultValue    =   10
      Min             =   1
      Max             =   32
      Value           =   10
   End
End
Attribute VB_Name = "options_Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Menus panel
'Copyright 2025-2026 by Tanner Helland
'Created: 04/April/25
'Last updated: 04/April/25
'Last update: initial build
'
'This form contains a single subpanel worth of program options.  At run-time, it is dynamically
' made a child of FormOptions.  It will only be loaded if/when the user interacts with this category.
'
'All Tools > Options child panels contain some mandatory public functions, including ones for loading
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

    btsMnemonics.AddItem "auto", 0
    btsMnemonics.AddItem "on", 1
    btsMnemonics.AddItem "off", 2
    
    btsMRUStyle.AddItem "compact (filename only)", 0
    btsMRUStyle.AddItem "verbose (filename and path)", 1
    
End Sub

Public Sub LoadUserPreferences()
    
    Select Case UserPrefs.GetPref_Long("Interface", "display-mnemonics", PD_BOOL_AUTO)
        Case PD_BOOL_FALSE
            btsMnemonics.ListIndex = 2
        Case PD_BOOL_TRUE
            btsMnemonics.ListIndex = 1
        Case Else
            btsMnemonics.ListIndex = 0
    End Select
    
    tudRecentFiles.Value = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
    btsMRUStyle.ListIndex = UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0)
    
End Sub

Public Sub SaveUserPreferences()
    
    'Immediately relay the mnemonics setting to the menu engine
    Select Case btsMnemonics.ListIndex
        Case 0
            UserPrefs.SetPref_Long "Interface", "display-mnemonics", PD_BOOL_AUTO
        Case 1
            UserPrefs.SetPref_Long "Interface", "display-mnemonics", PD_BOOL_TRUE
        Case 2
            UserPrefs.SetPref_Long "Interface", "display-mnemonics", PD_BOOL_FALSE
    End Select
    Menus.SetMnemonicsBehavior UserPrefs.GetPref_Long("Interface", "display-mnemonics", PD_BOOL_AUTO)
    Menus.UpdateAgainstCurrentTheme True
    
    'Changes to the recent files list (including count and how it's displayed) may require us to
    ' trigger a full rebuild of the menu
    Dim mruNeedsToBeRebuilt As Boolean
    mruNeedsToBeRebuilt = (btsMRUStyle.ListIndex <> UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0))
    UserPrefs.SetPref_Long "Interface", "MRU Caption Length", btsMRUStyle.ListIndex

    Dim newMaxRecentFiles As Long
    If tudRecentFiles.IsValid Then newMaxRecentFiles = tudRecentFiles.Value Else newMaxRecentFiles = 10
    If (Not mruNeedsToBeRebuilt) Then mruNeedsToBeRebuilt = (newMaxRecentFiles <> UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10))
    UserPrefs.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value

    'If any MRUs need to be rebuilt, do so now
    If mruNeedsToBeRebuilt Then
        g_RecentFiles.NotifyMaxLimitChanged
        g_RecentMacros.MRU_NotifyNewMaxLimit
    End If

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
    
    btsMnemonics.AssignTooltip "Access keys (mnemonics) allow you to navigate menus using the keyboard.  In ""auto"" mode, languages without native access keys will display the U.S. English access key for each menu."
    
    lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent files to remember: ")
    tudRecentFiles.SetLeft lblRecentFileCount.GetLeft + lblRecentFileCount.GetWidth + Interface.FixDPI(8)
    
    btsMRUStyle.AssignTooltip "The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations."
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub
