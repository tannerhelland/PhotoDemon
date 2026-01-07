VERSION 5.00
Begin VB.Form options_Updates 
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
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   3495
      Left            =   240
      Top             =   3000
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6165
      Caption         =   "(disclaimer populated at run-time)"
      FontSize        =   9
      Layout          =   1
   End
   Begin PhotoDemon.pdDropDown cboUpdates 
      Height          =   735
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "automatically check for updates:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboUpdates 
      Height          =   735
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1350
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "allow updates from these tracks:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      Caption         =   "update options"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdCheckBox chkUpdates 
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   2400
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   582
      Caption         =   "notify when an update is ready"
   End
End
Attribute VB_Name = "options_Updates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Updates panel
'Copyright 2002-2026 by Tanner Helland
'Created: 8/November/02
'Last updated: 02/April/25
'Last update: split this panel into a standalone form
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

    'Update preferences
    cboUpdates(0).Clear
    cboUpdates(0).AddItem "each session", 0
    cboUpdates(0).AddItem "weekly", 1
    cboUpdates(0).AddItem "monthly", 2
    cboUpdates(0).AddItem "never (not recommended)", 3
    cboUpdates(0).AssignTooltip "Because PhotoDemon is a portable application, it can only check for updates when the program is running.  By default, PhotoDemon will check for updates whenever the program is launched, but you can reduce this frequency if desired."
    
    cboUpdates(1).Clear
    cboUpdates(1).AddItem "stable releases", 0
    cboUpdates(1).AddItem "stable and beta releases", 1
    cboUpdates(1).AddItem "stable, beta, and developer releases", 2
    cboUpdates(1).AssignTooltip "One of the best ways to support PhotoDemon is to help test new releases.  By default, PhotoDemon will suggest both stable and beta releases, but the truly adventurous can also try developer releases.  (Developer releases give you immediate access to the latest program enhancements, but you might encounter some bugs.)"
    
    chkUpdates(0).AssignTooltip "PhotoDemon can notify you when it's ready to apply an update.  This allows you to use the updated version immediately."
    
    'In normal (portable) mode, I like to provide a short explanation of how automatic updates work.
    ' In non-portable mode, however, we don't have write access to our own folder (because the user
    ' probably stuck us in an access-restricted folder).  When this happens, we disable all update
    ' options and use the explanation label to explain "why".
    If UserPrefs.IsNonPortableModeActive() Then
    
        'This is a non-portable install.  Disable all update controls, then explain why.
        Dim i As Long
        For i = cboUpdates.lBound() To cboUpdates.UBound()
            cboUpdates(i).Enabled = False
        Next i
        
        For i = chkUpdates.lBound() To chkUpdates.UBound()
            chkUpdates(i).Enabled = False
        Next i
        
        lblExplanation.Caption = g_Language.TranslateMessage("You have placed PhotoDemon in a restricted system folder.  Security precautions prevent PhotoDemon from modifying this folder, so automatic updates are now disabled.  To restore them, you must move PhotoDemon to a non-admin folder, like Desktop, Documents, or Downloads." & vbCrLf & vbCrLf & "(If you leave PhotoDemon where it is, please don't forget to visit photodemon.org from time to time to check for new versions.)")
        
    'This is a normal (portable) install.  Populate the network access disclaimer in the "Update" panel.
    Else
        lblExplanation.Caption = g_Language.TranslateMessage("The developers of PhotoDemon take privacy very seriously, so no information - statistical or otherwise - is uploaded during the update process.  Updates simply involve downloading several small XML files from photodemon.org. These files contain the latest software, plugin, and language version numbers. If updated versions are found, and user preferences allow, the updated files are then downloaded and patched automatically." & vbCrLf & vbCrLf & "If you still choose to disable updates, don't forget to visit photodemon.org from time to time to check for new versions.")
    End If
    
End Sub

Public Sub LoadUserPreferences()

    'Update preferences
    cboUpdates(0).ListIndex = UserPrefs.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
    cboUpdates(1).ListIndex = UserPrefs.GetPref_Long("Updates", "Update Track", ut_Beta)
    chkUpdates(0).Value = UserPrefs.GetPref_Boolean("Updates", "Update Notifications", True)
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Long "Updates", "Update Frequency", cboUpdates(0).ListIndex
    UserPrefs.SetPref_Long "Updates", "Update Track", cboUpdates(1).ListIndex
    UserPrefs.SetPref_Boolean "Updates", "Update Notifications", chkUpdates(0).Value
    
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
