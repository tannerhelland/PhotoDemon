VERSION 5.00
Begin VB.Form options_Loading 
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
   Begin PhotoDemon.pdCheckBox chkSplash 
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   3600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   582
      Caption         =   "display splash screen"
   End
   Begin PhotoDemon.pdButtonStrip btsMultiInstance 
      Height          =   975
      Left            =   150
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1720
      Caption         =   "when images arrive from an external source (like Windows Explorer):"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdCheckBox chkToneMapping 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   1920
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "display tone mapping options when importing HDR and RAW images"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   0
      Top             =   1560
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "high-dynamic range (HDR) images"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "app instances"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   5
      Left            =   0
      Top             =   2400
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "session restore"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdCheckBox chkSystemReboots 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   2760
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "automatically restore sessions interrupted by system updates or reboots"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   3240
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "startup"
      FontSize        =   12
      ForeColor       =   5263440
   End
End
Attribute VB_Name = "options_Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Loading panel
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

    btsMultiInstance.AddItem "load into this instance", 0
    btsMultiInstance.AddItem "load into a new PhotoDemon instance", 1
    
End Sub

Public Sub LoadUserPreferences()
    
    If UserPrefs.GetPref_Boolean("Loading", "Single Instance", False) Then btsMultiInstance.ListIndex = 0 Else btsMultiInstance.ListIndex = 1
    chkToneMapping.Value = UserPrefs.GetPref_Boolean("Loading", "Tone Mapping Prompt", True)
    chkSystemReboots.Value = UserPrefs.GetPref_Boolean("Loading", "RestoreAfterReboot", False)
    chkSplash.Value = UserPrefs.GetPref_Boolean("Loading", "splash-screen", True)
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Boolean "Loading", "Single Instance", (btsMultiInstance.ListIndex = 0)
    
    UserPrefs.SetPref_Boolean "Loading", "Tone Mapping Prompt", chkToneMapping.Value
    
    'Restore after reboot behavior requires an immediate API to de/activate
    UserPrefs.SetPref_Boolean "Loading", "RestoreAfterReboot", chkSystemReboots.Value
    OS.SetRestartRestoreBehavior chkSystemReboots.Value
    
    UserPrefs.SetPref_Boolean "Loading", "splash-screen", chkSplash.Value
    
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
    
    chkToneMapping.AssignTooltip "HDR and RAW images contain more colors than PC screens can physically display.  Before displaying such images, a tone mapping operation must be applied to the original image data."
    chkSystemReboots.AssignTooltip "If your PC reboots while PhotoDemon is running, PhotoDemon can automatically restore your previous session."
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub
