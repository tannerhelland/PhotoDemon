VERSION 5.00
Begin VB.Form options_Input 
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
   Begin PhotoDemon.pdButtonStrip btsMouseWheel 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1720
      Caption         =   "mouse wheel behavior"
   End
   Begin PhotoDemon.pdSpinner spnSnapDistance 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      DefaultValue    =   8
      Min             =   1
      Max             =   255
      Value           =   8
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   23
      Left            =   0
      Top             =   2400
      Width           =   4020
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "snap distance (in pixels)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   25
      Left            =   4080
      Top             =   2400
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   503
      Caption         =   "angle snap distance (in degrees)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdSpinner spnSnapDistance 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      DefaultValue    =   5
      Min             =   1
      Max             =   15
      SigDigits       =   1
      Value           =   5
   End
   Begin PhotoDemon.pdButtonStrip btsMouseHighRes 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1720
      Caption         =   "high-resolution mouse input"
   End
End
Attribute VB_Name = "options_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Input panel
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
    
    'High-res mouse input only needs to be deactivated if there are obvious glitches.  This is a Windows-level
    ' problem that seems to show up on VMs and Remote Desktop (see https://forums.getpaint.net/topic/28852-line-jumpsskips-to-top-of-window-while-drawing/)
    btsMouseHighRes.AddItem "off", 0
    btsMouseHighRes.AddItem "on", 1
    
    btsMouseWheel.AddItem "scroll", 0
    btsMouseWheel.AddItem "zoom", 1
    
End Sub

Public Sub LoadUserPreferences()
    
    If UserPrefs.GetPref_Boolean("Tools", "HighResMouseInput", True) Then btsMouseHighRes.ListIndex = 1 Else btsMouseHighRes.ListIndex = 0
    
    If UserPrefs.GetPref_Boolean("Interface", "wheel-zoom", False) Then
        btsMouseWheel.ListIndex = 1
    Else
        btsMouseWheel.ListIndex = 0
    End If
    
    spnSnapDistance(0).Value = UserPrefs.GetPref_Long("Interface", "snap-distance", 8&)
    spnSnapDistance(1).Value = UserPrefs.GetPref_Float("Interface", "snap-degrees", 7.5)
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Boolean "Interface", "wheel-zoom", (btsMouseWheel.ListIndex = 1)
    UserPrefs.SetZoomWithWheel (btsMouseWheel.ListIndex = 1)

    UserPrefs.SetPref_Long "Interface", "snap-distance", spnSnapDistance(0).Value
    Snap.SetSnap_Distance spnSnapDistance(0).Value
    
    UserPrefs.SetPref_Long "Interface", "snap-degrees", spnSnapDistance(1).Value
    Snap.SetSnap_Degrees spnSnapDistance(1).Value

    If (btsMouseHighRes.ListIndex = 1) Then UserPrefs.SetPref_Boolean "Tools", "HighResMouseInput", True Else UserPrefs.SetPref_Boolean "Tools", "HighResMouseInput", False
    Tools.SetToolSetting_HighResMouse (btsMouseHighRes.ListIndex = 1)
    
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
    
    btsMouseHighRes.AssignTooltip "When using Remote Desktop or a VM (Virtual Machine), high-resolution mouse input may not work correctly.  This is a long-standing Windows bug.  In these situations, you can use this setting to restore correct mouse behavior."
    
End Sub
