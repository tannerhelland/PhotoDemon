VERSION 5.00
Begin VB.Form options_Performance 
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
   Begin PhotoDemon.pdSlider sltUndoCompression 
      Height          =   765
      Left            =   180
      TabIndex        =   0
      Top             =   4170
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   873
      Caption         =   "compress undo/redo data at the following level:"
      FontSizeCaption =   10
      Max             =   9
      SliderTrackStyle=   1
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdDropDown cboPerformance 
      Height          =   690
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   1217
      Caption         =   "when decorating interface elements:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboPerformance 
      Height          =   690
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   1217
      Caption         =   "when generating image and layer thumbnail images:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboPerformance 
      Height          =   690
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   2850
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   1217
      Caption         =   "when rendering the image canvas:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   16
      Left            =   0
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   503
      Caption         =   "interface"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblPNGCompression 
      Height          =   240
      Index           =   3
      Left            =   300
      Top             =   5040
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   503
      Caption         =   "no compression (fastest)"
      FontSize        =   8
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblPNGCompression 
      Height          =   240
      Index           =   2
      Left            =   3960
      Top             =   5040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "maximum compression (slowest)"
      FontSize        =   8
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   18
      Left            =   0
      Top             =   3780
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      Caption         =   "undo/redo"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   15
      Left            =   0
      Top             =   1260
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   503
      Caption         =   "thumbnails"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   17
      Left            =   0
      Top             =   2490
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "viewport"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "options_Performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Performance panel
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

    'Perf prefs
    Dim i As Long
    For i = 0 To cboPerformance.UBound
        cboPerformance(i).Clear
        cboPerformance(i).AddItem "maximize quality", 0
        cboPerformance(i).AddItem "balance performance and quality", 1
        cboPerformance(i).AddItem "maximize performance", 2
    Next i
        
    cboPerformance(0).AssignTooltip "Some interface elements receive custom decorations (like drop shadows).  On older PCs, these decorations can be suspended for a small performance boost."
    cboPerformance(1).AssignTooltip "PhotoDemon generates many thumbnail images, especially when images contain multiple layers.  Thumbnail quality can be lowered to improve performance."
    cboPerformance(2).AssignTooltip "Rendering the primary image canvas is a common bottleneck for PhotoDemon's performance.  The automatic setting is recommended, but for older PCs, you can manually select the Maximize Performance option to sacrifice quality for raw performance."
    sltUndoCompression.AssignTooltip "Low compression settings require more disk space, but undo/redo operations will be faster.  High compression settings require less disk space, but undo/redo operations will be slower.  Undo data is erased when images are closed, so this setting only affects disk space while images are actively being edited."
    
End Sub

Public Sub LoadUserPreferences()

    'Performance preferences
    cboPerformance(0).ListIndex = g_InterfacePerformance
    cboPerformance(1).ListIndex = UserPrefs.GetThumbnailPerformancePref()
    cboPerformance(2).ListIndex = g_ViewportPerformance
    sltUndoCompression.Value = g_UndoCompressionLevel
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Long "Performance", "Interface Decoration Performance", cboPerformance(0).ListIndex
    g_InterfacePerformance = cboPerformance(0).ListIndex
    
    UserPrefs.SetPref_Long "Performance", "Thumbnail Performance", cboPerformance(1).ListIndex
    UserPrefs.SetThumbnailPerformancePref cboPerformance(1).ListIndex
    
    UserPrefs.SetPref_Long "Performance", "Viewport Render Performance", cboPerformance(2).ListIndex
    g_ViewportPerformance = cboPerformance(2).ListIndex
    
    UserPrefs.SetPref_Long "Performance", "Undo Compression", sltUndoCompression.Value
    g_UndoCompressionLevel = sltUndoCompression.Value
    
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
