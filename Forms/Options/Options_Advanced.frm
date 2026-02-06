VERSION 5.00
Begin VB.Form options_Advanced 
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
   Begin PhotoDemon.pdButton cmdReset 
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   3465
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1058
      Caption         =   "reset all program settings"
   End
   Begin PhotoDemon.pdButton cmdTmpPath 
      Height          =   450
      Left            =   7680
      TabIndex        =   1
      Top             =   4575
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   794
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtTempPath 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   4650
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      Text            =   "automatically generated at run-time"
   End
   Begin PhotoDemon.pdLabel lblMemoryUsageMax 
      Height          =   345
      Left            =   240
      Top             =   2655
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   609
      Caption         =   "memory usage will be displayed here"
      ForeColor       =   8405056
   End
   Begin PhotoDemon.pdLabel lblMemoryUsageCurrent 
      Height          =   345
      Left            =   240
      Top             =   2295
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   609
      Caption         =   "memory usage will be displayed here"
      ForeColor       =   8405056
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   5
      Left            =   0
      Top             =   1935
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   503
      Caption         =   "memory diagnostics"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   19
      Left            =   0
      Top             =   4200
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "temporary file location"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTempPathWarning 
      Height          =   480
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   847
      ForeColor       =   255
      Layout          =   1
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   0
      Top             =   3105
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   503
      Caption         =   "start over"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   20
      Left            =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "application settings folder"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblSettingsFolder 
      Height          =   285
      Left            =   240
      Top             =   360
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   503
      Caption         =   ""
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdButtonStrip btsDebug 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1720
      Caption         =   "generate debug logs"
   End
End
Attribute VB_Name = "options_Advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Advanced panel
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

'Resetting preferences involves some circular bullshit, so the magic happens back in FormOptions
' (which is dynamically made the parent of this window at run-time).
Private Sub cmdReset_Click()
    FormOptions.ResetAllPreferences
End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub cmdTmpPath_Click()
    Dim tString As String
    tString = Files.PathBrowseDialog(Me.hWnd, UserPrefs.GetTempPath)
    If (LenB(tString) <> 0) Then txtTempPath.Text = Files.PathAddBackslash(tString)
End Sub

'If the selected temp folder doesn't have write access, warn the user
Private Sub txtTempPath_Change()
    
    'Assign theme-specific error colors
    If Me.Visible Then
        lblTempPathWarning.ForeColor = g_Themer.GetGenericUIColor(UI_ErrorRed)
        lblTempPathWarning.UseCustomForeColor = True
    End If
    
    'Ensure the specified temp path exists.  If it doesn't (or if access is denied), let the user know that we will silently
    ' fall back to the system temp folder.
    If (Not Files.PathExists(Trim$(txtTempPath.Text), True)) Then
        lblTempPathWarning.Caption = g_Language.TranslateMessage("WARNING: this folder is invalid (access prohibited).  Please provide a valid folder.  If a new folder is not provided, PhotoDemon will use the system temp folder.")
        lblTempPathWarning.Visible = True
    Else
        lblTempPathWarning.Caption = g_Language.TranslateMessage("This new temporary folder location will not take effect until you restart the program.")
        lblTempPathWarning.Visible = Strings.StringsNotEqual(Trim$(txtTempPath.Text), UserPrefs.GetTempPath, True)
    End If
    
End Sub

Private Sub Form_Load()

    'Advanced preferences
    lblSettingsFolder.Caption = UserPrefs.GetDataPath()
    
    btsDebug.AddItem "auto", 0
    btsDebug.AddItem "no", 1
    btsDebug.AddItem "yes", 2
    btsDebug.AssignTooltip "In developer builds, debug data is automatically logged to the program's \Data\Debug folder.  If you encounter bugs in a stable release, please manually activate this setting.  This will help developers resolve your problem."
    
    lblMemoryUsageCurrent.Caption = g_Language.TranslateMessage("current PhotoDemon memory usage:") & " " & Format$(OS.AppMemoryUsageInMB(), "#,#") & " M"
    lblMemoryUsageMax.Caption = g_Language.TranslateMessage("max PhotoDemon memory usage this session:") & " " & Format$(OS.AppMemoryUsageInMB(True), "#,#") & " M"
    
    cmdTmpPath.AssignTooltip "Click to select a new temporary folder."
    cmdReset.AssignTooltip "This button resets all PhotoDemon settings.  If the program is behaving unexpectedly, this may resolve the problem."
    
End Sub

Public Sub LoadUserPreferences()

    'Advanced preferences
    lblSettingsFolder.Caption = UserPrefs.GetDataPath()
    btsDebug.ListIndex = UserPrefs.GetPref_Long("Core", "GenerateDebugLogs", 0)
    txtTempPath.Text = UserPrefs.GetTempPath
    
End Sub

Public Sub SaveUserPreferences()

    'First, see if the user has changed the debug log preference
    If (UserPrefs.GetDebugLogPreference <> btsDebug.ListIndex) Then
        
        'The user has changed the current setting.  Make a note of whether debug logs are currently being generated.
        ' (If this behavior changes, we may need to create and/or terminate the debugger.)
        Dim curLogBehavior As Boolean
        curLogBehavior = UserPrefs.GenerateDebugLogs()
        
        'Store the new preference
        UserPrefs.SetDebugLogPreference btsDebug.ListIndex
        
        'Invoke and/or terminate the current debugger, as necessary
        If (curLogBehavior <> UserPrefs.GenerateDebugLogs()) Then
            If UserPrefs.GenerateDebugLogs Then PDDebug.StartDebugger True, , False Else PDDebug.TerminateDebugger False
        End If
        
    End If
    
    If Strings.StringsNotEqual(Trim$(txtTempPath), UserPrefs.GetTempPath, True) Then UserPrefs.SetTempPath Trim$(txtTempPath)
    
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
