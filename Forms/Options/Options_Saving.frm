VERSION 5.00
Begin VB.Form options_Saving 
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
   Begin PhotoDemon.pdCheckBox chkConfirmUnsaved 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   5160
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "when closing images, warn about unsaved changes"
   End
   Begin PhotoDemon.pdDropDown cboDefaultSaveFormat 
      Height          =   690
      Left            =   180
      TabIndex        =   1
      Top             =   1680
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      Caption         =   "when using ""Save As"", set the default file format to:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboSaveBehavior 
      Height          =   690
      Left            =   180
      TabIndex        =   2
      Top             =   3000
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      Caption         =   "when ""Save"" is used:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   6
      Left            =   0
      Top             =   2640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      Caption         =   "safe saving"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   7
      Left            =   0
      Top             =   4800
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "unsaved changes"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   8
      Left            =   0
      Top             =   1320
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   503
      Caption         =   "default format"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   21
      Left            =   0
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   503
      Caption         =   "default folder"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdDropDown cboDefaultSaveFolder 
      Height          =   690
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      Caption         =   "when using ""Save As"", set the initial folder to:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboSaveAsBehavior 
      Height          =   690
      Left            =   180
      TabIndex        =   4
      Top             =   3840
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      Caption         =   "when ""Save as"" is used:"
      FontSizeCaption =   10
   End
End
Attribute VB_Name = "options_Saving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Saving panel
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

    'Saving prefs
    chkConfirmUnsaved.AssignTooltip "By default, PhotoDemon will warn you when you attempt to close an image with unsaved changes."
    
    cboDefaultSaveFolder.Clear
    cboDefaultSaveFolder.AddItem "the current image's folder", 0
    cboDefaultSaveFolder.AddItem "the last-used folder", 1
    cboDefaultSaveFolder.AssignTooltip "Most photo editors default to the current image's folder.  For workflows that involve loading images from one folder but saving to a new folder, use the last-used folder to save time."
    
    cboDefaultSaveFormat.Clear
    cboDefaultSaveFormat.AddItem "the current image's format", 0
    cboDefaultSaveFormat.AddItem "the last-used format", 1
    cboDefaultSaveFormat.AssignTooltip "Most photo editors default to the current image's format.  For workflows that involve loading images in one format (e.g. RAW) but saving to a new format (e.g. JPEG), use the last-used format to save time."
    
    cboSaveBehavior.Clear
    cboSaveBehavior.AddItem "overwrite the current file (standard behavior)", 0
    cboSaveBehavior.AddItem "save a new copy, e.g. ""filename (2).jpg"" (safe behavior)", 1
    cboSaveBehavior.AssignTooltip "In most photo editors, the ""Save"" command saves the image over its original version, erasing that copy forever.  PhotoDemon provides a ""safer"" option, where each save results in a new copy of the file."
    
    cboSaveAsBehavior.Clear
    cboSaveAsBehavior.AddItem "suggest the current filename (standard behavior)", 0
    cboSaveAsBehavior.AddItem "suggest a new copy, e.g. ""filename (2).jpg"" (safe behavior)", 1
    cboSaveAsBehavior.AssignTooltip "In most photo editors, the ""Save as"" command defaults to the current filename.  PhotoDemon also provides a ""safer"" option, where Save As will automatically increment filenames for you."
    
End Sub

Public Sub LoadUserPreferences()

    'Saving preferences
    chkConfirmUnsaved.Value = g_ConfirmClosingUnsaved
    cboDefaultSaveFormat.ListIndex = UserPrefs.GetPref_Long("Saving", "Suggested Format", 0)
    If UserPrefs.GetPref_Boolean("Saving", "Use Last Folder", False) Then cboDefaultSaveFolder.ListIndex = 1 Else cboDefaultSaveFolder.ListIndex = 0
    cboSaveBehavior.ListIndex = UserPrefs.GetPref_Long("Saving", "Overwrite Or Copy", 0)
    If UserPrefs.GetPref_Boolean("Saving", "save-as-autoincrement", True) Then cboSaveAsBehavior.ListIndex = 1 Else cboSaveAsBehavior.ListIndex = 0
    
End Sub

Public Sub SaveUserPreferences()

    g_ConfirmClosingUnsaved = chkConfirmUnsaved.Value
    UserPrefs.SetPref_Boolean "Saving", "Confirm Closing Unsaved", g_ConfirmClosingUnsaved
    
    If g_ConfirmClosingUnsaved Then
        toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "If the current image has not been saved, you will receive a prompt to save it before it closes.", "Close the current image"
    Else
        toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.", "Close the current image"
    End If
    
    UserPrefs.SetPref_Long "Saving", "Overwrite Or Copy", cboSaveBehavior.ListIndex
    UserPrefs.SetPref_Long "Saving", "save-as-autoincrement", (cboSaveAsBehavior.ListIndex = 1)
    UserPrefs.SetPref_Long "Saving", "Suggested Format", cboDefaultSaveFormat.ListIndex
    UserPrefs.SetPref_Boolean "Saving", "Use Last Folder", (cboDefaultSaveFolder.ListIndex = 1)
    
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
