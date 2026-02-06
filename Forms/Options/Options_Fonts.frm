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
   Begin PhotoDemon.pdButton cmdUserFonts 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Caption         =   "add another folder"
   End
   Begin PhotoDemon.pdListBox lstFonts 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2990
      Caption         =   "font folders:"
   End
   Begin PhotoDemon.pdLabel lblInfo 
      Height          =   615
      Index           =   0
      Left            =   120
      Top             =   3600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1085
      Alignment       =   2
      Caption         =   ""
      Layout          =   1
   End
   Begin PhotoDemon.pdDropDownFont ddFont 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1508
      Caption         =   "interface font"
   End
   Begin PhotoDemon.pdButton cmdUserFonts 
      Height          =   495
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Caption         =   "remove this folder"
   End
End
Attribute VB_Name = "options_Fonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Fonts panel
'Copyright 2025-2026 by Tanner Helland
'Created: 04/April/25
'Last updated: 07/April/25
'Last update: UI for users to add/remove custom font folders
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

Private Const FONT_PRESETS_FILE As String = "font_folders.txt"

Private Sub cmdUserFonts_Click(Index As Integer)
    
    'Add a new font folder
    If (Index = 0) Then
        
        'Default to the PD font folder, unless the user has already added another folder to the list
        ' (and selected it).
        Dim initFolder As String
        initFolder = UserPrefs.GetFontPath()
        If (lstFonts.ListIndex > 0) Then initFolder = lstFonts.List(lstFonts.ListIndex, False)
        
        Dim newFolder As String
        newFolder = Files.PathBrowseDialog(Me.hWnd, initFolder)
        If (LenB(newFolder) <> 0) Then
            If Files.PathExists(newFolder, False) Then
                lstFonts.AddItem newFolder
                lstFonts.ListIndex = lstFonts.ListCount - 1
            End If
        End If
        
    'Remove the selected font folder
    ElseIf (Index = 1) Then
        
        'Failsafe only.  (User is not allowed to delete index = 0, which is PD's default user font folder.)
        If (lstFonts.ListIndex > 0) Then
            lstFonts.RemoveItem lstFonts.ListIndex
            UpdateFontButtons
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    
    'Populate the font UI
    ddFont.InitializeFontList
    UpdateFontButtons
    
End Sub

Private Sub UpdateFontButtons()
    cmdUserFonts(0).Enabled = True
    cmdUserFonts(1).Enabled = (lstFonts.ListIndex > 0)
End Sub

Public Sub LoadUserPreferences()
    
    'The default PD font folder is always available, but the user can add more.
    lstFonts.Clear
    lstFonts.AddItem UserPrefs.GetFontPath(), 0, True
    
    'Load any/all custom font folders saved in previous sessions
    Dim srcFile As String
    srcFile = UserPrefs.GetPresetPath & FONT_PRESETS_FILE
    If Files.FileExists(srcFile) Then
        
        'Load preset file
        Dim srcList As String
        If Files.FileLoadAsString(srcFile, srcList, True) Then
            
            'Iterate lines in the file
            Dim cStack As pdStringStack: Set cStack = New pdStringStack
            If cStack.CreateFromMultilineString(srcList, vbCrLf) Then
                
                Dim i As Long
                For i = 0 To cStack.GetNumOfStrings - 1
                    srcFile = cStack.GetString(i)
                    If (LenB(srcFile) > 0) Then
                        If Files.PathExists(srcFile, False) Then lstFonts.AddItem srcFile
                    End If
                Next i
                
            End If
            
        End If
        
    End If
    
    'Technically, the active font name comes from the font engine, *not* user prefs.
    ' Start there, but if there's a difference, default to the one in the user prefs file.
    ' (This would mean the user has changed the font this session, but *not* restarted the app.)
    Dim curFontName As String, curFontPref As String
    curFontName = Fonts.GetUIFontName()
    curFontPref = UserPrefs.GetUIFontName()
    
    Dim targetName As String
    If Strings.StringsEqual(curFontName, curFontPref, True) Then
        targetName = curFontName
    Else
        
        'The user prefs value will be NULL until the user interacts with it
        If (LenB(curFontPref) > 0) Then
            
            'If the font doesn't exist on this PC, revert to PD's default UI font
            Dim cFont As pdFont: Set cFont = New pdFont
            If cFont.DoesFontExist(curFontPref) Then
                targetName = curFontPref
            Else
                targetName = curFontName
            End If
        
        'Use PD's default font
        Else
            targetName = curFontName
        End If
    End If
    
    'Default to the most appropriate font
    ddFont.ListIndex = ddFont.ListIndexByString(targetName, vbTextCompare)
    
End Sub

Public Sub SaveUserPreferences()
    
    'Settings belong in the central settings file...
    UserPrefs.SetPref_String "Interface", "UIFont", ddFont.List(ddFont.ListIndex, False)
    
    'Custom font folders are stored in a simple text file
    Dim dstFile As String
    dstFile = UserPrefs.GetPresetPath & FONT_PRESETS_FILE
    
    'Only create said list if the user has actually added font folders
    If (lstFonts.ListCount > 1) Then
        
        'Open a persistent text file
        Files.FileDeleteIfExists dstFile
        
        'Add all strings to a single string, then dump the whole thing to file.
        Dim cString As pdString
        Set cString = New pdString
        
        Dim i As Long
        For i = 1 To lstFonts.ListCount - 1
            cString.AppendLine lstFonts.List(i, False)
        Next i
        
        Files.FileSaveAsText cString.ToString(), dstFile
    
    'If the user previously added font files, then deleted them, clear the preset file
    Else
        Files.FileDeleteIfExists dstFile
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
    
    lblInfo(0).Caption = g_Language.TranslateMessage("Changes will take effect the next time you start PhotoDemon.")
    Interface.ApplyThemeAndTranslations Me
    
    UpdateFontButtons
    
End Sub

Private Sub lstFonts_Click()
    UpdateFontButtons
End Sub
