VERSION 5.00
Begin VB.Form FormRecordAPNGPrefs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animated screen capture"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565
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
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   6030
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdSlider sldFrameRate 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1296
      Caption         =   "maximum frame rate (fps)"
      Min             =   1
      Max             =   30
      SigDigits       =   1
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
End
Attribute VB_Name = "FormRecordAPNGPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2020 by Tanner Helland
'Created: 01/July/20
'Last updated: 11/July/20
'Last update: get frame optimizations up and running
'
'PD can write animated PNGs.  APNGs seem like a great fit for animated screen captures.
' Let's see if we can merge the two, eh?
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Destination filename, populated when the user hits OK
Private m_Filename As String

'Last-used settings must be stored manually, since we aren't using a dedicated command bar
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub cmdBar_OKClick()
    
    'Before proceeding, we need to prompt the user for a destination filename.
    
    'Start by validating m_Filename, which will be filled with the user's past destination
    ' filename (if one exists), or a default capture filename in the user's current
    ' "Save image" folder
    If ((LenB(m_Filename) = 0) Or (Not Files.PathExists(Files.FileGetPath(m_Filename)))) Then
    
        'm_Filename is bad.  Attempt to populate it with default values.
        Dim tmpPath As String, tmpFilename As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        m_Filename = tmpPath & IncrementFilename(tmpPath, tmpFilename, "apng") & ".apng"
    
    End If
    
    Dim cSave As pdOpenSaveDialog
    Set cSave = New pdOpenSaveDialog
    
    Dim okToProceed As Boolean, sFile As String
    sFile = m_Filename
    okToProceed = cSave.GetSaveFileName(sFile, Files.FileGetName(m_Filename), True, "Animated PNG (.apng)|*.apng;*.png", 1, Files.FileGetPath(m_Filename), "Save image", ".apng", Me.hWnd)
    
    If okToProceed Then
        
        'Save the current export path as the latest "save image" path
        m_Filename = sFile
        UserPrefs.SetPref_String "Paths", "Save Image", Files.FileGetPath(m_Filename)
        
        'Because this dialog is modal, it needs to be hidden before we invoke a modeless dialog
        Me.Hide
        
        'Launch the capture form, then note that the command bar will handle unloading this form
        FormRecordAPNG.ShowDialog m_Filename, sldFrameRate.Value
        
    Else
        cmdBar.DoNotUnloadForm
    End If
    
End Sub

Private Sub Form_Load()
    
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'When the dialog first loads, we want to manually update some of the paths
    ' (like the capture filename).
    
    'The OK button uses custom text
    If (Not g_Language Is Nothing) Then
        cmdBar.SetCustomOKText g_Language.TranslateMessage("Next: select destination filename"), 24
    End If
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_lastUsedSettings.SaveAllControlValues
    m_lastUsedSettings.SetParentForm Nothing
    Interface.ReleaseFormTheming Me
End Sub

Private Sub m_LastUsedSettings_AddCustomPresetData()
    If ((LenB(m_Filename) <> 0) And Files.PathExists(Files.FileGetPath(m_Filename))) Then m_lastUsedSettings.AddPresetData "dst-capture-filename", m_Filename
End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

    'Look for a previously saved destination filename.  If one does not exist,
    ' we want to populate the destination path with a good default suggestion.
    If (Not m_lastUsedSettings.DoesPresetExist("dst-capture-filename")) Then
    
        Dim tmpPath As String, tmpFilename As String, tmpCombined As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        tmpCombined = tmpPath & IncrementFilename(tmpPath, tmpFilename, "apng") & ".apng"
        
        m_Filename = tmpCombined
        
    Else
        m_Filename = m_lastUsedSettings.RetrievePresetData("dst-capture-filename", UserPrefs.GetPref_String("Paths", "Save Image", vbNullString) & "capture.apng")
    End If

End Sub
