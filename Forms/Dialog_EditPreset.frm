VERSION 5.00
Begin VB.Form dialog_AddPreset 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Saved presets"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
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
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   3690
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   2535
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      Begin PhotoDemon.pdTextBox txtName 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   405
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   661
         FontSize        =   11
      End
      Begin PhotoDemon.pdLabel lblName 
         Height          =   375
         Left            =   135
         Top             =   15
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   661
         Caption         =   "enter a name for this preset"
         FontSize        =   11
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   2535
      Index           =   1
      Left            =   120
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      Begin PhotoDemon.pdButton cmdMove 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1085
      End
      Begin PhotoDemon.pdButton cmdDelete 
         Height          =   615
         Left            =   2370
         TabIndex        =   6
         Top             =   1800
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   1085
         Caption         =   "delete preset"
      End
      Begin PhotoDemon.pdListBox lstPresets 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2355
         Caption         =   "saved presets for this tool"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdButton cmdMove 
         Height          =   615
         Index           =   1
         Left            =   1305
         TabIndex        =   4
         Top             =   1800
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1085
      End
   End
End
Attribute VB_Name = "dialog_AddPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Preset Editor Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 06/March/15
'Last updated: 02/September/15
'Last update: convert to the new mini-command-bar UC
'
'PD supports last-used and custom user-entered presets for pretty much every tool in the program.  This was a massive
' undertaking, and it still has a lot of papercuts that drive me nuts (e.g. not being able to delete past presets,
' short of manually editing the XML file yourself - ugh!).
'
'The correct solution is to provide some kind of editor form, where the user can add/rename/delete (sort?) presets at
' their leisure.  This dialog will eventually become that editor, but at present, it's missing a number of features.
' My first goal is just getting "add preset" working, so I can finally convert the command bar over to new PD-specific
' controls (particularly pdButtonToolbox).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user button click from the dialog (OK/Cancel)
Private m_userAnswer As VbMsgBoxResult

'Used to detect initial activation
Private m_IsNotFirstActivation As Boolean

'Because this form needs to interact with both the command bar that raises this dialog, and its preset manager,
' we must maintain references to both.  These references are initially supplied via the showDialog function.
Private m_Presets As pdToolPreset, m_CommandBar As pdCommandBar

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_userAnswer
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByRef srcPresetManager As pdToolPreset, ByRef srcCommandBar As pdCommandBar, ByRef parentForm As Form)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Make sure that a proper cursor is set
    Screen.MousePointer = 0
    
    'Maintain a persistent reference to the source command bar and its preset manager
    Set m_Presets = srcPresetManager
    Set m_CommandBar = srcCommandBar
    
    'Before making any changes to the preset object, back up its current contents
    m_Presets.BackupPresetsInternally
    m_Presets.ClearActivePresetName
    
    'Populate the "existing presets" list, if any exist
    UpdatePresetList
    
    UpdateEditButtons
    
    'Theme the dialog
    ApplyThemeAndTranslations Me
    
    'Display the dialog
    Me.Show vbModal, parentForm
    
End Sub

'Populate the "existing presets" list, if any exist
Private Sub UpdatePresetList()
    
    lstPresets.Clear
    
    Dim tmpStack As pdStringStack, numOfItems As Long, numValidItems As Long
    numOfItems = m_Presets.GetListOfPresets(tmpStack)
    numValidItems = 0
    
    If (numOfItems > 0) Then
        
        Dim i As Long, tmpString As String
        For i = 0 To numOfItems - 1
            tmpString = tmpStack.GetString(i)
            If (LenB(Trim$(tmpString)) > 0) Then
                If Strings.StringsNotEqual(tmpString, "last-used settings", True) And Strings.StringsNotEqual(tmpString, g_Language.TranslateMessage("last-used settings"), True) Then
                    lstPresets.AddItem tmpString
                    numValidItems = numValidItems + 1
                End If
            End If
        Next i
    
    End If
    
    If (numValidItems = 0) Then lstPresets.AddItem "(no saved presets)"
    
End Sub

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    UpdateVisiblePanel
End Sub

Private Sub cmdBarMini_CancelClick()
    
    m_userAnswer = vbCancel
    
    'Undo any changes we may have made to the parent preset object
    m_Presets.RestoreBackedUpPresets
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    'If the user is adding a preset, make sure a valid name was entered.
    If (btsOptions.ListIndex = 0) Then
    
        If (LenB(Trim$(txtName.Text)) <> 0) And (Strings.StringsNotEqual(txtName.Text, g_Language.TranslateMessage("(enter name here)"), True)) Then
            
            'A valid name was entered.  See if this name already exists in the preset manager.
            If m_Presets.DoesPresetExist(Trim$(txtName.Text)) Then
            
                'This name already exists.  Ask the user if an overwrite is okay.
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("A preset with this name already exists.  Do you want to overwrite it?", vbYesNoCancel Or vbExclamation, "Overwrite existing preset")
                
                'Based on the user's answer to the confirmation message box, continue or exit
                Select Case msgReturn
    
                    'If the user selects YES, continue on like normal
                    Case vbYes
                        m_userAnswer = vbOK
                        
                    'If the user selects NO, let them enter a new name
                    Case vbNo
                        txtName.Text = g_Language.TranslateMessage("(enter name here)")
                        txtName.SetFocusToEditBox True
                        cmdBarMini.DoNotUnloadForm
                        Exit Sub
    
                    'If the user selects CANCEL, exit the dialog entirely
                    Case vbCancel
                        m_userAnswer = vbCancel
                    
                End Select
                
            'This preset does not exist, so no special handling is required
            Else
                m_userAnswer = vbOK
            End If
            
            'If the user is okay with us proceeding, update the preset they have just entered.
            ' (Note that this will also update our locally shared m_Presets object.)
            If (m_userAnswer = vbOK) Then m_CommandBar.StorePreset Trim$(txtName.Text)
            
        Else
            
            PDMsgBox "Please enter a name for this preset.", vbInformation Or vbOKOnly, "Preset name required"
            
            txtName.Text = g_Language.TranslateMessage("(enter name here)")
            txtName.SetFocusToEditBox True
            
            cmdBarMini.DoNotUnloadForm
            Exit Sub
            
        End If
    
    'If the user is *not* adding a new preset, there is nothing to validate.
    Else
        m_userAnswer = vbOK
    End If
    
    'Note that this function may have exited prematurely due to the results of a modal dialog.
    If (m_userAnswer = vbOK) Then m_Presets.WritePresetFile
    
End Sub

Private Sub cmdDelete_Click()
    If (lstPresets.ListIndex >= 0) Then
        Dim backupIndex As Long
        backupIndex = lstPresets.ListIndex
        m_Presets.DeletePreset lstPresets.List(lstPresets.ListIndex)
        UpdatePresetList
        If (backupIndex < lstPresets.ListCount) Then
            If Strings.StringsNotEqual(lstPresets.List(backupIndex), g_Language.TranslateMessage("(no saved presets)"), True) Then lstPresets.ListIndex = backupIndex
        Else
            If Strings.StringsNotEqual(lstPresets.List(lstPresets.ListCount - 1), g_Language.TranslateMessage("(no saved presets)"), True) Then lstPresets.ListIndex = lstPresets.ListCount - 1
        End If
        UpdateEditButtons
    End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
    
    If m_Presets.MovePreset(lstPresets.List(lstPresets.ListIndex), (Index = 0)) Then
        
        'Maintain the current listindex
        Dim backupIndex As Long
        backupIndex = lstPresets.ListIndex
        UpdatePresetList
        If (Index = 0) Then lstPresets.ListIndex = backupIndex - 1 Else lstPresets.ListIndex = backupIndex + 1
        
        'Movement may result in the up and/or down buttons being deactivated
        UpdateEditButtons
        
    End If
    
End Sub

Private Sub Form_Activate()
    
    'On initial activation, clear the potentially auto-saved preset name box, then set focus to said box
    If (Not m_IsNotFirstActivation) Then
        
        'Set focus to the text entry box
        txtName.Text = vbNullString
        txtName.SetFocusToEditBox
        
        'Do not display on subsequent activations
        m_IsNotFirstActivation = True
        
    End If
    
End Sub

Private Sub Form_Load()
    
    btsOptions.AddItem "add new preset", 0
    btsOptions.AddItem "edit existing presets", 1
    btsOptions.ListIndex = 0
    UpdateVisiblePanel
    
    Dim dibSize As Long
    dibSize = Interface.FixDPI(38)
    cmdMove(0).AssignImage vbNullString, Interface.GetRuntimeUIDIB(pdri_ArrowUp, dibSize), dibSize, dibSize
    cmdMove(1).AssignImage vbNullString, Interface.GetRuntimeUIDIB(pdri_ArrowDown, dibSize), dibSize, dibSize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release our hold on the parent command bar
    Set m_Presets = Nothing
    Set m_CommandBar = Nothing
    Interface.ReleaseFormTheming Me
    
End Sub

Private Sub lstPresets_Click()
    UpdateEditButtons
End Sub

Private Sub UpdateVisiblePanel()
    Dim i As Long
    For i = 0 To btsOptions.ListCount - 1
        pnlOptions(i).Visible = (i = btsOptions.ListIndex)
    Next i
End Sub

Private Sub UpdateEditButtons()
    cmdDelete.Enabled = (lstPresets.ListIndex >= 0) And Strings.StringsNotEqual(lstPresets.List(lstPresets.ListIndex), g_Language.TranslateMessage("(no saved presets)"), True)
    cmdMove(0).Enabled = (lstPresets.ListIndex > 0)
    cmdMove(1).Enabled = (lstPresets.ListIndex < lstPresets.ListCount - 1) And (lstPresets.ListIndex >= 0)
End Sub
