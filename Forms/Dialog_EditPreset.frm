VERSION 5.00
Begin VB.Form dialog_AddPreset 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add new preset"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1323
      BackColor       =   14802140
      dontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdTextBox txtName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      FontSize        =   11
   End
   Begin PhotoDemon.pdLabel lblName 
      Height          =   375
      Left            =   120
      Top             =   150
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   661
      Caption         =   "enter a name for this preset"
      FontSize        =   11
   End
End
Attribute VB_Name = "dialog_AddPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Preset Editor Dialog
'Copyright 2014-2015 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user button click from the dialog (OK/Cancel)
Private userAnswer As VbMsgBoxResult

'The new preset name entered in the text box
Private presetName As String

'Because this form needs to interact with a preset manager (provided by the command bar that raises the dialog),
' we must maintain a reference to it.  This reference is initially supplied via the showDialog function.
Private m_Presets As pdToolPreset

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Public Property Get newPresetName() As String
    newPresetName = presetName
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog(ByRef srcPresetManager As pdToolPreset, ByRef parentForm As Form)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that a proper cursor is set
    Screen.MousePointer = 0
    
    'Maintain a persistent reference to the source command bar
    Set m_Presets = srcPresetManager
    
    'Theme the dialog
    makeFormPretty Me
    
    'Display the dialog
    Me.Show vbModal, parentForm
    
End Sub

Private Sub cmdBarMini_CancelClick()
    userAnswer = vbCancel
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Make sure a valid name was entered.
    If Len(Trim$(txtName.Text)) <> 0 Then
        
        'A valid name was entered.  See if this name already exists in the preset manager.
        If m_Presets.doesPresetExist(Trim$(txtName.Text)) Then
        
            'This name already exists.  Ask the user if an overwrite is okay.
            Dim msgReturn As VbMsgBoxResult
            msgReturn = pdMsgBox("A preset with this name already exists.  Do you want to overwrite it?", vbYesNoCancel + vbApplicationModal + vbInformation, "Overwrite existing preset")
            
            'Based on the user's answer to the confirmation message box, continue or exit
            Select Case msgReturn

                'If the user selects YES, continue on like normal
                Case vbYes
                    userAnswer = vbOK

                'If the user selects NO, let them enter a new name
                Case vbNo
                    txtName.Text = g_Language.TranslateMessage("(enter name here)")
                    txtName.SetFocus
                    txtName.selectAll
                    cmdBarMini.doNotUnloadForm
                    Exit Sub

                'If the user selects CANCEL, exit the dialog entirely
                Case vbCancel
                    userAnswer = vbCancel
                
            End Select
            
        'This preset does not exist, so no special handling is required
        Else
            userAnswer = vbOK
        End If
        
    Else
        
        pdMsgBox "Please enter a name for this preset.", vbInformation + vbOKOnly + vbApplicationModal, "Preset name required"
        
        txtName.Text = g_Language.TranslateMessage("(enter name here)")
        txtName.SetFocus
        txtName.selectAll
        
        cmdBarMini.doNotUnloadForm
        Exit Sub
        
    End If
    
    'Store the preset name, if any, before exiting
    presetName = Trim$(txtName.Text)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release our hold on the parent command bar
    Set m_Presets = Nothing
    
End Sub
