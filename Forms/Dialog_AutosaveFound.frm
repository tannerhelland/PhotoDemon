VERSION 5.00
Begin VB.Form dialog_AutosaveWarning 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Autosave data detected"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9165
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
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBox picWarning 
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdListBox lstAutosaves 
      Height          =   2730
      Left            =   240
      TabIndex        =   1
      Top             =   1830
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6085
   End
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   4740
      Width           =   3540
      _ExtentX        =   5821
      _ExtentY        =   1296
      Caption         =   "Attempt to recover"
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   2730
      Left            =   3960
      Top             =   1830
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   4815
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   735
      Left            =   5400
      TabIndex        =   2
      Top             =   4740
      Width           =   3540
      _ExtentX        =   5821
      _ExtentY        =   1296
      Caption         =   "Discard"
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   645
      Index           =   1
      Left            =   240
      Top             =   960
      Width           =   8745
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   525
      Index           =   0
      Left            =   1005
      Top             =   330
      Width           =   8055
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "Autosave data found.  Would you like to restore it?"
      FontSize        =   12
      ForeColor       =   2105376
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_AutosaveWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Autosave (unsafe shutdown) Prompt/Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 19/January/14
'Last updated: 10/January/17
'Last update: implement better theming support
'
'PhotoDemon now provides AutoSave functionality.  If the program terminates unexpectedly, this dialog will be raised,
' which gives the user an option to restore any in-progress image edits.
'
'Images that had been loaded by PhotoDemon but never modified will not be shown.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'Collection of Autosave XML entries found
Private m_numOfXMLFound As Long
Private m_XmlEntries() As AutosaveXML

'Theme-specific icons are fully supported
Private m_warningDIB As pdDIB

'When this dialog finally closes, the calling function can use this sub to retrieve the entries the user wants saved.
Friend Sub FillArrayWithSaveResults(ByRef dstArray() As AutosaveXML)
    
    ReDim dstArray(0 To m_numOfXMLFound - 1) As AutosaveXML
    
    Dim i As Long
    For i = 0 To m_numOfXMLFound - 1
        dstArray(i) = m_XmlEntries(i)
    Next i
    
End Sub

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog()
    
    'Prep a warning icon
    Dim warningIconSize As Long
    warningIconSize = Interface.FixDPI(32)
    
    If Not IconsAndCursors.LoadResourceToDIB("generic_warning", m_warningDIB, warningIconSize, warningIconSize, 0) Then
        Set m_warningDIB = Nothing
        picWarning.Visible = False
    End If
    
    'Display a brief explanation of the dialog at the top of the window
    lblWarning(1).Caption = g_Language.TranslateMessage("A previous PhotoDemon session terminated unexpectedly.  Would you like to automatically recover the following autosaved images?")
    
    'Provide a default answer of "do not restore" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbNo
    
    'Load command button images
    Dim buttonIconSize As Long
    buttonIconSize = Interface.FixDPI(32)
    cmdOK.AssignImage "generic_ok", , buttonIconSize, buttonIconSize
    cmdCancel.AssignImage "generic_cancel", , buttonIconSize, buttonIconSize
    
    'Apply any custom styles to the form
    ApplyThemeAndTranslations Me

    'Populate the AutoSave entry list box
    DisplayAutosaveEntries

    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

'If the user cancels, warn them that these image will be lost foreeeever.
Private Sub cmdCancel_Click()

    Dim msgReturn As VbMsgBoxResult
    msgReturn = PDMsgBox("If you exit now, this autosave data will be lost forever.  Are you sure you want to exit?", vbExclamation Or vbYesNo, "Warning: autosave data will be deleted")
    
    If (msgReturn = vbYes) Then
        userAnswer = vbNo
        Me.Hide
    End If

End Sub

Private Sub CmdOK_Click()
    userAnswer = vbYes
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update the active image preview in the top-right
Private Sub UpdatePreview(ByVal srcImagePath As String)
    
    'Display a preview of the selected image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    If tmpDIB.CreateFromFile(srcImagePath) Then
        picPreview.CopyDIB tmpDIB, , True, , True, , True
    Else
        picPreview.PaintText g_Language.TranslateMessage("preview not available")
    End If
    
End Sub

'Fill the AutoSave entries list with any images found from the Autosave engine
Private Function DisplayAutosaveEntries() As Boolean

    'Because we've arrived at this point, we know that the Autosave engine has found at least *some* usable image data.
    ' Our goal now is to present that image data to the user, so they can select which images (if any) they want us
    ' to restore.
    
    'The Autosaves module will already contain a list of all Undo XML files found by the Autosave engine.
    ' It has stored this data in its private m_XmlEntries() array.  We can request a copy of this array as follows:
    Autosaves.GetXMLAutosaveEntries m_XmlEntries(), m_numOfXMLFound
    
    'All XML entries will now have been matched up with their latest Undo entry.  Fill the listbox with their data,
    ' ignoring any entries that do not have binary image data attached.
    lstAutosaves.SetAutomaticRedraws False
    lstAutosaves.Clear
    
    Dim i As Long
    For i = 0 To m_numOfXMLFound - 1
        lstAutosaves.AddItem m_XmlEntries(i).friendlyName
    Next i
    
    'Select the entry at the top of the list by default
    lstAutosaves.ListIndex = 0
    lstAutosaves.SetAutomaticRedraws True, True
    
End Function

Private Sub lstAutosaves_Click()
    
    'PD always saves a thumbnail of the latest image state to the same Undo path as the XML file, but with
    ' the "pdasi" extension (which represents "PD autosave image").
    Dim previewPath As String
    previewPath = m_XmlEntries(lstAutosaves.ListIndex).xmlPath & ".pdasi"
    UpdatePreview previewPath
    
End Sub

Private Sub picWarning_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_warningDIB Is Nothing) Then m_warningDIB.AlphaBlendToDC targetDC, , (ctlWidth - m_warningDIB.GetDIBWidth) \ 2, (ctlHeight - m_warningDIB.GetDIBHeight) \ 2
End Sub
