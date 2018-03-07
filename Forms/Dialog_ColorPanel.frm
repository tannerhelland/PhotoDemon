VERSION 5.00
Begin VB.Form dialog_ColorPanel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Color panel settings"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9045
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
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsStyle 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1931
      Caption         =   "style"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2790
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
      Begin PhotoDemon.pdButton cmdPaletteChoose 
         Height          =   375
         Left            =   8280
         TabIndex        =   5
         Top             =   330
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "..."
      End
      Begin PhotoDemon.pdTextBox txtPaletteFile 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   450
         Caption         =   "palette to use:"
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
   End
End
Attribute VB_Name = "dialog_ColorPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color panel settings dialog
'Copyright 2018-2018 by Tanner Helland
'Created: 13/February/18
'Last updated: 13/February/18
'Last update: initial build
'
'The right-side color panel now supports multiple color selection modes.  Hopefully this gives creators
' increased freedom when deciding how they want to paint an image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog.  If the user cancels this dialog, default settings will be used.
Private m_CmdBarAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_CmdBarAnswer
End Property

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub ShowDialog()
    
    'Provide a default answer (in case the user closes the dialog via some means other than the command bar)
    m_CmdBarAnswer = vbCancel
    
    'Prep any dynamic UI objects
    btsStyle.AddItem "wheels + history", 0
    btsStyle.AddItem "palette", 1
    btsStyle.ListIndex = UserPrefs.GetPref_Long("Tools", "ColorPanelStyle", 0)
    txtPaletteFile.Text = UserPrefs.GetPref_String("Tools", "ColorPanelPaletteFile")
    UpdateVisiblePanel
    
    'Apply any custom styles to the form
    ApplyThemeAndTranslations Me

    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    UpdateVisiblePanel
End Sub

Private Sub cmdBar_CancelClick()
    m_CmdBarAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    'Write these preferences to file before exiting
    UserPrefs.SetPref_Long "Tools", "ColorPanelStyle", btsStyle.ListIndex
    If (btsStyle.ListIndex = 1) Then UserPrefs.SetPref_String "Tools", "ColorPanelPaletteFile", txtPaletteFile.Text
    
    m_CmdBarAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub cmdPaletteChoose_Click()
    Dim srcPaletteFile As String
    If Palettes.DisplayPaletteLoadDialog(vbNullString, srcPaletteFile) Then txtPaletteFile.Text = srcPaletteFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdateVisiblePanel()
    
    Dim i As Long
    For i = 0 To btsStyle.ListCount - 1
        pnlOptions(i).Visible = (i = btsStyle.ListIndex)
    Next i
    
End Sub
