VERSION 5.00
Begin VB.Form dialog_ExportLUT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
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
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsQuality 
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdSlider sldGridPoints 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
      Min             =   2
      Max             =   64
      Value           =   17
      NotchPosition   =   2
      NotchValueCustom=   17
   End
   Begin PhotoDemon.pdTextBox txtDescription 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      Caption         =   "description"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   1320
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      Caption         =   "copyright"
      FontSize        =   12
   End
   Begin PhotoDemon.pdTextBox txtCopyright 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   2
      Left            =   120
      Top             =   2520
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      Caption         =   "grid points"
      FontSize        =   12
   End
End
Attribute VB_Name = "dialog_ExportLUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color lookup table (LUT) export dialog
'Copyright 2022-2026 by Tanner Helland
'Created: 17/June/22
'Last updated: 17/June/22
'Last update: initial build
'
'This dialog works as a simple relay to the pdLUT3D class.  Look there for specific encoding details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Grid point values for fast/default/extreme modes
Private Const GRID_FAST As Long = 8
Private Const GRID_DEFAULT As Long = 17
Private Const GRID_EXTREME As Long = 32

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    Select Case buttonIndex
        Case 0
            sldGridPoints.Value = GRID_FAST
        Case 1
            sldGridPoints.Value = GRID_DEFAULT
        Case 2
            sldGridPoints.Value = GRID_EXTREME
        Case Else
            'Do nothing
    End Select
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "description", txtDescription.Text, True, False
    cParams.AddParam "copyright", txtCopyright.Text, True, False
    cParams.AddParam "grid-points", sldGridPoints.Value, True, False
    
    m_FormatParamString = cParams.GetParamString
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_ResetClick()
    btsQuality.ListIndex = 1
    txtDescription.Text = vbNullString
    txtCopyright.Text = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate any list elements
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "standard", 1
    btsQuality.AddItem "extreme", 2
    btsQuality.AddItem "custom", 3
    btsQuality.ListIndex = 1
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "Color lookup")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

Private Sub sldGridPoints_Change()
    If (sldGridPoints.Value = GRID_FAST) Then
        btsQuality.ListIndex = 0
    ElseIf (sldGridPoints.Value = GRID_DEFAULT) Then
        btsQuality.ListIndex = 1
    ElseIf (sldGridPoints.Value = GRID_EXTREME) Then
        btsQuality.ListIndex = 2
    Else
        btsQuality.ListIndex = 3
    End If
End Sub
