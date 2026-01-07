VERSION 5.00
Begin VB.Form dialog_ColorPanel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Color panel settings"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
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
   ScaleHeight     =   277
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
      Top             =   3420
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   2055
      Index           =   1
      Left            =   120
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
      Begin PhotoDemon.pdDropDown cboPalettes 
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1296
         Caption         =   "palettes in this file (%1)"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButton cmdPaletteChoose 
         Height          =   375
         Left            =   8280
         TabIndex        =   3
         Top             =   330
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   661
         Caption         =   "..."
      End
      Begin PhotoDemon.pdTextBox txtPaletteFile 
         Height          =   375
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   0
         Width           =   8655
         _ExtentX        =   15478
         _ExtentY        =   450
         Caption         =   "palette to use"
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   1335
      Index           =   0
      Left            =   120
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
'Copyright 2018-2026 by Tanner Helland
'Created: 13/February/18
'Last updated: 13/February/18
'Last update: initial build
'
'The right-side color panel now supports multiple color selection modes.  Hopefully this gives creators
' increased freedom when deciding how they want to paint an image.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog.  If the user cancels this dialog, default settings will be used.
Private m_CmdBarAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_CmdBarAnswer
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog()
    
    'Provide a default answer (in case the user closes the dialog via some means other than the command bar)
    m_CmdBarAnswer = vbCancel
    
    'Prep any dynamic UI objects
    btsStyle.AddItem "wheels + history", 0
    btsStyle.AddItem "palette", 1
    btsStyle.ListIndex = UserPrefs.GetPref_Long("Tools", "ColorPanelStyle", 0)
    txtPaletteFile.Text = UserPrefs.GetPref_String("Tools", "ColorPanelPaletteFile")
    
    'If the palette file is valid, update the group list to match
    If UpdatePaletteGroups() Then
        Dim curPaletteGroup As Long
        curPaletteGroup = UserPrefs.GetPref_Long("Tools", "ColorPanelPaletteGroup", -1)
        If (curPaletteGroup >= 0) And (curPaletteGroup < cboPalettes.ListCount) Then cboPalettes.ListIndex = curPaletteGroup
    End If
    
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
    
    'We need to write all preferences to PD's central user prefs manager before exiting
    
    'First, if the user selected "palette" color mode, we need to validate their selected palette
    ' before attempting to load it in the main window.
    Dim finalStyle As Long
    finalStyle = btsStyle.ListIndex
    If (finalStyle = 1) Then
        Dim tmpPalette As pdPalette
        Set tmpPalette = New pdPalette
        If (Not tmpPalette.LoadPaletteFromFile(txtPaletteFile.Text)) Then finalStyle = 0
    End If
    
    UserPrefs.SetPref_Long "Tools", "ColorPanelStyle", finalStyle
    
    'Palette mode requires a few other preferences; we can ignore these if palette mode isn't being used
    If (finalStyle = 1) Then
        UserPrefs.SetPref_String "Tools", "ColorPanelPaletteFile", txtPaletteFile.Text
        UserPrefs.SetPref_String "Tools", "ColorPanelPaletteGroup", cboPalettes.ListIndex
    End If
    
    m_CmdBarAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub cmdPaletteChoose_Click()
    Dim srcPaletteFile As String
    If Palettes.DisplayPaletteLoadDialog(vbNullString, srcPaletteFile) Then
        txtPaletteFile.Text = srcPaletteFile
        UpdatePaletteGroups
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function UpdatePaletteGroups() As Boolean

    Dim tmpPalette As pdPalette
    Set tmpPalette = New pdPalette
    If tmpPalette.LoadPaletteFromFile(txtPaletteFile.Text) Then
        
        If (Not g_Language Is Nothing) Then cboPalettes.Caption = g_Language.TranslateMessage("palettes in this file (%1)", tmpPalette.GetPaletteGroupCount)
        
        cboPalettes.SetAutomaticRedraws False
        cboPalettes.Clear
        
        Dim i As Long
        For i = 1 To tmpPalette.GetPaletteGroupCount
            cboPalettes.AddItem tmpPalette.GetPaletteName(i - 1), i - 1
        Next i
        
        cboPalettes.ListIndex = 0
        cboPalettes.SetAutomaticRedraws True, True
        cboPalettes.Visible = True
        
        UpdatePaletteGroups = True
        
    'If palette validation failed, clear various display bits
    Else
        'cboPalettes.Caption = g_Language.TranslateMessage("no valid palettes found")
        cboPalettes.Visible = False
    End If
        
End Function

Private Sub UpdateVisiblePanel()
    
    Dim i As Long
    For i = 0 To btsStyle.ListCount - 1
        pnlOptions(i).Visible = (i = btsStyle.ListIndex)
    Next i
    
    'Resize the form depending on the open panel
    If (Not g_WindowManager Is Nothing) Then
        
        Dim curWinRect As winRect, curClientRect As winRect
        g_WindowManager.GetWindowRect_API Me.hWnd, curWinRect
        g_WindowManager.GetClientWinRect Me.hWnd, curClientRect
        
        Dim ncHeight As Long
        ncHeight = (curWinRect.y2 - curWinRect.y1) - (curClientRect.y2 - curClientRect.y1) + cmdBar.GetHeight + Interface.FixDPI(8)
        
        If (btsStyle.ListIndex = 0) Then
            g_WindowManager.SetSizeByHWnd Me.hWnd, curWinRect.x2 - curWinRect.x1, ncHeight + (btsStyle.GetHeight + btsStyle.GetTop * 2), True
        
        ElseIf (btsStyle.ListIndex = 1) Then
            g_WindowManager.SetSizeByHWnd Me.hWnd, curWinRect.x2 - curWinRect.x1, ncHeight + (pnlOptions(1).GetTop + pnlOptions(1).GetHeight + btsStyle.GetTop), True
            
        End If
    End If
    
    'Ensure the listed palette group data is valid
    UpdatePaletteGroups
    
End Sub
