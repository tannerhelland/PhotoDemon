VERSION 5.00
Begin VB.Form toolpanel_Fill 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Toolpanel_Fill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   945
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "opacity"
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   435
      Left            =   885
      TabIndex        =   9
      Top             =   885
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   767
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdDropDown cboSource 
      Height          =   735
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Caption         =   "fill source"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdCheckBox chkAntialiasing 
      Height          =   375
      Left            =   12480
      TabIndex        =   7
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "antialiased"
   End
   Begin PhotoDemon.pdBrushSelector bsFillStyle 
      Height          =   495
      Left            =   165
      TabIndex        =   4
      Top             =   855
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   873
      FontSize        =   10
   End
   Begin PhotoDemon.pdDropDown cboFillCompare 
      Height          =   450
      Left            =   5340
      TabIndex        =   0
      Top             =   870
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
   End
   Begin PhotoDemon.pdButtonStripVertical btsFillArea 
      Height          =   1305
      Left            =   7920
      TabIndex        =   1
      Top             =   30
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2302
      Caption         =   "area"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sldFillTolerance 
      CausesValidation=   0   'False
      Height          =   675
      Left            =   5280
      TabIndex        =   2
      Top             =   30
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1191
      Caption         =   "tolerance"
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdButtonStripVertical btsFillMerge 
      Height          =   1305
      Left            =   10200
      TabIndex        =   3
      Top             =   30
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2302
      Caption         =   "sampling area"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboFillBlendMode 
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   30
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1296
      Caption         =   "blend / alpha mode"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboFillAlphaMode 
      Height          =   450
      Left            =   2985
      TabIndex        =   6
      Top             =   870
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   794
      FontSizeCaption =   10
   End
End
Attribute VB_Name = "toolpanel_Fill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Bucket Fill Tool Panel
'Copyright 2017-2021 by Tanner Helland
'Created: 30/August/17
'Last updated: 04/September/17
'Last update: continued work on initial build
'
'This form includes all user-editable settings for PD's bucket fill tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit


'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub bsFillStyle_BrushChanged()
    Tools_Fill.SetFillBrush bsFillStyle.Brush
End Sub

Private Sub btsFillArea_Click(ByVal buttonIndex As Long)
    Tools_Fill.SetFillSearchMode buttonIndex
End Sub

Private Sub btsFillMerge_Click(ByVal buttonIndex As Long)
    Tools_Fill.SetFillSampleMerged (buttonIndex = 0)
End Sub

Private Sub cboFillAlphaMode_Click()
    Tools_Fill.SetFillAlphaMode cboFillAlphaMode.ListIndex
End Sub

Private Sub cboFillBlendMode_Click()
    Tools_Fill.SetFillBlendMode cboFillBlendMode.ListIndex
End Sub

Private Sub cboFillCompare_Click()
    
    'Limit the accuracy of the tolerance for certain comparison methods.
    If (cboFillCompare.ListIndex > 1) Then
        sldFillTolerance.SigDigits = 0
    Else
        sldFillTolerance.SigDigits = 1
    End If
    
    Tools_Fill.SetFillCompareMode cboFillCompare.ListIndex
    
End Sub

Private Sub cboSource_Click()
    
    sldOpacity.Visible = (cboSource.ListIndex = 0)
    lblTitle(0).Visible = (cboSource.ListIndex = 0)
    bsFillStyle.Visible = (cboSource.ListIndex = 1)
    
    If (cboSource.ListIndex = 0) Then
        Tools_Fill.SetFillBrushSource fts_ColorOpacity
        Tools_Fill.SetFillBrushColor layerpanel_Colors.GetCurrentColor()
        Tools_Fill.SetFillBrushOpacity sldOpacity.Value
    Else
        Tools_Fill.SetFillBrushSource fts_CustomBrush
        Tools_Fill.SetFillBrush bsFillStyle.Brush
    End If
    
End Sub

Private Sub chkAntialiasing_Click()
    Tools_Fill.SetFillAA chkAntialiasing.Value
End Sub

Private Sub sldFillTolerance_Change()
    Tools_Fill.SetFillTolerance sldFillTolerance.Value
End Sub

Private Sub Form_Load()
    
    'Magic wand options
    cboSource.AddItem "current color", 0
    cboSource.AddItem "custom brush", 1
    cboSource.ListIndex = 0
    bsFillStyle.Visible = False
    
    btsFillMerge.AddItem "image", 0
    btsFillMerge.AddItem "layer", 1
    btsFillMerge.ListIndex = 0
    
    btsFillArea.AddItem "contiguous", 0
    btsFillArea.AddItem "global", 1
    btsFillArea.ListIndex = 0
    
    Interface.PopulateFloodFillTypes cboFillCompare
    Interface.PopulateBlendModeDropDown cboFillBlendMode
    Interface.PopulateAlphaModeDropDown cboFillAlphaMode
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If

End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllFillSettingsToUI()
    Tools_Fill.SetFillAA chkAntialiasing.Value
    Tools_Fill.SetFillAlphaMode cboFillAlphaMode.ListIndex
    Tools_Fill.SetFillBlendMode cboFillBlendMode.ListIndex
    Tools_Fill.SetFillBrush bsFillStyle.Brush
    Tools_Fill.SetFillBrushColor layerpanel_Colors.GetCurrentColor()
    Tools_Fill.SetFillBrushOpacity sldOpacity.Value
    Tools_Fill.SetFillCompareMode cboFillCompare.ListIndex
    Tools_Fill.SetFillSampleMerged (btsFillMerge.ListIndex = 0)
    Tools_Fill.SetFillSearchMode btsFillArea.ListIndex
    Tools_Fill.SetFillTolerance sldFillTolerance.Value
End Sub

Public Sub UpdateAgainstCurrentTheme()

    ApplyThemeAndTranslations Me
    
    bsFillStyle.AssignTooltip "Fills support many different styles.  Click to switch between solid color, pattern, and gradient styles."
    sldFillTolerance.AssignTooltip "Tolerance controls how similar two pixels must be before spreading the fill between them."
    btsFillMerge.AssignTooltip "Normally, fill operations analyze the entire image.  You can also analyze just the active layer."
    btsFillArea.AssignTooltip "Normally, fills spread out from a target pixel, adding neighboring pixels as it goes.  You can alternatively set it to analyze the entire image, without regard for continuity."
    cboFillCompare.AssignTooltip "This option controls how pixels are analyzed when adding them to the fill."

End Sub

Private Sub sldOpacity_Change()
    Tools_Fill.SetFillBrushOpacity sldOpacity.Value
End Sub
