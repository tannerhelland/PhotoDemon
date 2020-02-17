VERSION 5.00
Begin VB.Form toolpanel_Clone 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1023
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Left            =   10380
      Top             =   15
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      Caption         =   "source settings"
   End
   Begin PhotoDemon.pdDropDown cboBrushSetting 
      Height          =   735
      Index           =   2
      Left            =   10380
      TabIndex        =   7
      Top             =   690
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      Caption         =   "pattern mode"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboBrushSetting 
      Height          =   735
      Index           =   0
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "blend / alpha mode"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1217
      Caption         =   "size"
      FontSizeCaption =   10
      Min             =   1
      Max             =   1000
      SigDigits       =   1
      ScaleStyle      =   1
      ScaleExponent   =   3
      Value           =   1
      NotchPosition   =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1217
      Caption         =   "opacity"
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      Value           =   100
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboBrushSetting 
      Height          =   375
      Index           =   1
      Left            =   7905
      TabIndex        =   3
      Top             =   900
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   661
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1217
      Caption         =   "hardness"
      FontSizeCaption =   10
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdCheckBox chkSampleMerged 
      Height          =   345
      Left            =   12240
      TabIndex        =   5
      Top             =   345
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   609
      Caption         =   "sample all layers"
   End
   Begin PhotoDemon.pdCheckBox chkAligned 
      Height          =   345
      Left            =   10470
      TabIndex        =   6
      Top             =   345
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      Caption         =   "aligned"
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   3
      Left            =   3960
      TabIndex        =   8
      Top             =   660
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1217
      Caption         =   "flow"
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
End
Attribute VB_Name = "toolpanel_Clone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Clone Stamp Tool Panel
'Copyright 2016-2020 by Tanner Helland
'Created: 31/October/16
'Last updated: 16/September/19
'Last update: split off from normal softbrush tool panel
'
'This form includes all user-editable settings for the "clone stamp" canvas tool.
'
'Some brush settings in this panel are currently commented out.  This is not a bug - these features
' are already implemented in the tools_clone module, but PD's brush UI is being reworked to remove
' some of these dedicated controls in favor of merging them into a separate brush UI (where the user
' can pick brushes from a pre-built list or design their own, similar to the gradients dialog).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Private Sub btsSpacing_Click(ByVal buttonIndex As Long)
'    UpdateSpacingVisibility
'End Sub

Private Sub cboBrushSetting_Click(Index As Integer)

    Select Case Index
    
        'Blend mode
        Case 0
            Tools_Clone.SetBrushBlendMode cboBrushSetting(Index).ListIndex
        
        'Alpha mode
        Case 1
            Tools_Clone.SetBrushAlphaMode cboBrushSetting(Index).ListIndex
        
        'Wrap mode
        Case 2
            Tools_Clone.SetBrushWrapMode GetWrapModeFromIndex(cboBrushSetting(Index).ListIndex)
        
    End Select
    
End Sub

Private Sub chkAligned_Click()
    Tools_Clone.SetBrushAligned chkAligned.Value
End Sub

Private Sub chkSampleMerged_Click()
    Tools_Clone.SetBrushSampleMerged chkSampleMerged.Value
End Sub

Private Sub Form_Load()
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboBrushSetting(0), BM_Normal
    Interface.PopulateAlphaModeDropDown cboBrushSetting(1), AM_Normal
    
    cboBrushSetting(2).SetAutomaticRedraws False
    cboBrushSetting(2).AddItem "off", 0
    cboBrushSetting(2).AddItem "tile", 1
    cboBrushSetting(2).AddItem "tile + flip horizontal", 2
    cboBrushSetting(2).AddItem "tile + flip vertical", 3
    cboBrushSetting(2).AddItem "tile + flip both", 4
    cboBrushSetting(2).ListIndex = 0
    cboBrushSetting(2).SetAutomaticRedraws True, True
    
    ''Populate any other list-style UI elements
    'btsSpacing.AddItem "auto", 0
    'btsSpacing.AddItem "manual", 1
    'btsSpacing.ListIndex = 0
    'UpdateSpacingVisibility
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me

End Sub

'Private Sub sldSpacing_Change()
'    Tools_Clone.SetBrushSpacing sldSpacing.Value
'End Sub

Private Sub sltBrushSetting_Change(Index As Integer)
    
    Select Case Index
    
        'Radius
        Case 0
            Tools_Clone.SetBrushSize sltBrushSetting(Index).Value
        
        'Opacity
        Case 1
            Tools_Clone.SetBrushOpacity sltBrushSetting(Index).Value
            
        'Hardness
        Case 2
            Tools_Clone.SetBrushHardness sltBrushSetting(Index).Value
            
        'Flow
        Case 3
            Tools_Clone.SetBrushFlow sltBrushSetting(Index).Value
    
    End Select
    
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllPaintbrushSettingsToUI()
    Tools_Clone.SetBrushSize sltBrushSetting(0).Value
    Tools_Clone.SetBrushOpacity sltBrushSetting(1).Value
    Tools_Clone.SetBrushHardness sltBrushSetting(2).Value
    Tools_Clone.SetBrushBlendMode cboBrushSetting(0).ListIndex
    Tools_Clone.SetBrushAlphaMode cboBrushSetting(1).ListIndex
    Tools_Clone.SetBrushSampleMerged chkSampleMerged.Value
    Tools_Clone.SetBrushAligned chkAligned.Value
    Tools_Clone.SetBrushWrapMode GetWrapModeFromIndex(cboBrushSetting(2).ListIndex)
    Tools_Clone.SetBrushFlow sltBrushSetting(3).Value
    'If (btsSpacing.ListIndex = 0) Then Tools_Clone.SetBrushSpacing 0# Else Tools_Clone.SetBrushSpacing sldSpacing.Value
End Sub

'If you want to synchronize all UI elements to match current paintbrush settings, use this function
Public Sub SyncUIToAllPaintbrushSettings()
    sltBrushSetting(0).Value = Tools_Clone.GetBrushSize()
    sltBrushSetting(1).Value = Tools_Clone.GetBrushOpacity()
    sltBrushSetting(2).Value = Tools_Clone.GetBrushHardness()
    cboBrushSetting(0).ListIndex = Tools_Clone.GetBrushBlendMode()
    cboBrushSetting(1).ListIndex = Tools_Clone.GetBrushAlphaMode()
    chkSampleMerged.Value = Tools_Clone.GetBrushSampleMerged()
    chkAligned.Value = Tools_Clone.GetBrushAligned()
    cboBrushSetting(2).ListIndex = GetIndexFromWrapMode(Tools_Clone.GetBrushWrapMode())
    sltBrushSetting(3).Value = Tools_Clone.GetBrushFlow()
    'If (Tools_Clone.GetBrushSpacing() = 0#) Then
    '    btsSpacing.ListIndex = 0
    'Else
    '    btsSpacing.ListIndex = 1
    '    sldSpacing.Value = Tools_Clone.GetBrushSpacing()
    'End If
End Sub

'Helper functions to translate between dropdown "pattern mode" index and PD_2D_WrapMode enum
Private Function GetWrapModeFromIndex(ByVal srcIndex As Long) As PD_2D_WrapMode

    If (srcIndex = 0) Then
        GetWrapModeFromIndex = P2_WM_Clamp
    ElseIf (srcIndex = 1) Then
        GetWrapModeFromIndex = P2_WM_Tile
    ElseIf (srcIndex = 2) Then
        GetWrapModeFromIndex = P2_WM_TileFlipX
    ElseIf (srcIndex = 3) Then
        GetWrapModeFromIndex = P2_WM_TileFlipY
    ElseIf (srcIndex = 4) Then
        GetWrapModeFromIndex = P2_WM_TileFlipXY
    
    'Failsafe only; should never trigger
    Else
        GetWrapModeFromIndex = P2_WM_Clamp
    End If
    
End Function

Private Function GetIndexFromWrapMode(ByVal srcMode As PD_2D_WrapMode) As Long

    If (srcMode = P2_WM_Clamp) Then
        GetIndexFromWrapMode = 0
    ElseIf (srcMode = P2_WM_Tile) Then
        GetIndexFromWrapMode = 1
    ElseIf (srcMode = P2_WM_TileFlipX) Then
        GetIndexFromWrapMode = 2
    ElseIf (srcMode = P2_WM_TileFlipY) Then
        GetIndexFromWrapMode = 3
    ElseIf (srcMode = P2_WM_TileFlipXY) Then
        GetIndexFromWrapMode = 4
    
    'Failsafe only; should never trigger
    Else
        GetIndexFromWrapMode = 0
    End If

End Function

'Private Sub UpdateSpacingVisibility()
'    If (btsSpacing.ListIndex = 0) Then
'        sldSpacing.Visible = False
'        Tools_Clone.SetBrushSpacing 0#
'    Else
'        sldSpacing.Visible = True
'        Tools_Clone.SetBrushSpacing sldSpacing.Value
'    End If
'End Sub
