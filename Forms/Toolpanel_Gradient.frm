VERSION 5.00
Begin VB.Form toolpanel_Gradient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
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
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdGradientSelector grdPrimary 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2143
      Caption         =   "gradient"
      FontSize        =   10
   End
   Begin PhotoDemon.pdDropDown cboSetting 
      Height          =   735
      Index           =   0
      Left            =   5520
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "blend / alpha mode"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sldSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1217
      Caption         =   "opacity"
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      Value           =   100
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboSetting 
      Height          =   375
      Index           =   1
      Left            =   5625
      TabIndex        =   2
      Top             =   900
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   661
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboSetting 
      Height          =   735
      Index           =   2
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "repeat"
      FontSizeCaption =   10
   End
End
Attribute VB_Name = "toolpanel_Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Tool Panel
'Copyright 2018-2019 by Tanner Helland
'Created: 31/December/18
'Last updated: 31/December/18
'Last update: initial build
'
'This form includes all user-editable settings for the "gradient" canvas tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub cboSetting_Click(Index As Integer)

    Select Case Index
    
        'Blend mode
        Case 0
            Tools_Gradient.SetGradientBlendMode cboSetting(Index).ListIndex
        
        'Alpha mode
        Case 1
            Tools_Gradient.SetGradientAlphaMode cboSetting(Index).ListIndex
            
        Case 2
            Tools_Gradient.SetGradientRepeat cboSetting(Index).ListIndex
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboSetting(0), BL_NORMAL
    Interface.PopulateAlphaModeDropDown cboSetting(1), LA_NORMAL
    
    'Populate any custom dropdowns
    cboSetting(2).SetAutomaticRedraws False
    cboSetting(2).Clear
    cboSetting(2).AddItem "none", 0
    cboSetting(2).AddItem "wrap", 1
    cboSetting(2).AddItem "reflect", 2
    cboSetting(2).SetAutomaticRedraws True, True
    
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

Private Sub sldSetting_Change(Index As Integer)
    
    Select Case Index
        
        'Opacity
        Case 0
            Tools_Gradient.SetGradientOpacity sldSetting(Index).Value
    
    End Select
    
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllGradientSettingsToUI()
    Tools_Gradient.SetGradientOpacity sldSetting(0).Value
    Tools_Gradient.SetGradientBlendMode cboSetting(0).ListIndex
    Tools_Gradient.SetGradientAlphaMode cboSetting(1).ListIndex
    Tools_Gradient.SetGradientRepeat cboSetting(2).ListIndex
    'If chkAntialiasing.Value Then tools_gradient.SetgradientAntialiasing P2_AA_HighQuality Else tools_gradient.SetgradientAntialiasing P2_AA_None
End Sub

'If you want to synchronize all UI elements to match current paintgradient settings, use this function
Public Sub SyncUIToAllGradientSettings()
    sldSetting(0).Value = Tools_Gradient.GetGradientOpacity
    cboSetting(0).ListIndex = Tools_Gradient.GetGradientBlendMode()
    cboSetting(1).ListIndex = Tools_Gradient.GetGradientAlphaMode()
    cboSetting(2).ListIndex = Tools_Gradient.GetGradientRepeat()
    'chkAntialiasing.Value = (tools_gradient.GetgradientAntialiasing = P2_AA_HighQuality)
End Sub
