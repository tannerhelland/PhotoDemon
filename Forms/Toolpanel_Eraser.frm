VERSION 5.00
Begin VB.Form toolpanel_Eraser 
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
   Icon            =   "Toolpanel_Eraser.frx":0000
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
   Begin PhotoDemon.pdSlider sldSpacing 
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Min             =   1
      Max             =   1000
      ScaleStyle      =   1
      ScaleExponent   =   5
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdButtonStrip btsSpacing 
      Height          =   855
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1508
      Caption         =   "spacing"
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
      Max             =   2000
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
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   2
      Left            =   3960
      TabIndex        =   2
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
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   690
      Index           =   3
      Left            =   3960
      TabIndex        =   5
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
Attribute VB_Name = "toolpanel_Eraser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Eraser Tool Panel
'Copyright 2016-2021 by Tanner Helland
'Created: 31/Oct/16
'Last updated: 28/March/18
'Last update: forcibly set scratch layer alpha mode to "normal"; otherwise, switching from a paintbrush in
'             "locked" alpha mode will cause this tool to inherit that "locked" setting, preventing the
'             eraser from working at all!
'
'This form includes all user-editable settings for the "eraser" canvas tool.
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

Private Sub btsSpacing_Click(ByVal buttonIndex As Long)
    UpdateSpacingVisibility
End Sub

Private Sub Form_Load()
    
    'Populate any other list-style UI elements
    btsSpacing.AddItem "auto", 0
    btsSpacing.AddItem "manual", 1
    btsSpacing.ListIndex = 0
    UpdateSpacingVisibility
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
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

Private Sub sldSpacing_Change()
    Tools_Paint.SetBrushSpacing sldSpacing.Value
End Sub

Private Sub sltBrushSetting_Change(Index As Integer)
    
    Select Case Index
    
        'Radius
        Case 0
            Tools_Paint.SetBrushSize sltBrushSetting(Index).Value
        
        'Opacity
        Case 1
            Tools_Paint.SetBrushOpacity sltBrushSetting(Index).Value
            
        'Hardness
        Case 2
            Tools_Paint.SetBrushHardness sltBrushSetting(Index).Value
            
        'Flow
        Case 3
            Tools_Paint.SetBrushFlow sltBrushSetting(Index).Value
    
    End Select
    
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllPaintbrushSettingsToUI()
    Tools_Paint.SetBrushSize sltBrushSetting(0).Value
    Tools_Paint.SetBrushOpacity sltBrushSetting(1).Value
    Tools_Paint.SetBrushHardness sltBrushSetting(2).Value
    Tools_Paint.SetBrushFlow sltBrushSetting(3).Value
    Tools_Paint.SetBrushBlendMode BM_Erase
    Tools_Paint.SetBrushAlphaMode AM_Normal
    If (btsSpacing.ListIndex = 0) Then Tools_Paint.SetBrushSpacing 0# Else Tools_Paint.SetBrushSpacing sldSpacing.Value
End Sub

'If you want to synchronize all UI elements to match current paintbrush settings, use this function
Public Sub SyncUIToAllPaintbrushSettings()
    sltBrushSetting(0).Value = Tools_Paint.GetBrushSize()
    sltBrushSetting(1).Value = Tools_Paint.GetBrushOpacity()
    sltBrushSetting(2).Value = Tools_Paint.GetBrushHardness()
    sltBrushSetting(3).Value = Tools_Paint.GetBrushFlow()
    If (Tools_Paint.GetBrushSpacing() = 0#) Then
        btsSpacing.ListIndex = 0
    Else
        btsSpacing.ListIndex = 1
        sldSpacing.Value = Tools_Paint.GetBrushSpacing()
    End If
End Sub

Private Sub UpdateSpacingVisibility()
    If (btsSpacing.ListIndex = 0) Then
        sldSpacing.Visible = False
        Tools_Paint.SetBrushSpacing 0#
    Else
        sldSpacing.Visible = True
        Tools_Paint.SetBrushSpacing sldSpacing.Value
    End If
End Sub
