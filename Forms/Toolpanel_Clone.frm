VERSION 5.00
Begin VB.Form toolpanel_Clone 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13170
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
   Icon            =   "Toolpanel_Clone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   878
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1935
      Index           =   3
      Left            =   8280
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   3
         Left            =   3240
         TabIndex        =   16
         Top             =   1440
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboBrushSetting 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   1296
         Caption         =   "pattern mode"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsSampleMerged 
         Height          =   945
         Left            =   150
         TabIndex        =   18
         Top             =   0
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   1667
         Caption         =   "sample from"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdCheckBox chkAligned 
      Height          =   345
      Left            =   8400
      TabIndex        =   0
      Top             =   420
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   609
      Caption         =   "aligned"
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2415
      Index           =   2
      Left            =   5880
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4260
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   2880
         TabIndex        =   1
         Top             =   1815
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltBrushSetting 
         CausesValidation=   0   'False
         Height          =   690
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1217
         Caption         =   "flow"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdSlider sldSpacing 
         Height          =   495
         Left            =   180
         TabIndex        =   3
         Top             =   1800
         Width           =   2670
         _ExtentX        =   4710
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1508
         Caption         =   "spacing"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   360
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
      FontSizeCaption =   10
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdDropDown cboBrushSetting 
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   375
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
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
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   635
      Caption         =   "size"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Top             =   330
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltBrushSetting 
         CausesValidation=   0   'False
         Height          =   690
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   1217
         Caption         =   "opacity"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         Value           =   100
         DefaultValue    =   100
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Top             =   0
      Width           =   2400
      _ExtentX        =   5292
      _ExtentY        =   635
      Caption         =   "blend mode"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   2
      Left            =   5520
      TabIndex        =   12
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   635
      Caption         =   "hardness"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   1
      Left            =   2640
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1508
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboBrushSetting 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         Caption         =   "alpha mode"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   3
      Left            =   8400
      TabIndex        =   15
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   635
      Caption         =   "source settings"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_Clone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Clone Stamp Tool Panel
'Copyright 2016-2026 by Tanner Helland
'Created: 31/October/16
'Last updated: 01/December/21
'Last update: update UI to new flyout design
'
'This form includes all user-editable settings for the "clone stamp" canvas tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub btsSampleMerged_Click(ByVal buttonIndex As Long)
    Tools_Clone.SetBrushSampleMerged (buttonIndex = 0)
End Sub

Private Sub btsSampleMerged_GotFocusAPI()
    UpdateFlyout 3, True
End Sub

Private Sub btsSampleMerged_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkAligned.hWnd
    Else
        newTargetHwnd = Me.cboBrushSetting(2).hWnd
    End If
End Sub

Private Sub btsSpacing_Click(ByVal buttonIndex As Long)
    UpdateSpacingVisibility
End Sub

Private Sub btsSpacing_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub btsSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sltBrushSetting(3).hWndSpinner
    Else
        If Me.sldSpacing.Visible Then
            newTargetHwnd = Me.sldSpacing.hWndSlider
        Else
            newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
        End If
    End If
End Sub

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

Private Sub cboBrushSetting_GotFocusAPI(Index As Integer)
    If (Index < 2) Then
        UpdateFlyout 1, True
    Else
        UpdateFlyout 3, True
    End If
End Sub

Private Sub cboBrushSetting_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(1).hWnd
            Else
                newTargetHwnd = Me.cboBrushSetting(1).hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cboBrushSetting(0).hWnd
            Else
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.btsSampleMerged.hWnd
            Else
                newTargetHwnd = Me.cmdFlyoutLock(3).hWnd
            End If
    End Select
End Sub

Private Sub chkAligned_Click()
    Tools_Clone.SetBrushAligned chkAligned.Value
End Sub

Private Sub chkAligned_GotFocusAPI()
    UpdateFlyout 3, True
End Sub

Private Sub chkAligned_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(3).hWnd
    Else
        newTargetHwnd = Me.btsSampleMerged.hWnd
    End If
End Sub

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub cmdFlyoutLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = Me.sltBrushSetting(1).hWndSpinner
            Case 1
                newTargetHwnd = Me.cboBrushSetting(1).hWnd
            Case 2
                If Me.sldSpacing.Visible Then
                    newTargetHwnd = Me.sldSpacing.hWndSpinner
                Else
                    newTargetHwnd = Me.btsSpacing.hWnd
                End If
            Case 3
                newTargetHwnd = Me.cboBrushSetting(2).hWnd
        End Select
    Else
        Dim newIndex As Long
        newIndex = Index + 1
        If (newIndex > 3) Then newIndex = 0
        newTargetHwnd = Me.ttlPanel(newIndex).hWnd
    End If
End Sub

Private Sub Form_Load()
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboBrushSetting(0), BM_Normal
    Interface.PopulateAlphaModeDropDown cboBrushSetting(1), AM_Normal
    
    btsSampleMerged.AddItem "image", 0
    btsSampleMerged.AddItem "layer", 1
    btsSampleMerged.ListIndex = 0
    
    cboBrushSetting(2).SetAutomaticRedraws False
    cboBrushSetting(2).AddItem "off", 0
    cboBrushSetting(2).AddItem "tile", 1
    cboBrushSetting(2).AddItem "tile + flip horizontal", 2
    cboBrushSetting(2).AddItem "tile + flip vertical", 3
    cboBrushSetting(2).AddItem "tile + flip both", 4
    cboBrushSetting(2).ListIndex = 0
    cboBrushSetting(2).SetAutomaticRedraws True, True
    
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

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldSpacing_Change()
    Tools_Clone.SetBrushSpacing sldSpacing.Value
End Sub

Private Sub sldSpacing_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub sldSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsSpacing.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    End If
End Sub

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

Private Sub sltBrushSetting_GotFocusAPI(Index As Integer)
    Select Case Index
        Case 0, 1
            UpdateFlyout 0, True
        Case Else
            UpdateFlyout 2, True
    End Select
End Sub

Private Sub sltBrushSetting_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(0).hWnd
            Else
                newTargetHwnd = Me.sltBrushSetting(1).hWndSlider
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sltBrushSetting(0).hWndSpinner
            Else
                newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(2).hWnd
            Else
                newTargetHwnd = Me.sltBrushSetting(3).hWndSlider
            End If
        Case 3
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sltBrushSetting(2).hWndSpinner
            Else
                newTargetHwnd = Me.btsSpacing.hWnd
            End If
    End Select
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    Dim newIndex As Long
    
    If shiftTabWasPressed Then
        newIndex = Index - 1
        If (newIndex < 0) Then newIndex = 3
        newTargetHwnd = Me.cmdFlyoutLock(newIndex).hWnd
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.sltBrushSetting(0).hWndSlider
            Case 1
                newTargetHwnd = Me.cboBrushSetting(0).hWnd
            Case 2
                newTargetHwnd = Me.sltBrushSetting(2).hWndSlider
            Case 3
                newTargetHwnd = Me.chkAligned.hWnd
        End Select
    End If
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me

End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllPaintbrushSettingsToUI()
    Tools_Clone.SetBrushSize sltBrushSetting(0).Value
    Tools_Clone.SetBrushOpacity sltBrushSetting(1).Value
    Tools_Clone.SetBrushHardness sltBrushSetting(2).Value
    Tools_Clone.SetBrushBlendMode cboBrushSetting(0).ListIndex
    Tools_Clone.SetBrushAlphaMode cboBrushSetting(1).ListIndex
    Tools_Clone.SetBrushSampleMerged (btsSampleMerged.ListIndex = 0)
    Tools_Clone.SetBrushAligned chkAligned.Value
    Tools_Clone.SetBrushWrapMode GetWrapModeFromIndex(cboBrushSetting(2).ListIndex)
    Tools_Clone.SetBrushFlow sltBrushSetting(3).Value
    If (btsSpacing.ListIndex = 0) Then Tools_Clone.SetBrushSpacing 0# Else Tools_Clone.SetBrushSpacing sldSpacing.Value
End Sub

'If you want to synchronize all UI elements to match current paintbrush settings, use this function
Public Sub SyncUIToAllPaintbrushSettings()
    sltBrushSetting(0).Value = Tools_Clone.GetBrushSize()
    sltBrushSetting(1).Value = Tools_Clone.GetBrushOpacity()
    sltBrushSetting(2).Value = Tools_Clone.GetBrushHardness()
    cboBrushSetting(0).ListIndex = Tools_Clone.GetBrushBlendMode()
    cboBrushSetting(1).ListIndex = Tools_Clone.GetBrushAlphaMode()
    If Tools_Clone.GetBrushSampleMerged() Then btsSampleMerged.ListIndex = 0 Else btsSampleMerged.ListIndex = 1
    chkAligned.Value = Tools_Clone.GetBrushAligned()
    cboBrushSetting(2).ListIndex = GetIndexFromWrapMode(Tools_Clone.GetBrushWrapMode())
    sltBrushSetting(3).Value = Tools_Clone.GetBrushFlow()
    If (Tools_Clone.GetBrushSpacing() = 0#) Then
        btsSpacing.ListIndex = 0
    Else
        btsSpacing.ListIndex = 1
        sldSpacing.Value = Tools_Clone.GetBrushSpacing()
    End If
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

'Update the actively displayed flyout (if any).  Note that the flyout manager will automatically
' hide any other open flyouts, as necessary.
Private Sub UpdateFlyout(ByVal flyoutIndex As Long, Optional ByVal newState As Boolean = True)
    
    'Ensure we have a flyout manager
    If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
    
    'Exit if we're already in the process of synchronizing
    If m_Flyout.GetFlyoutSyncState() Then Exit Sub
    m_Flyout.SetFlyoutSyncState True
    
    'Ensure we have a flyout manager, then raise the corresponding panel
    If newState Then
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex, IIf(flyoutIndex = 0, 0, Interface.FixDPI(-8))
    Else
        If (flyoutIndex = m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.HideFlyout
    End If
    
    'Update titlebar state(s) to match
    Dim i As Long
    For i = ttlPanel.lBound To ttlPanel.UBound
        If (i = m_Flyout.GetFlyoutTrackerID()) Then
            If (Not ttlPanel(i).Value) Then ttlPanel(i).Value = True
        Else
            If ttlPanel(i).Value Then ttlPanel(i).Value = False
        End If
    Next i
    
    'Clear the synchronization flag before exiting
    m_Flyout.SetFlyoutSyncState False
    
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
