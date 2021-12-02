VERSION 5.00
Begin VB.Form toolpanel_Paintbrush 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
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
   Icon            =   "Toolpanel_Paintbrush.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2415
      Index           =   2
      Left            =   5880
      Top             =   840
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4260
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   3360
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
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
         TabIndex        =   12
         Top             =   1800
         Width           =   3090
         _ExtentX        =   5450
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
         TabIndex        =   13
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
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
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
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
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
      TabIndex        =   6
      Top             =   0
      Width           =   2400
      _ExtentX        =   5292
      _ExtentY        =   635
      Caption         =   "blend mode"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   1
      Left            =   2640
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1508
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   0
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   1296
         Caption         =   "alpha mode"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   635
      Caption         =   "hardness"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_Paintbrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Paintbrush Tool Panel
'Copyright 2016-2021 by Tanner Helland
'Created: 31/Oct/16
'Last updated: 01/December/21
'Last update: migrate UI to new flyout design
'
'This form includes all user-editable settings for the "paintbrush" canvas tool.
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
            Tools_Paint.SetBrushBlendMode cboBrushSetting(Index).ListIndex
        
        'Alpha mode
        Case 1
            Tools_Paint.SetBrushAlphaMode cboBrushSetting(Index).ListIndex
            
    End Select
    
End Sub

Private Sub cboBrushSetting_GotFocusAPI(Index As Integer)
    UpdateFlyout 1, True
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
    End Select
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
        End Select
    Else
        Dim newIndex As Long
        newIndex = Index + 1
        If (newIndex > 2) Then newIndex = 0
        newTargetHwnd = Me.ttlPanel(newIndex).hWnd
    End If
End Sub

Private Sub Form_Load()
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboBrushSetting(0), BM_Normal
    Interface.PopulateAlphaModeDropDown cboBrushSetting(1), AM_Normal
    
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

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldSpacing_Change()
    Tools_Paint.SetBrushSpacing sldSpacing.Value
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
        If (newIndex < 0) Then newIndex = 2
        newTargetHwnd = Me.cmdFlyoutLock(newIndex).hWnd
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.sltBrushSetting(0).hWndSlider
            Case 1
                newTargetHwnd = Me.cboBrushSetting(0).hWnd
            Case 2
                newTargetHwnd = Me.sltBrushSetting(2).hWndSlider
        End Select
    End If
    
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllPaintbrushSettingsToUI()
    Tools_Paint.SetBrushSize sltBrushSetting(0).Value
    Tools_Paint.SetBrushOpacity sltBrushSetting(1).Value
    Tools_Paint.SetBrushHardness sltBrushSetting(2).Value
    Tools_Paint.SetBrushFlow sltBrushSetting(3).Value
    Tools_Paint.SetBrushSourceColor layerpanel_Colors.GetCurrentColor()
    Tools_Paint.SetBrushBlendMode cboBrushSetting(0).ListIndex
    Tools_Paint.SetBrushAlphaMode cboBrushSetting(1).ListIndex
    If (btsSpacing.ListIndex = 0) Then Tools_Paint.SetBrushSpacing 0# Else Tools_Paint.SetBrushSpacing sldSpacing.Value
End Sub

'If you want to synchronize all UI elements to match current paintbrush settings, use this function
Public Sub SyncUIToAllPaintbrushSettings()
    sltBrushSetting(0).Value = Tools_Paint.GetBrushSize()
    sltBrushSetting(1).Value = Tools_Paint.GetBrushOpacity()
    sltBrushSetting(2).Value = Tools_Paint.GetBrushHardness()
    sltBrushSetting(3).Value = Tools_Paint.GetBrushFlow()
    cboBrushSetting(0).ListIndex = Tools_Paint.GetBrushBlendMode()
    cboBrushSetting(1).ListIndex = Tools_Paint.GetBrushAlphaMode()
    If (Tools_Paint.GetBrushSpacing() = 0#) Then
        btsSpacing.ListIndex = 0
    Else
        btsSpacing.ListIndex = 1
        sldSpacing.Value = Tools_Paint.GetBrushSpacing()
    End If
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Flyout lock controls use the same behavior across all instances
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(16)
    
    Dim i As Long
    For i = cmdFlyoutLock.lBound To cmdFlyoutLock.UBound
        cmdFlyoutLock(i).AssignImage "generic_invisible", , buttonSize, buttonSize
        cmdFlyoutLock(i).AssignImage_Pressed "generic_visible", , buttonSize, buttonSize
        cmdFlyoutLock(i).AssignTooltip UserControls.GetCommonTranslation(pduct_FlyoutLockTooltip), UserControls.GetCommonTranslation(pduct_FlyoutLockTitle)
        cmdFlyoutLock(i).Value = False
    Next i
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me

End Sub

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
