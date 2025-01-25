VERSION 5.00
Begin VB.Form toolpanel_Crop 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
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
   Icon            =   "Toolpanel_Crop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButton cmdCommit 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   13
      Top             =   405
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   975
      Index           =   0
      Left            =   3720
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1095
      Index           =   1
      Left            =   8160
      Top             =   960
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   8705
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   1
         Left            =   120
         Top             =   0
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   423
         Caption         =   "display options"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   3210
         TabIndex        =   4
         Top             =   600
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
      Min             =   -32000
      Max             =   32000
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   420
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
      Min             =   -32000
      Max             =   32000
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   661
      Caption         =   "aspect ratio"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   2
      Left            =   2820
      TabIndex        =   5
      Top             =   420
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
      DefaultValue    =   1
      Min             =   1
      Max             =   32000
      Value           =   1
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   3
      Left            =   4380
      TabIndex        =   6
      Top             =   420
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
      DefaultValue    =   1
      Min             =   1
      Max             =   32000
      Value           =   1
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   7
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      Caption         =   "commit"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   2
      Left            =   30
      Top             =   30
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   423
      Caption         =   "position (x, y)"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   0
      Left            =   2700
      Top             =   30
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   423
      Caption         =   "size (w, h)"
   End
   Begin PhotoDemon.pdLabel lblColon 
      Height          =   375
      Index           =   0
      Left            =   7320
      Top             =   405
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   ":"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButtonToolbox cmdLock 
      Height          =   360
      Index           =   2
      Left            =   8760
      TabIndex        =   8
      Top             =   405
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   4
      Left            =   6240
      TabIndex        =   9
      Top             =   420
      Width           =   1080
      _ExtentX        =   2328
      _ExtentY        =   714
      DefaultValue    =   1
      Min             =   1
      Max             =   32000
      Value           =   1
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   5
      Left            =   7680
      TabIndex        =   10
      Top             =   420
      Width           =   1080
      _ExtentX        =   2328
      _ExtentY        =   714
      DefaultValue    =   1
      Min             =   1
      Max             =   32000
      Value           =   1
      ShowResetButton =   0   'False
   End
   Begin PhotoDemon.pdButtonToolbox cmdLock 
      Height          =   360
      Index           =   1
      Left            =   5460
      TabIndex        =   11
      Top             =   405
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdLock 
      Height          =   360
      Index           =   0
      Left            =   3900
      TabIndex        =   12
      Top             =   405
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdButton cmdCommit 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   14
      Top             =   405
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   661
   End
End
Attribute VB_Name = "toolpanel_Crop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Crop Tool Panel
'Copyright 2024-2025 by Tanner Helland
'Created: 11/Nov/24
'Last updated: 11/Nov/24
'Last update: initial build
'
'This form includes all user-editable settings for the Crop on-canvas tool.
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
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub cmdFlyoutLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            'If shiftTabWasPressed Then newTargetHwnd = Me.cboLayerResizeQuality.hWnd Else newTargetHwnd = Me.ttlPanel(1).hWnd
        Case 1
            'If shiftTabWasPressed Then newTargetHwnd = Me.sltLayerShearY.hWndSpinner Else newTargetHwnd = Me.ttlPanel(2).hWnd
        Case 2
            If shiftTabWasPressed Then
                'newTargetHwnd = Me.chkRotateNode.hWnd
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
    End Select
End Sub

Private Sub cmdLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    Dim lockedValue As Variant
    If (Index = 0) Then
        lockedValue = tudCrop(2).Value
    ElseIf (Index = 1) Then
        lockedValue = tudCrop(3).Value
    Else
        If (tudCrop(4).Value <> 0#) And (tudCrop(5).Value <> 0#) Then
            lockedValue = tudCrop(4).Value / tudCrop(5).Value
        Else
            lockedValue = 1#
        End If
    End If

    'When setting a new lock, unlock any other locks.
    If cmdLock(Index).Value Then
        If (Index = 0) Then
            cmdLock(1).Value = False
            cmdLock(2).Value = False
        ElseIf (Index = 1) Then
            cmdLock(0).Value = False
            cmdLock(2).Value = False
        ElseIf (Index = 2) Then
            cmdLock(0).Value = False
            cmdLock(1).Value = False
        End If
    End If
    
    If cmdLock(Index).Value Then
        Tools_Crop.LockProperty Index, lockedValue
    Else
        Tools_Crop.UnlockProperty Index
    End If
    
End Sub

Private Sub Form_Load()
    
    Tools.SetToolBusyState True
    
    'TODO:
    'Ensure our corresponding tool manager is synchronized with default layer rendering styles
    'Tools_Move.SetDrawDistances chkDistances.Value
    'Tools_Move.SetDrawLayerCornerNodes chkLayerNodes.Value
    'Tools_Move.SetDrawLayerRotateNodes chkRotateNode.Value
    'Tools_Move.SetMoveSelectedPixels_SampleMerged (btsSampleMerged.ListIndex = 0)
    'Tools_Move.SetMoveSelectedPixels_DefaultCut (btsCopyCut.ListIndex = 1)
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues True
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues True
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'Failsafe only
    If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    Set m_Flyout = Nothing
    
End Sub

'Whenever an active flyout panel is closed, we need to reset the matching titlebar to "closed" state
Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

'Because this tool synchronizes the vast majority of its properties to the active layer
' in the current image, we do *not* automatically save all settings for this tool.
' Instead, we custom-save only the settings we absolutely want.
Private Sub m_LastUsedSettings_AddCustomPresetData()

    'TODO:
    With m_lastUsedSettings
        '.AddPresetData "move-size-auto-activate", chkAutoActivateLayer.Value
        '.AddPresetData "move-size-ignore-transparent", chkIgnoreTransparent.Value
        '.AddPresetData "move-size-selection-sample-merged", btsSampleMerged.ListIndex
        '.AddPresetData "move-size-selection-default-cut", btsCopyCut.ListIndex
        '.AddPresetData "move-size-show-distances", chkDistances.Value
        '.AddPresetData "move-size-show-resize-nodes", chkLayerNodes.Value
        '.AddPresetData "move-size-show-rotate-nodes", chkRotateNode.Value
        '.AddPresetData "move-size-lock-aspect-ratio", chkAspectRatio.Value
    End With

End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

'TODO:
    With m_lastUsedSettings
'        chkAutoActivateLayer.Value = .RetrievePresetData("move-size-auto-activate", True)
'        chkIgnoreTransparent.Value = .RetrievePresetData("move-size-ignore-transparent", True)
'
'        btsSampleMerged.ListIndex = CLng(.RetrievePresetData("move-size-selection-sample-merged", "1"))
'        btsCopyCut.ListIndex = CLng(.RetrievePresetData("move-size-selection-default-cut", "1"))
'
'        chkDistances.Value = .RetrievePresetData("move-size-show-distances", False)
'        chkLayerNodes.Value = .RetrievePresetData("move-size-show-resize-nodes", True)
'        chkRotateNode.Value = .RetrievePresetData("move-size-show-rotate-nodes", True)
'
'        'The "lock aspect ratio" control is tricky, because we don't want to set this value
'        ' if the current image has variable aspect ratio enabled; otherwise, it will forcibly
'        ' modify the active layer's size value(s).
'        Dim okToLoad As Boolean
'        okToLoad = True
'
'        If Tools.CanvasToolsAllowed(False) Then
'            okToLoad = (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier() = 1#) And (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier() = 1#)
'        End If
'
'        If okToLoad Then
'            chkAspectRatio.Value = .RetrievePresetData("move-size-lock-aspect-ratio", False)
'        Else
'            chkAspectRatio.Value = False
'        End If
'
    End With
    
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
            Case 1
                newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
            Case 2
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
        End Select
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.tudCrop(0).hWnd
            Case 1
                'newTargetHwnd = Me.sltLayerAngle.hWndSlider
            Case 2
                'newTargetHwnd = Me.chkAutoActivateLayer.hWnd
        End Select
    End If
End Sub

Private Sub tudCrop_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.GetToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    Select Case Index

        'Crop position (x, y)
        Case 0
            If tudCrop(Index).IsValid Then Tools_Crop.RelayCropChangesFromUI pdd_Left, tudCrop(Index).Value
        
        Case 1
            If tudCrop(Index).IsValid Then Tools_Crop.RelayCropChangesFromUI pdd_Top, tudCrop(Index).Value

        'Crop size (w, h)
        Case 2
            If tudCrop(Index).IsValid Then Tools_Crop.RelayCropChangesFromUI pdd_Width, tudCrop(Index).Value
            
        Case 3
            If tudCrop(Index).IsValid Then Tools_Crop.RelayCropChangesFromUI pdd_Height, tudCrop(Index).Value
        
        'Aspect ratio (x, y)
        Case 4, 5
            
            'Because aspect ratio calculations involve division, ensure validity before continuing
            Dim aspW As Double, aspH As Double, aspFinal As Double, newDimension As Long
            
            If tudCrop(4).IsValid Then aspW = tudCrop(4).Value
            If tudCrop(5).IsValid Then aspH = tudCrop(5).Value
            
            If (aspW <> 0#) And (aspH <> 0#) Then
                
                'As a convenience, adjust width OR height depending on the aspect ratio value that is changing.
                ' (Note that we also have to account for locked width or height here - if a dimension is locked,
                '  we must adjust the *other* dimension.)
                If ((Index = 4) And (Not cmdLock(0).Value)) Or cmdLock(1).Value Then
                    aspFinal = aspW / aspH
                    newDimension = tudCrop(3).Value * aspFinal
                    Tools_Crop.RelayCropChangesFromUI pdd_AspectRatioW, newDimension, CSng(aspFinal)
                Else
                    aspFinal = aspH / aspW
                    newDimension = tudCrop(2).Value * aspFinal
                    Tools_Crop.RelayCropChangesFromUI pdd_AspectRatioH, newDimension, CSng(aspFinal)
                End If
                
            End If
                    
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub tudCrop_GotFocusAPI(Index As Integer)
    If (Index = 4) Or (Index = 5) Then UpdateFlyout 0, True
End Sub

Private Sub tudCrop_LostFocusAPI(Index As Integer)
    If (Not PDImages.IsImageActive()) Then Exit Sub
End Sub

'Because these controls are laid out in a non-standard pattern, we want to manually specify tab and
' shift+tab focus targets.
Private Sub tudCrop_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        If (Index > 0) Then
            newTargetHwnd = tudCrop(Index - 1).hWnd
        Else
            newTargetHwnd = Me.ttlPanel(0).hWnd
        End If
    Else
        If (Index < 3) Then
            newTargetHwnd = tudCrop(Index + 1).hWnd
        Else
            'newTargetHwnd = chkAspectRatio.hWnd
        End If
    End If

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'TODO:
    
    'UI images must be updated against theme-specific colors
    'Dim buttonSize As Long
    'buttonSize = Interface.FixDPI(32)
    'cmdLayerAffinePermanent.AssignImage "generic_commit", , buttonSize, buttonSize
    
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(16)
    
    Dim i As Long
    For i = 0 To cmdLock.Count - 1
        cmdLock(i).AssignImage "generic_unlock", , buttonSize, buttonSize
        cmdLock(i).AssignImage_Pressed "generic_lock", , buttonSize, buttonSize
        cmdLock(i).AssignTooltip "Lock this value.  (Only one value can be locked at a time.  If you lock a new value, previously locked values will unlock.)"
    Next i
    
    cmdCommit(0).AssignImage "generic_ok", , buttonSize, buttonSize
    cmdCommit(1).AssignImage "generic_cancel", , buttonSize, buttonSize
    
    'Tooltips must be localized
    'cmdLayerAffinePermanent.AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    Interface.ApplyThemeAndTranslations Me
    
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
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex, IIf(flyoutIndex > 0, Interface.FixDPI(-8), 0)
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
