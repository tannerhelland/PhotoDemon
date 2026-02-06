VERSION 5.00
Begin VB.Form toolpanel_Fill 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
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
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdDropDown cboSource 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdButtonStrip btsFillArea 
      Height          =   450
      Left            =   8280
      TabIndex        =   0
      Top             =   345
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   794
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdSlider sldFillTolerance 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      FontSizeCaption =   10
      Max             =   100
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdDropDown cboFillBlendMode 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   375
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   635
      Caption         =   "fill source"
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
         TabIndex        =   5
         Top             =   345
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdBrushSelector bsFillStyle 
         Height          =   735
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Caption         =   "brush style"
         FontSize        =   10
      End
      Begin PhotoDemon.pdSlider sldOpacity 
         CausesValidation=   0   'False
         Height          =   690
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2700
         _ExtentX        =   4763
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
      TabIndex        =   7
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
      Left            =   3600
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1508
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboFillAlphaMode 
         Height          =   735
         Left            =   120
         TabIndex        =   9
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
      TabIndex        =   11
      Top             =   0
      Width           =   2400
      _ExtentX        =   5292
      _ExtentY        =   635
      Caption         =   "tolerance"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   2
      Left            =   4800
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   14631
      _ExtentY        =   3625
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboFillCompare 
         Height          =   765
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1349
         Caption         =   "compare pixels by"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   345
      Index           =   3
      Left            =   8160
      TabIndex        =   14
      Top             =   0
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   609
      Caption         =   "fill area"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1455
      Index           =   3
      Left            =   8160
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   14631
      _ExtentY        =   3625
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   3
         Left            =   3000
         TabIndex        =   15
         Top             =   900
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonStrip btsFillMerge 
         Height          =   810
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1429
         Caption         =   "sample from"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdCheckBox chkAntialiasing 
         Height          =   375
         Left            =   225
         TabIndex        =   17
         Top             =   75
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "antialiased"
      End
   End
End
Attribute VB_Name = "toolpanel_Fill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Bucket Fill Tool Panel
'Copyright 2017-2026 by Tanner Helland
'Created: 30/August/17
'Last updated: 02/December/21
'Last update: migrate UI to new flyout design
'
'This form includes all user-editable settings for PD's bucket fill tool.
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

Private Sub bsFillStyle_BrushChanged()
    Tools_Fill.SetFillBrush bsFillStyle.Brush
End Sub

Private Sub bsFillStyle_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub bsFillStyle_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboSource.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

Private Sub btsFillArea_Click(ByVal buttonIndex As Long)
    Tools_Fill.SetFillSearchMode buttonIndex
End Sub

Private Sub btsFillArea_GotFocusAPI()
    UpdateFlyout 3, True
End Sub

Private Sub btsFillArea_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(3).hWnd
    Else
        newTargetHwnd = Me.chkAntialiasing.hWnd
    End If
End Sub

Private Sub btsFillMerge_Click(ByVal buttonIndex As Long)
    Tools_Fill.SetFillSampleMerged (buttonIndex = 0)
End Sub

Private Sub btsFillMerge_GotFocusAPI()
    UpdateFlyout 3, True
End Sub

Private Sub btsFillMerge_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkAntialiasing.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(3).hWnd
    End If
End Sub

Private Sub cboFillAlphaMode_Click()
    Tools_Fill.SetFillAlphaMode cboFillAlphaMode.ListIndex
End Sub

Private Sub cboFillAlphaMode_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub cboFillAlphaMode_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboFillBlendMode.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
    End If
End Sub

Private Sub cboFillBlendMode_Click()
    Tools_Fill.SetFillBlendMode cboFillBlendMode.ListIndex
End Sub

Private Sub cboFillBlendMode_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub cboFillBlendMode_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(1).hWnd
    Else
        newTargetHwnd = Me.cboFillAlphaMode.hWnd
    End If
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

Private Sub cboFillCompare_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub cboFillCompare_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sldFillTolerance.hWndSpinner
    Else
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    End If
End Sub

Private Sub cboSource_Click()
    
    sldOpacity.Visible = (cboSource.ListIndex = 0)
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

Private Sub cboSource_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub cboSource_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(0).hWnd
    Else
        If Me.sldOpacity.Visible Then
            newTargetHwnd = Me.sldOpacity.hWndSlider
        Else
            newTargetHwnd = Me.bsFillStyle.hWnd
        End If
    End If
End Sub

Private Sub chkAntialiasing_Click()
    Tools_Fill.SetFillAA chkAntialiasing.Value
End Sub

Private Sub chkAntialiasing_GotFocusAPI()
    UpdateFlyout 3, True
End Sub

Private Sub chkAntialiasing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsFillArea.hWnd
    Else
        newTargetHwnd = Me.btsFillMerge.hWnd
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
                If Me.sldOpacity.Visible Then
                    newTargetHwnd = Me.sldOpacity.hWndSpinner
                Else
                    newTargetHwnd = Me.bsFillStyle.hWnd
                End If
            Case 1
                newTargetHwnd = Me.cboFillAlphaMode.hWnd
            Case 2
                newTargetHwnd = Me.cboFillCompare.hWnd
            Case 3
                newTargetHwnd = Me.btsFillMerge.hWnd
        End Select
    Else
        Dim newIndex As Long
        newIndex = Index + 1
        If (newIndex > Me.ttlPanel.UBound) Then newIndex = Me.ttlPanel.lBound
        newTargetHwnd = Me.ttlPanel(newIndex).hWnd
    End If
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

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldFillTolerance_Change()
    Tools_Fill.SetFillTolerance sldFillTolerance.Value
End Sub

Private Sub sldFillTolerance_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub sldFillTolerance_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(2).hWnd
    Else
        newTargetHwnd = Me.cboFillCompare.hWnd
    End If
End Sub

Private Sub sldOpacity_Change()
    Tools_Fill.SetFillBrushOpacity sldOpacity.Value
End Sub

Private Sub sldOpacity_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub sldOpacity_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboSource.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        Dim newIndex As Long
        newIndex = Index - 1
        If (newIndex < ttlPanel.lBound) Then newIndex = ttlPanel.UBound
        newTargetHwnd = Me.cmdFlyoutLock(newIndex).hWnd
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.cboSource.hWnd
            Case 1
                newTargetHwnd = Me.cboFillBlendMode.hWnd
            Case 2
                newTargetHwnd = Me.sldFillTolerance.hWndSlider
            Case 3
                newTargetHwnd = Me.btsFillArea.hWnd
        End Select
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

    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    ApplyThemeAndTranslations Me
    
    bsFillStyle.AssignTooltip "Fills support many different styles.  Click to switch between solid color, pattern, and gradient styles."
    sldFillTolerance.AssignTooltip "Tolerance controls how similar two pixels must be before spreading the fill between them."
    btsFillMerge.AssignTooltip "Normally, fill operations analyze the entire image.  You can also analyze just the active layer."
    btsFillArea.AssignTooltip "Normally, fills spread out from a target pixel, adding neighboring pixels as it goes.  You can alternatively set it to analyze the entire image, without regard for continuity."
    cboFillCompare.AssignTooltip "This option controls how pixels are analyzed when adding them to the fill."

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
