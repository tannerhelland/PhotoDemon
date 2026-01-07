VERSION 5.00
Begin VB.Form toolpanel_Gradient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
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
   Icon            =   "Toolpanel_Gradient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdSlider sldOffset 
      Height          =   735
      Left            =   8160
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1296
      Caption         =   "center offset"
      FontSizeCaption =   10
      Min             =   -99
      Max             =   99
      SigDigits       =   1
      Value           =   75
      DefaultValue    =   75
   End
   Begin PhotoDemon.pdGradientSelector grdPrimary 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   794
      FontSize        =   10
   End
   Begin PhotoDemon.pdDropDown cboSetting 
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   1
      Top             =   375
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboSetting 
      Height          =   375
      Index           =   0
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
      Caption         =   "gradient"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   2760
         TabIndex        =   5
         Top             =   330
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sldSetting 
         CausesValidation=   0   'False
         Height          =   690
         Index           =   0
         Left            =   0
         TabIndex        =   6
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
      Left            =   3120
      Top             =   2160
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
      Begin PhotoDemon.pdDropDown cboSetting 
         Height          =   735
         Index           =   1
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
      TabIndex        =   10
      Top             =   0
      Width           =   2400
      _ExtentX        =   5292
      _ExtentY        =   635
      Caption         =   "shape"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   855
      Index           =   2
      Left            =   6000
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   14631
      _ExtentY        =   3625
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboSetting 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1296
         Caption         =   "repeat"
         FontSizeCaption =   10
      End
   End
End
Attribute VB_Name = "toolpanel_Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Tool Panel
'Copyright 2018-2026 by Tanner Helland
'Created: 31/December/18
'Last updated: 02/December/21
'Last update: migrate UI to new flyout design
'
'This form includes all user-editable settings for the "gradient" canvas tool.
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

Private Sub cboSetting_Click(Index As Integer)

    Select Case Index
    
        'Blend mode
        Case 0
            Tools_Gradient.SetGradientBlendMode cboSetting(Index).ListIndex
        
        'Alpha mode
        Case 1
            Tools_Gradient.SetGradientAlphaMode cboSetting(Index).ListIndex
        
        'Shape
        Case 2
            Tools_Gradient.SetGradientShape cboSetting(Index).ListIndex
            sldOffset.Visible = (cboSetting(Index).ListIndex = gs_Spherical)
        
        'Wrap
        Case 3
            Tools_Gradient.SetGradientRepeat cboSetting(Index).ListIndex
            
    End Select
    
End Sub

Private Sub cboSetting_GotFocusAPI(Index As Integer)
    If ((Index = 0) Or (Index = 1)) Then
        UpdateFlyout 1, True
    ElseIf (Index = 2) Or (Index = 3) Then
        UpdateFlyout 2, True
    End If
End Sub

Private Sub cboSetting_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(1).hWnd
            Else
                newTargetHwnd = Me.cboSetting(1).hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cboSetting(0).hWnd
            Else
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(2).hWnd
            Else
                newTargetHwnd = Me.cboSetting(3).hWnd
            End If
        Case 3
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cboSetting(2).hWnd
            Else
                newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
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
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sldSetting(0).hWndSpinner
            Else
                newTargetHwnd = Me.ttlPanel(Index + 1).hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cboSetting(1).hWnd
            Else
                newTargetHwnd = Me.ttlPanel(Index + 1).hWnd
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cboSetting(3).hWnd
            Else
                If Me.sldOffset.Visible Then
                    newTargetHwnd = Me.sldOffset.hWndSlider
                Else
                    newTargetHwnd = Me.ttlPanel(0).hWnd
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboSetting(0), BM_Normal
    Interface.PopulateAlphaModeDropDown cboSetting(1), AM_Normal
    
    'Populate any custom dropdowns
    cboSetting(2).SetAutomaticRedraws False
    cboSetting(2).Clear
    cboSetting(2).AddItem "linear", 0
    cboSetting(2).AddItem "reflection", 1
    cboSetting(2).AddItem "radial", 2
    cboSetting(2).AddItem "spherical", 3
    cboSetting(2).AddItem "square", 4
    cboSetting(2).AddItem "diamond", 5
    cboSetting(2).AddItem "conical", 6
    cboSetting(2).AddItem "spiral", 7
    cboSetting(2).ListIndex = 0
    cboSetting(2).SetAutomaticRedraws True, True
    
    cboSetting(3).SetAutomaticRedraws False
    cboSetting(3).Clear
    cboSetting(3).AddItem "none", 0
    cboSetting(3).AddItem "clamp", 1
    cboSetting(3).AddItem "wrap", 2
    cboSetting(3).AddItem "reflect", 3
    cboSetting(3).ListIndex = 1
    cboSetting(3).SetAutomaticRedraws True, True
    
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

Private Sub grdPrimary_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub grdPrimary_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(0).hWnd
    Else
        newTargetHwnd = Me.sldSetting(0).hWnd
    End If
End Sub

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldOffset_Change()
    Tools_Gradient.SetGradientRadialOffset sldOffset.Value * 0.01
End Sub

Private Sub sldOffset_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub sldOffset_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    Else
        newTargetHwnd = Me.ttlPanel(0).hWnd
    End If
End Sub

Private Sub sldSetting_Change(Index As Integer)
    
    Select Case Index
        
        'Opacity
        Case 0
            Tools_Gradient.SetGradientOpacity sldSetting(Index).Value
    
    End Select
    
End Sub

Private Sub sldSetting_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub sldSetting_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.grdPrimary.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                If Me.sldOffset.Visible Then
                    newTargetHwnd = Me.sldOffset.hWndSpinner
                Else
                    newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
                End If
            Else
                newTargetHwnd = Me.grdPrimary.hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(Index - 1).hWnd
            Else
                newTargetHwnd = Me.cboSetting(0).hWnd
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(Index - 1).hWnd
            Else
                newTargetHwnd = Me.cboSetting(2).hWnd
            End If
    End Select
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllGradientSettingsToUI()
    Tools_Gradient.SetGradientOpacity sldSetting(0).Value
    Tools_Gradient.SetGradientBlendMode cboSetting(0).ListIndex
    Tools_Gradient.SetGradientAlphaMode cboSetting(1).ListIndex
    Tools_Gradient.SetGradientShape cboSetting(2).ListIndex
    Tools_Gradient.SetGradientRepeat cboSetting(3).ListIndex
    Tools_Gradient.SetGradientRadialOffset sldOffset.Value * 0.01
End Sub

'If you want to synchronize all UI elements to match current paintgradient settings, use this function
Public Sub SyncUIToAllGradientSettings()
    sldSetting(0).Value = Tools_Gradient.GetGradientOpacity
    cboSetting(0).ListIndex = Tools_Gradient.GetGradientBlendMode()
    cboSetting(1).ListIndex = Tools_Gradient.GetGradientAlphaMode()
    cboSetting(2).ListIndex = Tools_Gradient.GetGradientShape()
    cboSetting(3).ListIndex = Tools_Gradient.GetGradientRepeat()
    sldOffset.Value = Tools_Gradient.GetGradientRadialOffset * 100#
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
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
