VERSION 5.00
Begin VB.Form toolpanel_Eraser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
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
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   3000
      Index           =   1
      Left            =   3840
      Top             =   840
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   5292
      Begin PhotoDemon.pdCheckBox chkStrictPixel 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "align to pixel grid"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   3360
         TabIndex        =   0
         Top             =   2430
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
         TabIndex        =   1
         Top             =   600
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
         TabIndex        =   2
         Top             =   2400
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
         TabIndex        =   3
         Top             =   1440
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
      Left            =   2880
      TabIndex        =   4
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
   Begin PhotoDemon.pdSlider sltBrushSetting 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   5
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
      TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
      TabIndex        =   9
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   635
      Caption         =   "hardness"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_Eraser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Eraser Tool Panel
'Copyright 2016-2026 by Tanner Helland
'Created: 31/Oct/16
'Last updated: 01/December/21
'Last update: migrate UI to new flyout design
'
'This form includes all user-editable settings for the "eraser" canvas tool.
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
    UpdateFlyout 1, True
End Sub

Private Sub btsSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sltBrushSetting(3).hWndSpinner
    Else
        If Me.sldSpacing.Visible Then
            newTargetHwnd = Me.sldSpacing.hWndSlider
        Else
            newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
        End If
    End If
End Sub

Private Sub chkStrictPixel_Click()
    Tools_Paint.SetStrictPixelAlignment chkStrictPixel.Value
End Sub

Private Sub chkStrictPixel_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub chkStrictPixel_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sltBrushSetting(2).hWndSpinner
    Else
        newTargetHwnd = Me.sltBrushSetting(3).hWndSlider
    End If
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
                newTargetHwnd = Me.sltBrushSetting(1).hWndSpinner
            Else
                newTargetHwnd = Me.ttlPanel(1).hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                If Me.sldSpacing.Visible Then
                    newTargetHwnd = Me.sldSpacing.hWnd
                Else
                    newTargetHwnd = Me.btsSpacing.hWnd
                End If
            Else
                If Me.ttlPanel(0).Enabled Then
                    newTargetHwnd = Me.ttlPanel(0).hWnd
                Else
                    newTargetHwnd = Me.sltBrushSetting(0).hWndSlider
                End If
            End If
    End Select
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

Private Sub Form_Resize()
    ReflowUI
End Sub

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldSpacing_Change()
    Tools_Paint.SetBrushSpacing sldSpacing.Value
End Sub

Private Sub sldSpacing_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub sldSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsSpacing.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
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
            If Me.ttlPanel(0).Enabled Then UpdateFlyout 0, True
        Case 2, 3
            UpdateFlyout 1, True
    End Select
End Sub

Private Sub sltBrushSetting_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                If Me.ttlPanel(0).Enabled Then
                    newTargetHwnd = Me.ttlPanel(0).hWnd
                Else
                    newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
                End If
            Case 1
                newTargetHwnd = Me.sltBrushSetting(0).hWndSpinner
            Case 2
                newTargetHwnd = Me.ttlPanel(1).hWnd
            Case 3
                newTargetHwnd = chkStrictPixel.hWnd
        End Select
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.sltBrushSetting(1).hWndSlider
            Case 1
                If Me.ttlPanel(0).Enabled Then
                    newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
                Else
                    newTargetHwnd = Me.ttlPanel(1).hWnd
                End If
            Case 2
                newTargetHwnd = Me.chkStrictPixel.hWnd
            Case 3
                newTargetHwnd = Me.btsSpacing.hWnd
        End Select
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                newTargetHwnd = Me.sltBrushSetting(0).hWndSlider
            End If
        Case 1
            If shiftTabWasPressed Then
                If Me.ttlPanel(0).Enabled Then
                    newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
                Else
                    newTargetHwnd = Me.sltBrushSetting(1).hWndSpinner
                End If
            Else
                newTargetHwnd = Me.sltBrushSetting(2).hWnd
            End If
    End Select
End Sub

'When the form is resized, we can possibly move some controls out of their flyout panels and into
' the main toolpanel area.
Private Sub ReflowUI()
    
    'Skip reflow in designer mode
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'High-DPI displays cause trouble for measurements using internal VB layout properties.
    ' Use WAPI or PD-specific layout properties (GetLeft, GetWidth etc) for correct measurements.
    Dim parentWidth As Long, parentHeight As Long
    If (Not g_WindowManager Is Nothing) Then
        parentWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        parentHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    Else
        Exit Sub
    End If
    
    'Solve for available width, plus standardized padding
    Dim stdPadding As Long, stdPaddingTitle As Long
    stdPadding = Interface.FixDPI(20)
    stdPaddingTitle = Interface.FixDPI(8)
    
    'Determine if we're already in "wide screen mode" or "narrow screen mode"
    Dim inWideModeNow As Boolean
    inWideModeNow = (Not Me.ttlPanel(0).Enabled)
    
    'We now need to branch according to "already using wide toolbar layout"
    ' 1) If we're already in "wide mode", we need to ensure the available space hasn't shrunk too far.
    '    (Basically, see if the screen is too small and move stuff back into flyouts.)
    ' 2) If we're not in "wide mode", we need to see if we have available space to expand the layout.
    '    (Basically, is there room to stick flyout UI bits into the main toolpanel.)
    Dim useWideMode As Boolean
    Dim availablePixels As Long, minAvailableLeft As Long
    
    'If in wide mode, see if the toolpanel has gotten too crammed
    If inWideModeNow Then
        
        'The minimum available left position needs to be calculated against the right-most control
        ' in the toolpanel.  This must be custom-coded because that control varies by toolpanel,
        ' and some controls (like checkboxes) will deliberately make themselves as wide as possible
        ' to allow for long translations in non-US locales, so a special Width() function needs
        ' to be called.
        minAvailableLeft = Me.ttlPanel(1).GetLeft + Me.ttlPanel(1).GetWidth + stdPaddingTitle
        useWideMode = (minAvailableLeft < parentWidth)
        
    'If *not* in "wide mode", see if we have enough spare space to activate wide mode.
    Else
    
        minAvailableLeft = Me.ttlPanel(1).GetLeft + Me.ttlPanel(1).GetWidth + stdPadding + stdPaddingTitle
        availablePixels = parentWidth - minAvailableLeft
        
        'useWideMode needs to now be compared against the object we want to move into the toolpanel
        ' (in this case, the brush opacity slider).
        useWideMode = (availablePixels > Me.sltBrushSetting(1).GetWidth)
        
    End If
    
    Dim xOffset As Long
    
    'We now need to compare "useWideMode" to "inWideModeNow", and ensure the two values are in sync.
    If useWideMode Then
        
        'If we're already in "wide mode", we don't need to move anything!
        If (Not inWideModeNow) Then
            
            'Before doing anything else, hide any open flyouts
            UserControls.HideOpenFlyouts 0&
            
            'In this run, we're targeting the Opacity slider for inclusion in the toolpanel.
            
            'For this particular control, we can actually move the entire flyout panel into the toolpanel,
            ' but we are *not* sticking it at the end of the panel - instead, we're sticking it next to the
            ' size slider, which is its natural position in the flyout order.  This means we need to shift
            ' all controls after it to the right.
            
            'Start by moving the target control into position, and note that we use the opacity slider's
            ' width here (*not* the panel's width, as it includes the panel flyout lock button).
            xOffset = Me.ttlPanel(0).GetLeft + Me.ttlPanel(0).GetWidth + stdPadding
            cntrPopOut(0).SetPosition xOffset, 0
            xOffset = xOffset + Me.sltBrushSetting(1).GetWidth + stdPadding
            
            'Because the top panel uses a slightly taller layout (to account for taller controls),
            ' we need to slightly increase padding of the new slider to make it align.
            Me.sltBrushSetting(1).CaptionPadding = 2
            
            'Send the flyout panel to the back of the zorder so we don't have to mess with resizing it.
            cntrPopOut(0).ZOrder vbSendToBack
            
            'Shift everything past this control to the right.
            ' (This step could probably be automated across windows, but it would require a *lot* more code.)
            Me.ttlPanel(1).SetLeft xOffset
            Me.sltBrushSetting(2).SetLeft xOffset
            
            'Now forcibly disable (or enable) all controls associated with the old flyout,
            ' including the parent titlebar of the flyout and the panel lock button *on* the flyout.
            Me.cmdFlyoutLock(0).Visible = False
            Me.cntrPopOut(0).Visible = True
            Me.ttlPanel(0).Enabled = False
            
        End If
        
    'Argh, there's not enough room to expand the toolpanel.  If we're currently using wide mode,
    ' we must remove any embedded flyouts, while also re-enabling the flyout titlebar and flyout
    ' lock button(s).
    Else
        
        'If we're already not in wide mode, we don't need to move anything!
        If inWideModeNow Then
            
            'Before doing anything else, hide any open flyouts
            UserControls.HideOpenFlyouts 0&
            
            'Reset all disabled/enabled states
            Me.cmdFlyoutLock(0).Visible = True
            Me.cntrPopOut(0).Visible = False
            Me.ttlPanel(0).Enabled = True
            
            'Move the flyout panel off the parent toolpanel
            cntrPopOut(0).SetPosition 0, Me.ScaleHeight + stdPadding
            
            'Restore original caption padding of the slider on the flyout
            Me.sltBrushSetting(1).CaptionPadding = 0
            
            'Reset the position of everything left on the parent toolpanel.
            xOffset = Me.ttlPanel(0).GetLeft + Me.ttlPanel(0).GetWidth + stdPadding
            
            'Continue moving everything back into its original position
            Me.ttlPanel(1).SetLeft xOffset
            Me.sltBrushSetting(2).SetLeft xOffset
            
        End If
            
    End If
    
End Sub

'If you want to set all paintbrush settings at once, use this function
Public Sub SyncAllPaintbrushSettingsToUI()
    Tools_Paint.SetBrushSize sltBrushSetting(0).Value
    Tools_Paint.SetBrushOpacity sltBrushSetting(1).Value
    Tools_Paint.SetBrushHardness sltBrushSetting(2).Value
    Tools_Paint.SetBrushFlow sltBrushSetting(3).Value
    Tools_Paint.SetBrushBlendMode BM_Erase
    Tools_Paint.SetBrushAlphaMode AM_Normal
    If (btsSpacing.ListIndex = 0) Then Tools_Paint.SetBrushSpacing 0! Else Tools_Paint.SetBrushSpacing sldSpacing.Value
    Tools_Paint.SetStrictPixelAlignment chkStrictPixel.Value
End Sub

'If you want to synchronize all UI elements to match current paintbrush settings, use this function
Public Sub SyncUIToAllPaintbrushSettings()
    sltBrushSetting(0).Value = Tools_Paint.GetBrushSize()
    sltBrushSetting(1).Value = Tools_Paint.GetBrushOpacity()
    sltBrushSetting(2).Value = Tools_Paint.GetBrushHardness()
    sltBrushSetting(3).Value = Tools_Paint.GetBrushFlow()
    If (Tools_Paint.GetBrushSpacing() = 0!) Then
        btsSpacing.ListIndex = 0
    Else
        btsSpacing.ListIndex = 1
        sldSpacing.Value = Tools_Paint.GetBrushSpacing()
    End If
    chkStrictPixel.Value = Tools_Paint.GetStrictPixelAlignment()
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
    
    'Assign a tooltip to the "strict pixel alignment" checkbox, as it's somewhat unintuitive.
    chkStrictPixel.AssignTooltip "This setting forcibly aligns paint strokes to pixel grid centerpoints.  This improves precision (especially at small brush sizes), but brush strokes may appear less natural."
    
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
        Tools_Paint.SetBrushSpacing 0!
    Else
        sldSpacing.Visible = True
        Tools_Paint.SetBrushSpacing sldSpacing.Value
    End If
End Sub
