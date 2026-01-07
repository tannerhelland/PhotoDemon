VERSION 5.00
Begin VB.Form toolpanel_Crop 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12180
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
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButtonToolbox cmdAspectSwap 
      Height          =   345
      Left            =   7320
      TabIndex        =   16
      Top             =   420
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      AutoToggle      =   -1  'True
   End
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
      Height          =   1455
      Index           =   1
      Left            =   4200
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2566
      Begin PhotoDemon.pdDropDown ddPreset 
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1138
         Caption         =   "presets"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsOrientation 
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   23
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   3480
         TabIndex        =   3
         Top             =   960
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   3975
      Index           =   2
      Left            =   8400
      Top             =   960
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5741
      Begin PhotoDemon.pdDropDown ddGuides 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1296
         Caption         =   "guides"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsTarget 
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   90
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1508
         Caption         =   "target"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sldHighlight 
         Height          =   375
         Left            =   1140
         TabIndex        =   19
         Top             =   2325
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   661
         Min             =   1
         Max             =   100
         Value           =   50
         DefaultValue    =   50
      End
      Begin PhotoDemon.pdColorSelector clrHighlight 
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   2325
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         curColor        =   0
         ShowMainWindowColor=   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkHighlight 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "highlight crop area"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   3540
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkDelete 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "delete cropped pixels"
      End
      Begin PhotoDemon.pdCheckBox chkAllowGrowing 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "allow enlarging"
      End
   End
   Begin PhotoDemon.pdSpinner tudCrop 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
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
      TabIndex        =   1
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
      Index           =   1
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
      Index           =   2
      Left            =   9480
      TabIndex        =   7
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      Caption         =   "apply"
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
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   661
      Caption         =   "size (w, h)"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1695
      Index           =   0
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2990
      Begin PhotoDemon.pdSlider sldDPI 
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   900
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         Caption         =   "resolution (PPI)"
         FontSizeCaption =   10
         Min             =   1
         Max             =   1200
         ScaleStyle      =   1
         Value           =   300
         GradientColorRight=   1703935
         NotchPosition   =   2
         NotchValueCustom=   300
      End
      Begin PhotoDemon.pdDropDown ddPreset 
         Height          =   645
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1138
         Caption         =   "presets"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsOrientation 
         Height          =   495
         Index           =   0
         Left            =   2760
         TabIndex        =   26
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   3480
         TabIndex        =   27
         Top             =   1200
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
End
Attribute VB_Name = "toolpanel_Crop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Crop Tool Panel
'Copyright 2024-2026 by Tanner Helland
'Created: 11/November/24
'Last updated: 21/March/25
'Last update: presets are now pulled from a central template manager
'
'This form includes all user-editable settings for the on-canvas Crop tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Prevent synchronization loops
Private m_DontUpdate As Boolean

'Two lists of presets: aspect ratio, and physical sizes
Private m_PresetSize() As PD_SizeTemplate, m_PresetAspect() As PD_SizeTemplate
Private m_numPresetSize As Long, m_numPresetAspect As Long

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub btsOrientation_Click(Index As Integer, ByVal buttonIndex As Long)
    
    If m_DontUpdate Then Exit Sub
    
    'Keep both preset lists in sync (portrait vs landscape)
    m_DontUpdate = True
    
    Dim idxListBackup As Long
    idxListBackup = ddPreset(Index).ListIndex
    
    If (Index = 0) Then
        btsOrientation(1).ListIndex = buttonIndex
    Else
        btsOrientation(0).ListIndex = buttonIndex
    End If
    m_DontUpdate = False
    
    'Change the applied size to match the current orientation
    ApplyNewPreset Index
    
End Sub

Private Sub btsOrientation_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub btsOrientation_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then newTargetHwnd = Me.ddPreset(Index).hWnd Else newTargetHwnd = Me.sldDPI.hWndSlider
    Else
        If shiftTabWasPressed Then newTargetHwnd = Me.ddPreset(Index).hWnd Else newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
    End If
End Sub

Private Sub btsTarget_Click(ByVal buttonIndex As Long)
    Tools_Crop.SetCropAllLayers (btsTarget.ListIndex = 0)
    UpdateEnabledControls
End Sub

Private Sub btsTarget_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub btsTarget_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If Me.cmdCommit(1).Enabled Then
            newTargetHwnd = Me.cmdCommit(1).hWnd
        Else
            newTargetHwnd = Me.ttlPanel(2).hWnd
        End If
    Else
        If Me.chkAllowGrowing.Enabled Then
            newTargetHwnd = Me.chkAllowGrowing.hWnd
        ElseIf Me.chkDelete.Enabled Then
            newTargetHwnd = Me.chkDelete.hWnd
        Else
            newTargetHwnd = Me.chkHighlight.hWnd
        End If
    End If
End Sub

'When toggling the allow/don't allow enlarge setting, we also need to modify min/max values of
' the position and size spin controls.  All handling (including relaying changes back to the crop engine)
' take place in SyncMinMaxAgainstImage.
Private Sub chkAllowGrowing_Click()
    SyncMinMaxAgainstImage
End Sub

Private Sub chkAllowGrowing_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkAllowGrowing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsTarget.hWnd
    Else
        newTargetHwnd = Me.chkDelete.hWnd
    End If
End Sub

Private Sub chkDelete_Click()
    Tools_Crop.SetCropDeletePixels chkDelete.Value
End Sub

Private Sub chkDelete_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkDelete_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkAllowGrowing.hWnd
    Else
        newTargetHwnd = chkHighlight.hWnd
    End If
End Sub

Private Sub chkHighlight_Click()
    Tools_Crop.SetCropHighlight chkHighlight.Value
End Sub

Private Sub chkHighlight_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkHighlight_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If Me.chkDelete.Enabled Then
            newTargetHwnd = Me.chkDelete.hWnd
        ElseIf Me.chkAllowGrowing.Enabled Then
            newTargetHwnd = Me.chkAllowGrowing.hWnd
        Else
            newTargetHwnd = Me.btsTarget.hWnd
        End If
    Else
        newTargetHwnd = clrHighlight.hWnd
    End If
End Sub

Private Sub clrHighlight_ColorChanged()
    Tools_Crop.SetCropHighlightColor clrHighlight.Color
End Sub

Private Sub clrHighlight_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub clrHighlight_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.chkHighlight.hWnd Else newTargetHwnd = Me.sldHighlight.hWndSlider
End Sub

Private Sub cmdAspectSwap_Click(ByVal Shift As ShiftConstants)
    If tudCrop(4).IsValid And tudCrop(5).IsValid Then Tools_Crop.RelayCropChangesFromUI pdd_SwapAspectRatio
End Sub

Private Sub cmdAspectSwap_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub cmdAspectSwap_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = tudCrop(4).hWnd Else newTargetHwnd = tudCrop(5).hWnd
End Sub

Private Sub cmdCommit_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Tools_Crop.CommitCurrentCrop
        Case 1
            Tools_Crop.RemoveCurrentCrop
    End Select
    
End Sub

Private Sub cmdCommit_GotFocusAPI(Index As Integer)
    UpdateFlyout 2, True
End Sub

Private Sub cmdCommit_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then newTargetHwnd = Me.ttlPanel(2).hWnd Else newTargetHwnd = cmdCommit(1).hWnd
    Else
        If shiftTabWasPressed Then newTargetHwnd = Me.cmdCommit(0).hWnd Else newTargetHwnd = Me.btsTarget.hWnd
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
            If shiftTabWasPressed Then newTargetHwnd = Me.sldDPI.hWndSpinner Else newTargetHwnd = Me.ttlPanel(1).hWnd
        Case 1
            If shiftTabWasPressed Then newTargetHwnd = Me.btsOrientation(1).hWnd Else newTargetHwnd = Me.ttlPanel(2).hWnd
        Case 2
            If shiftTabWasPressed Then newTargetHwnd = Me.ddGuides.hWnd Else newTargetHwnd = Me.tudCrop(0).hWnd
    End Select
End Sub

Private Sub cmdLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    SynchronizeLockStates Index
End Sub

Private Sub SynchronizeLockStates(ByVal srcIndex As Long)

    Dim lockedValue As Variant, lockedValue2 As Variant
    If (srcIndex = 0) Then
        lockedValue = tudCrop(2).Value
    ElseIf (srcIndex = 1) Then
        lockedValue = tudCrop(3).Value
    Else
        If (tudCrop(4).Value <> 0#) And (tudCrop(5).Value <> 0#) Then
            lockedValue = tudCrop(4).Value
            lockedValue2 = tudCrop(5).Value
        Else
            lockedValue = 1#
            lockedValue2 = 1#
        End If
    End If

    'When setting a new lock, unlock any other locks.
    If cmdLock(srcIndex).Value Then
        If (srcIndex = 0) Then
            cmdLock(1).Value = False
            cmdLock(2).Value = False
        ElseIf (srcIndex = 1) Then
            cmdLock(0).Value = False
            cmdLock(2).Value = False
        ElseIf (srcIndex = 2) Then
            cmdLock(0).Value = False
            cmdLock(1).Value = False
        End If
    End If
    
    If cmdLock(srcIndex).Value Then
        Tools_Crop.LockProperty srcIndex, lockedValue, lockedValue2
    Else
        Tools_Crop.UnlockProperty srcIndex
    End If
    
End Sub

Private Sub cmdLock_GotFocusAPI(Index As Integer)
    If (Index = 0) Or (Index = 1) Then
        UpdateFlyout 0, True
    ElseIf (Index = 2) Then
        UpdateFlyout 1, True
    End If
End Sub

Private Sub cmdLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then newTargetHwnd = Me.tudCrop(2).hWnd Else newTargetHwnd = Me.tudCrop(3).hWnd
        Case 1
            If shiftTabWasPressed Then newTargetHwnd = Me.tudCrop(3).hWnd Else newTargetHwnd = Me.ddPreset(0).hWnd
        Case 2
            If shiftTabWasPressed Then newTargetHwnd = Me.tudCrop(5).hWnd Else newTargetHwnd = Me.ddPreset(1).hWnd
    End Select
End Sub

Private Sub ddGuides_Click()
    Tools_Crop.SetCropGuide ddGuides.ListIndex
End Sub

Private Sub ddGuides_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sldHighlight.hWndSpinner
    Else
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    End If
End Sub

Private Sub ddPreset_Click(Index As Integer)
    
    'When matching current crop size and aspect ratio to an existing preset, we *don't* want to redraw the screen
    ' (because subtle differences in aspect ratio may trigger an awkward redraw with a slightly different crop rect).
    ' (Also, the m_DontUpdate value is set by the MatchPresetsToCurrentValues() function.)
    If m_DontUpdate Then Exit Sub
    ApplyNewPreset Index
    
End Sub

'When a preset box is clicked (either aspect ratio or size), apply that preset immediately.
Private Sub ApplyNewPreset(ByVal idxTrigger As Long)
    
    'Ignore the top entry of either dropdown (it's just a blank entry for "freeform" behavior)
    Dim idxListbox As Long
    idxListbox = ddPreset(idxTrigger).ListIndex
    If (idxListbox > 0) Then
        
        'Unlock all measurements
        Me.cmdLock(0).Value = False
        Me.cmdLock(1).Value = False
        Me.cmdLock(2).Value = False
        
        'Landscape or portrait orientation can be toggled in the UI; we silently use whichever is selected
        Dim firstValue As Long, secondValue As Long
        
        'Size
        If (idxTrigger = 0) Then
        
            'Failsafe only for presets being initialized correctly
            If (m_numPresetSize >= idxListbox) Then
                
                'Convert the current size from its source unit to pixels, using the user-entered DPI value
                Dim pxSizeH As Long, pxSizeV As Long
                pxSizeH = Units.ConvertOtherUnitToPixels(m_PresetSize(idxListbox).thisSizeUnit, m_PresetSize(idxListbox).hSizeOrig, Me.sldDPI.Value)
                pxSizeV = Units.ConvertOtherUnitToPixels(m_PresetSize(idxListbox).thisSizeUnit, m_PresetSize(idxListbox).vSizeOrig, Me.sldDPI.Value)
                
                'Switch between landscape and portrait (based on the relevant toggle in the UI)
                If (btsOrientation(idxTrigger).ListIndex = 0) Then
                    firstValue = PDMath.Max2Int(pxSizeH, pxSizeV)
                    secondValue = PDMath.Min2Int(pxSizeH, pxSizeV)
                Else
                    secondValue = PDMath.Max2Int(pxSizeH, pxSizeV)
                    firstValue = PDMath.Min2Int(pxSizeH, pxSizeV)
                End If
                
                'Assign the value; the _Change event of the controls will relay the change to the crop engine
                Me.tudCrop(2).Value = firstValue
                Me.tudCrop(3).Value = secondValue
                
            End If
        
        'Aspect ratio
        Else
            
            'Failsafe only for presets being initialized correctly
            If (m_numPresetAspect >= idxListbox) Then
                
                'Aspect ratio doesn't require unit handling
                If (btsOrientation(idxTrigger).ListIndex = 0) Then
                    firstValue = PDMath.Max2Int(m_PresetAspect(idxListbox).hSizeOrig, m_PresetAspect(idxListbox).vSizeOrig)
                    secondValue = PDMath.Min2Int(m_PresetAspect(idxListbox).hSizeOrig, m_PresetAspect(idxListbox).vSizeOrig)
                Else
                    secondValue = PDMath.Max2Int(m_PresetAspect(idxListbox).hSizeOrig, m_PresetAspect(idxListbox).vSizeOrig)
                    firstValue = PDMath.Min2Int(m_PresetAspect(idxListbox).hSizeOrig, m_PresetAspect(idxListbox).vSizeOrig)
                End If
                
                'We don't want the aspect ratio relayed until *both* text boxes are set
                Tools.SetToolBusyState True
                Me.tudCrop(4).Value = firstValue
                Me.tudCrop(5).Value = secondValue
                Tools.SetToolBusyState False
                Tools_Crop.RelayCropChangesFromUI pdd_AspectBoth, Me.tudCrop(4).Value, Me.tudCrop(5).Value
                
                'Ensure the viewport was redrawn after the second relay event
                Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                
            End If
            
        End If
        
        'pdButtonToolbox controls do not auto-synchronize on assignment via code (by design),
        ' so manually sync lock states with the crop module now
        SynchronizeLockStates 0
        SynchronizeLockStates 1
        SynchronizeLockStates 2
        
    End If
    
End Sub

Private Sub ddPreset_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub ddPreset_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then newTargetHwnd = Me.cmdLock(1).hWnd Else newTargetHwnd = btsOrientation(0).hWnd
    ElseIf (Index = 1) Then
        If shiftTabWasPressed Then newTargetHwnd = Me.cmdLock(2).hWnd Else newTargetHwnd = btsOrientation(1).hWnd
    End If
End Sub

Private Sub Form_Activate()
    
    'Only enable commit/clear buttons while a crop is actually active
    Me.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    Me.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
End Sub

Private Sub Form_Load()
    
    Tools.SetToolBusyState True
    
    'Populate any run-time UI elements
    btsTarget.AddItem "image", 0
    btsTarget.AddItem "layer", 1
    btsTarget.ListIndex = 0
    
    btsOrientation(0).AddItem vbNullString, 0
    btsOrientation(0).AddItem vbNullString, 1
    btsOrientation(0).ListIndex = 0
    
    btsOrientation(1).AddItem vbNullString, 0
    btsOrientation(1).AddItem vbNullString, 1
    btsOrientation(1).ListIndex = 0
    
    ddGuides.Clear
    ddGuides.AddItem "none", 0
    ddGuides.AddItem "centerlines", 1
    ddGuides.AddItem "rule of thirds", 2
    ddGuides.AddItem "rule of fifths", 3
    ddGuides.AddItem "golden ratio", 4
    ddGuides.AddItem "diagonals", 5
    ddGuides.ListIndex = 0
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues True
    
    'Because last-used settings may update crop render settings, immediately relay any changes to the crop tool
    Tools_Crop.SetCropAllowEnlarge Me.chkAllowGrowing.Value
    Tools_Crop.SetCropDeletePixels Me.chkDelete.Value
    Tools_Crop.SetCropHighlight chkHighlight.Value
    Tools_Crop.SetCropHighlightColor clrHighlight.Color
    Tools_Crop.SetCropHighlightOpacity sldHighlight.Value
    Tools_Crop.SetCropGuide ddGuides.ListIndex
    
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

'Non-measurement settings are stored between sessions
Private Sub m_LastUsedSettings_AddCustomPresetData()
    
    With m_lastUsedSettings
        .AddPresetData "crop-tool-allow-enlarge", Trim$(Str$(Me.chkAllowGrowing.Value))
        .AddPresetData "crop-tool-delete-pixels", Trim$(Str$(Me.chkDelete.Value))
        .AddPresetData "crop-tool-highlight", Trim$(Str$(Me.chkHighlight.Value))
        .AddPresetData "crop-tool-guides", Trim$(Str$(Me.ddGuides.ListIndex))
        .AddPresetData "crop-tool-target-image", Trim$(Str$((Me.btsTarget.ListIndex = 0)))
    End With

End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

    Const STR_FALSE As String = "False", STR_TRUE As String = "True"
    With m_lastUsedSettings
        Me.chkAllowGrowing.Value = Strings.StringsEqual(.RetrievePresetData("crop-tool-allow-enlarge", STR_FALSE), STR_TRUE, True)
        Tools_Crop.SetCropAllowEnlarge chkAllowGrowing.Value
        Me.chkDelete.Value = Strings.StringsEqual(.RetrievePresetData("crop-tool-delete-pixels", STR_TRUE), STR_TRUE, True)
        Tools_Crop.SetCropDeletePixels chkDelete.Value
        Me.chkHighlight.Value = Strings.StringsEqual(.RetrievePresetData("crop-tool-highlight", STR_TRUE), STR_TRUE, True)
        Tools_Crop.SetCropHighlight chkHighlight.Value
        Tools_Crop.SetCropHighlightColor Me.clrHighlight.Color
        Tools_Crop.SetCropHighlightOpacity Me.sldHighlight.Value
        Me.ddGuides.ListIndex = CLng(.RetrievePresetData("crop-tool-guides", "0"))
        Tools_Crop.SetCropGuide Me.ddGuides.ListIndex
        If Strings.StringsEqual(.RetrievePresetData("crop-tool-target-image", STR_TRUE), STR_TRUE, True) Then
            Me.btsTarget.ListIndex = 0
        Else
            Me.btsTarget.ListIndex = 1
        End If
        Tools_Crop.SetCropAllLayers (Me.btsTarget.ListIndex = 0)
        UpdateEnabledControls
    End With
    
End Sub

Private Sub sldDPI_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.btsOrientation(0).hWnd Else newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
End Sub

Private Sub sldHighlight_Change()
    Tools_Crop.SetCropHighlightOpacity sldHighlight.Value
End Sub

Private Sub sldHighlight_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub sldHighlight_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.clrHighlight.hWnd Else newTargetHwnd = Me.ddGuides.hWnd
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = tudCrop(1).hWnd
            Case 1
                newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
            Case 2
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
        End Select
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.tudCrop(2).hWnd
            Case 1
                newTargetHwnd = Me.tudCrop(4).hWnd
            Case 2
                If Me.cmdCommit(0).Enabled Then
                    newTargetHwnd = Me.cmdCommit(0).hWnd
                Else
                    newTargetHwnd = Me.btsTarget.hWnd
                End If
        End Select
    End If
End Sub

Private Sub tudCrop_Change(Index As Integer)
    
    'When the user changes size and/or aspect ratio, auto-match the corresponding dropdown
    ' (same needs to be done for *both* size and aspect ratio).  This happens even if changes
    ' are suspended to prevent recursion (the matching function has its own protections against this).
    MatchPresetsToCurrentValues
    
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
                
                'If either of the width or height is locked, the user wants to retain that dimension - so we need
                ' to calculate the *opposite* dimension and modify only that.
                If (cmdLock(0).Value Or cmdLock(1).Value) Then
                    
                    If ((Index = 4) And (Not cmdLock(0).Value)) Or cmdLock(1).Value Then
                        aspFinal = aspW / aspH
                        newDimension = tudCrop(3).Value * aspFinal
                        Tools_Crop.RelayCropChangesFromUI pdd_AspectRatioW, newDimension, tudCrop(4).Value, tudCrop(5).Value
                    Else
                        aspFinal = aspH / aspW
                        newDimension = tudCrop(2).Value * aspFinal
                        Tools_Crop.RelayCropChangesFromUI pdd_AspectRatioH, newDimension, tudCrop(4).Value, tudCrop(5).Value
                    End If
                    
                'If neither width nor height is locked, it doesn't matter which one we modify - we simply want
                ' to preserve the current aspect ratio (as a fraction), which can be tricky if the crop can't
                ' easily be kept in-bounds.
                Else
                    Tools_Crop.RelayCropChangesFromUI pdd_AspectBoth, aspW, aspH
                End If
                
            End If
            
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'When the user changes size and/or aspect ratio, we want to iterate all presets and match up the current
' size/apsect ratio settings to a given preset (if one matches).  This is also important for clearing
' a previous preset if the user changes sizes manually to a different value.
Private Sub MatchPresetsToCurrentValues()
    
    'Failsafe only
    If (m_numPresetSize < 2) Then Exit Sub
    
    If m_DontUpdate Then Exit Sub
    m_DontUpdate = True
    
    'Start with size matching
    Dim firstSize As Long, secondSize As Long
    firstSize = tudCrop(2).Value
    secondSize = tudCrop(3).Value
    
    Dim curResolution As Long
    curResolution = Me.sldDPI.Value
    
    'Skip the first entry in the list, because it's blank (it represents "no preset match").
    Dim i As Long, idxTarget As Long
    idxTarget = 0
    
    For i = 1 To m_numPresetSize - 1
        
        Dim pxWidth As Long, pxHeight As Long
        
        'As relevant, convert the source size from real-world units to pixels
        If (m_PresetSize(i).thisSizeUnit = mu_Pixels) Then
            pxWidth = m_PresetSize(i).hSizeOrig
            pxHeight = m_PresetSize(i).vSizeOrig
        Else
            pxWidth = Units.ConvertOtherUnitToPixels(m_PresetSize(i).thisSizeUnit, m_PresetSize(i).hSizeOrig, curResolution)
            pxHeight = Units.ConvertOtherUnitToPixels(m_PresetSize(i).thisSizeUnit, m_PresetSize(i).vSizeOrig, curResolution)
        End If
        
        'Compare in both landscape *and* portrait orientation
        If (firstSize = pxWidth) And (secondSize = pxHeight) Then
            idxTarget = i
            Exit For
        ElseIf (firstSize = pxHeight) And (secondSize = pxWidth) Then
            idxTarget = i
            Exit For
        End If
        
    Next i
    
    If (ddPreset(0).ListIndex <> idxTarget) Then ddPreset(0).ListIndex = idxTarget
    
    'With size handled, repeat the above steps for aspect ratio.
    ' (Aspect ratio is simpler because we don't need to account for unit conversion.)
    firstSize = tudCrop(4).Value
    secondSize = tudCrop(5).Value
    
    idxTarget = 0
    For i = 1 To m_numPresetAspect - 1
        
        'As relevant, convert the source size from real-world units to pixels
        pxWidth = m_PresetAspect(i).hSizeOrig
        pxHeight = m_PresetAspect(i).vSizeOrig
        
        'Compare in both landscape *and* portrait orientation
        If (firstSize = pxWidth) And (secondSize = pxHeight) Then
            idxTarget = i
            Exit For
        ElseIf (firstSize = pxHeight) And (secondSize = pxWidth) Then
            idxTarget = i
            Exit For
        End If
        
    Next i
    
    If (ddPreset(1).ListIndex <> idxTarget) Then ddPreset(1).ListIndex = idxTarget
    m_DontUpdate = False
    
End Sub

Private Sub tudCrop_GotFocusAPI(Index As Integer)
    If (Index = 2) Or (Index = 3) Then
        UpdateFlyout 0, True
    ElseIf (Index = 4) Or (Index = 5) Then
        UpdateFlyout 1, True
    End If
End Sub

Private Sub tudCrop_LostFocusAPI(Index As Integer)
    If (Not PDImages.IsImageActive()) Then Exit Sub
End Sub

'Because these controls are laid out in a non-standard pattern, we want to manually specify tab and
' shift+tab focus targets.
Private Sub tudCrop_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = cmdFlyoutLock(2).hWnd
            Case 1
                newTargetHwnd = tudCrop(0).hWnd
            Case 2
                newTargetHwnd = ttlPanel(0).hWnd
            Case 3
                newTargetHwnd = cmdLock(0).hWnd
            Case 4
                newTargetHwnd = ttlPanel(1).hWnd
            Case 5
                newTargetHwnd = cmdAspectSwap.hWnd
        End Select
    Else
        Select Case Index
            Case 0
                newTargetHwnd = tudCrop(1).hWnd
            Case 1
                newTargetHwnd = ttlPanel(0).hWnd
            Case 2
                newTargetHwnd = cmdLock(0).hWnd
            Case 3
                newTargetHwnd = cmdLock(1).hWnd
            Case 4
                newTargetHwnd = cmdAspectSwap.hWnd
            Case 5
                newTargetHwnd = cmdLock(2).hWnd
        End Select
    End If

End Sub

'When the active image changes (or the "allow growing" toggle changes), we should adjust max/min values
' of spin controls to prevent the user from
Private Sub SyncMinMaxAgainstImage()
    
    Const MAX_SPIN_VALUE As Long = 65000
    
    'To make this behavior more intuitive, max/min of spin controls should change to match.
    
    'Set minimums (only applies to position)...
    If chkAllowGrowing.Value Or (Not PDImages.IsImageActive()) Then
        tudCrop(0).Min = -1& * MAX_SPIN_VALUE
        tudCrop(1).Min = tudCrop(0).Min
    Else
        tudCrop(0).Min = 0
        tudCrop(1).Min = 0
    End If
    
    '...and maximums (applies to position *and* size)
    If chkAllowGrowing.Value Or (Not PDImages.IsImageActive()) Then
        tudCrop(0).Max = MAX_SPIN_VALUE
        tudCrop(1).Max = MAX_SPIN_VALUE
        tudCrop(2).Max = MAX_SPIN_VALUE
        tudCrop(3).Max = MAX_SPIN_VALUE
    Else
        
        'Redundant failsafe for image existence
        If PDImages.IsImageActive Then
            
            'Position (x/y)
            tudCrop(0).Max = PDImages.GetActiveImage.Width - 1
            tudCrop(1).Max = PDImages.GetActiveImage.Height - 1
            
            'Size (w/h)
            tudCrop(2).Max = PDImages.GetActiveImage.Width
            tudCrop(3).Max = PDImages.GetActiveImage.Height
            
        End If
        
    End If
    
    'Relay any changes to the actual crop engine
    Dim cropChanged As Boolean
    If chkAllowGrowing.Value Then
        Tools_Crop.SetCropAllowEnlarge True
        cropChanged = Tools_Crop.NotifyCropMaxSizes(0, 0)
    Else
        cropChanged = Tools_Crop.NotifyCropMaxSizes(tudCrop(2).Max, tudCrop(3).Max)
        Tools_Crop.SetCropAllowEnlarge False
    End If
    
    If PDImages.IsImageActive() Then
        If (Tools_Crop.IsValidCropActive() Or cropChanged) Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
    End If
    
End Sub

'This panel *does* need to be notified of active image changes, because things like max/min of spin controls
' may change based on user settings and image dimensions.
Public Sub NotifyActiveImageChanged()
    SyncMinMaxAgainstImage
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Lock/unlock buttons are standardized across *all* toolpanels
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(16)
    
    Dim i As Long
    For i = 0 To cmdLock.Count - 1
        cmdLock(i).AssignImage "generic_unlock", , buttonSize, buttonSize
        cmdLock(i).AssignImage_Pressed "generic_lock", , buttonSize, buttonSize
        cmdLock(i).AssignTooltip "Lock this value.  (Only one value can be locked at a time.  If you lock a new value, previously locked values will unlock.)"
    Next i
    
    'Commit and cancel buttons use generic ok/cancel images
    cmdCommit(0).AssignImage "generic_ok", , buttonSize, buttonSize
    cmdCommit(1).AssignImage "generic_cancel", , buttonSize, buttonSize
    
    buttonSize = Interface.FixDPI(14)
    cmdAspectSwap.AssignImage "edit_repeat", , buttonSize, buttonSize
    cmdAspectSwap.AssignTooltip "Swap width and height"
    
    'Orientation toggle uses images only
    Dim imgOrientationColor As Long
    If (Not g_Themer Is Nothing) Then imgOrientationColor = g_Themer.GetGenericUIColor(UI_GrayDark)
    buttonSize = Interface.FixDPI(18)
    
    Dim tmpDIB As pdDIB
    IconsAndCursors.LoadResourceToDIB "generic_image", tmpDIB, buttonSize, buttonSize, 0&, imgOrientationColor, False, GP_IM_HighQualityBicubic
    btsOrientation(0).AssignImageToItem 0, vbNullString, tmpDIB
    btsOrientation(1).AssignImageToItem 0, vbNullString, tmpDIB
    IconsAndCursors.LoadResourceToDIB "generic_imageportrait", tmpDIB, buttonSize, buttonSize, 0&, imgOrientationColor, False, GP_IM_HighQualityBicubic
    btsOrientation(0).AssignImageToItem 1, vbNullString, tmpDIB
    btsOrientation(1).AssignImageToItem 1, vbNullString, tmpDIB
    
    'Build the first copy of preset lists for both size and aspect ratio
    UpdatePresetLists
    
    'Next, apply localized tooltips to any other UI items that require it
    chkAllowGrowing.AssignTooltip "Allow cropping outside image boundaries (which will enlarge the image)."
    sldDPI.AssignTooltip "Resolution is used to convert real-world measurements to pixels.  It is ignored for measurements already in pixels."
    
    Dim ttOrientation As String
    If (Not g_Language Is Nothing) Then ttOrientation = g_Language.TranslateMessage("You can use this measurement in landscape or portrait orientation.")
    btsOrientation(0).AssignTooltip ttOrientation
    btsOrientation(1).AssignTooltip ttOrientation
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub UpdateEnabledControls()
    If (Me.btsTarget.ListIndex = 0) Then
        Me.chkAllowGrowing.Enabled = True
        Me.chkDelete.Enabled = True
        Tools_Crop.SetCropDeletePixels Me.chkDelete.Value
    Else
        Me.chkAllowGrowing.Value = True
        Me.chkAllowGrowing.Enabled = False
        Me.chkDelete.Value = True
        Tools_Crop.SetCropDeletePixels True
        Me.chkDelete.Enabled = False
    End If
End Sub

'Update the size and aspect ratio preset lists
Private Sub UpdatePresetLists()
    
    'Reset preset and aspect ratio collections
    Const INIT_PRESET_SIZE As Long = 8
    If (m_numPresetSize = 0) Then ReDim m_PresetSize(0 To INIT_PRESET_SIZE - 1) As PD_SizeTemplate
    If (m_numPresetAspect = 0) Then ReDim m_PresetAspect(0 To INIT_PRESET_SIZE - 1) As PD_SizeTemplate
    m_numPresetSize = 0
    m_numPresetAspect = 0
    
    'Clear existing preset lists
    ddPreset(0).Clear
    ddPreset(1).Clear
    
    'Add blank presets for size and aspect ratio.  (These are used for "free-form" or "custom" values.)
    AddPreset 0, 0, 0
    AddPreset 1, 0, 0
    
    'Grab copies of the central size and aspect ratio collections
    Dim tmpPresets() As PD_SizeTemplate, tmpAspects() As PD_SizeTemplate
    If Units.GetCopyOfSizeAndAspectTemplates(tmpPresets, tmpAspects) Then
        
        'Add all presets to our collections.  Note that separators are added when switching between units
        ' (e.g. "1920x1080" followed by "8.5x11 US Letter" will auto-insert a separator)
        Dim i As Long, separatorNow As Boolean
        For i = 0 To UBound(tmpPresets)
            separatorNow = False
            If (i < UBound(tmpPresets)) Then separatorNow = (tmpPresets(i).thisSizeUnit <> tmpPresets(i + 1).thisSizeUnit)
            AddPreset 0, tmpPresets(i).hSizeOrig, tmpPresets(i).vSizeOrig, tmpPresets(i).localizedName, tmpPresets(i).thisSizeUnit, separatorNow
        Next i
        
        For i = 0 To UBound(tmpAspects)
            AddPreset 1, tmpAspects(i).hSizeOrig, tmpAspects(i).vSizeOrig
        Next i
        
    End If
    
End Sub

Private Sub AddPreset(ByVal idxDropdown As Long, ByVal szHorizontal As Single, ByVal szVertical As Single, Optional ByVal szName As String = vbNullString, Optional ByVal UnitOfMeasurement As PD_MeasurementUnit = mu_Pixels, Optional ByVal showDividerAfter As Boolean = False)
    
    'This function must construct a UI string for this entry
    Dim finalSize As String
    
    'A 0-size indicates "free" or "custom" scaling, so ignore any further processing
    If (szHorizontal > 0!) Then
        
        'Store this preset (in pixel measurements only) at module-level, so we can access it on dropdown clicks
        If (idxDropdown = 0) Then
            If (m_numPresetSize > UBound(m_PresetSize)) Then ReDim Preserve m_PresetSize(0 To m_numPresetSize * 2 - 1) As PD_SizeTemplate
            m_PresetSize(m_numPresetSize).hSizeOrig = szHorizontal
            m_PresetSize(m_numPresetSize).vSizeOrig = szVertical
            m_PresetSize(m_numPresetSize).thisSizeUnit = UnitOfMeasurement
            m_numPresetSize = m_numPresetSize + 1
        Else
            If (m_numPresetAspect > UBound(m_PresetAspect)) Then ReDim Preserve m_PresetAspect(0 To m_numPresetAspect * 2 - 1) As PD_SizeTemplate
            m_PresetAspect(m_numPresetAspect).hSizeOrig = szHorizontal
            m_PresetAspect(m_numPresetAspect).vSizeOrig = szVertical
            m_PresetAspect(m_numPresetAspect).thisSizeUnit = mu_Pixels
            m_numPresetAspect = m_numPresetAspect + 1
        End If
        
        'Convert that to a consistently formatted string for the given unit
        finalSize = Units.GetValueFormattedForUnit(UnitOfMeasurement, szHorizontal, False) & " x " & Units.GetValueFormattedForUnit(UnitOfMeasurement, szVertical, (idxDropdown = 0))
        
    Else
        If (idxDropdown = 0) Then
            m_numPresetSize = m_numPresetSize + 1
        Else
            m_numPresetAspect = m_numPresetAspect + 1
        End If
        finalSize = vbNullString
    End If
            
    If (LenB(finalSize) > 0) Then
        
        'Append the given name, if any
        If (LenB(szName) <> 0) Then finalSize = finalSize & "  (" & szName & ")"
        Me.ddPreset(idxDropdown).AddItem finalSize, -1, showDividerAfter
        
    Else
        Me.ddPreset(idxDropdown).AddItem vbNullString, 0, True
    End If
    
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
