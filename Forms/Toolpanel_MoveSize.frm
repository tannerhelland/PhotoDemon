VERSION 5.00
Begin VB.Form toolpanel_MoveSize 
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
   Begin PhotoDemon.pdButtonStripVertical btsMoveOptions 
      Height          =   1320
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2328
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboLayerResizeQuality 
         Height          =   660
         Left            =   5190
         TabIndex        =   2
         Top             =   60
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1164
         Caption         =   "transform quality"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   9
         Left            =   135
         Top             =   75
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         Caption         =   "layer position (x, y)"
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   10
         Left            =   2655
         Top             =   75
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         Caption         =   "layer size (w, h)"
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   420
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerMove 
         Height          =   570
         Index           =   0
         Left            =   8520
         TabIndex        =   7
         Top             =   420
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   12
         Left            =   8400
         Top             =   60
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   503
         Caption         =   "other options"
      End
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltLayerAngle 
         Height          =   765
         Left            =   120
         TabIndex        =   14
         Top             =   60
         Width           =   4950
         _ExtentX        =   7223
         _ExtentY        =   1349
         Caption         =   "layer angle"
         FontSizeCaption =   10
         Min             =   -360
         Max             =   360
         SigDigits       =   2
      End
      Begin PhotoDemon.pdSlider sltLayerShearX 
         Height          =   765
         Left            =   5400
         TabIndex        =   16
         Top             =   60
         Width           =   4950
         _ExtentX        =   7223
         _ExtentY        =   1349
         Caption         =   "layer shear (x, y)"
         FontSizeCaption =   10
         Min             =   -5
         Max             =   5
         SigDigits       =   2
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAffinePermanent 
         Height          =   570
         Left            =   10800
         TabIndex        =   17
         Top             =   360
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltLayerShearY 
         Height          =   405
         Left            =   5400
         TabIndex        =   18
         Top             =   840
         Width           =   4950
         _ExtentX        =   7223
         _ExtentY        =   714
         Min             =   -5
         Max             =   5
         SigDigits       =   2
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   4
         Left            =   10800
         Top             =   60
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   503
         Caption         =   "other options"
      End
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   0
         Left            =   135
         Top             =   75
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         Caption         =   "interaction options"
      End
      Begin PhotoDemon.pdCheckBox chkAutoActivateLayer 
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "automatically activate layer beneath mouse"
      End
      Begin PhotoDemon.pdCheckBox chkIgnoreTransparent 
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "ignore transparent pixels when auto-activating layers"
      End
      Begin PhotoDemon.pdCheckBox chkLayerBorder 
         Height          =   330
         Left            =   5760
         TabIndex        =   11
         Top             =   360
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show layer borders"
      End
      Begin PhotoDemon.pdCheckBox chkLayerNodes 
         Height          =   330
         Left            =   5760
         TabIndex        =   0
         Top             =   720
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show resize nodes"
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   1
         Left            =   5655
         Top             =   75
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         Caption         =   "display options"
      End
      Begin PhotoDemon.pdCheckBox chkRotateNode 
         Height          =   330
         Left            =   5760
         TabIndex        =   15
         Top             =   1080
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show rotate nodes"
      End
   End
End
Attribute VB_Name = "toolpanel_MoveSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Move/Size Tool Panel
'Copyright 2013-2017 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 06/November/17
'Last update: improve the rate at which things like taskbar icons are refreshed during non-destructive modifications
'
'This form includes all user-editable settings for the Move/Size canvas tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Two sub-panels are available on the "move options" panel
Private Sub btsMoveOptions_Click(ByVal buttonIndex As Long)
    UpdateSubpanel
End Sub

Private Sub UpdateSubpanel()
    Dim i As Long
    For i = 0 To ctlMoveContainer.UBound
        ctlMoveContainer(i).Visible = (i = btsMoveOptions.ListIndex)
    Next i
End Sub

Private Sub cboLayerResizeQuality_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Apply the new quality mode
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerResizeQuality cboLayerResizeQuality.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboLayerResizeQuality_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ResizeQuality, cboLayerResizeQuality.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboLayerResizeQuality_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ResizeQuality, cboLayerResizeQuality.ListIndex
End Sub

'En/disable the "ignore transparent layer bits on click activations" setting if the auto-activate clicked layer setting changes
Private Sub chkAutoActivateLayer_Click()
    chkIgnoreTransparent.Enabled = CBool(chkAutoActivateLayer)
End Sub

'Show/hide layer borders while using the move tool
Private Sub chkLayerBorder_Click()
    ViewportEngine.Stage4_FlipBufferAndDrawUI pdImages(g_CurrentImage), FormMain.mainCanvas(0)
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    ViewportEngine.Stage4_FlipBufferAndDrawUI pdImages(g_CurrentImage), FormMain.mainCanvas(0)
End Sub

Private Sub chkRotateNode_Click()
    ViewportEngine.Stage4_FlipBufferAndDrawUI pdImages(g_CurrentImage), FormMain.mainCanvas(0)
End Sub

Private Sub cmdLayerAffinePermanent_Click()
    If (g_OpenImageCount = 0) Then Exit Sub
    Process "Make layer changes permanent", , BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Layer
End Sub

Private Sub cmdLayerMove_Click(Index As Integer)
    
    If (g_OpenImageCount = 0) Then Exit Sub
    
    Select Case Index
    
        'Make non-destructive resize permanent
        Case 0
            Process "Make layer changes permanent", , BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Layer
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    Tools.SetToolBusyState True
    
    'Initialize move tool panels
    btsMoveOptions.AddItem "size and position", 0
    btsMoveOptions.AddItem "angle and shear", 1
    btsMoveOptions.AddItem "tool settings", 2
    btsMoveOptions.ListIndex = 0
    UpdateSubpanel
    
    'Several reset/apply buttons on this form have very similar purposes
    cmdLayerMove(0).AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    cmdLayerAffinePermanent.AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    
    cboLayerResizeQuality.Clear
    cboLayerResizeQuality.AddItem "Nearest neighbor", 0
    cboLayerResizeQuality.AddItem "Bilinear", 1
    cboLayerResizeQuality.AddItem "Bicubic", 2
    cboLayerResizeQuality.ListIndex = 1
        
    'Load any last-used settings for this form
    'NOTE: this is currently disabled, as all settings on this form are synched to the active layer
    'Set lastUsedSettings = New pdLastUsedSettings
    'lastUsedSettings.SetParentForm Me
    'lastUsedSettings.LoadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub sltLayerAngle_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerAngle sltLayerAngle.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerAngle_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerAngle_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_Angle, sltLayerAngle.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltLayerAngle_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_Angle, sltLayerAngle.Value
End Sub

Private Sub sltLayerShearX_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerShearX sltLayerShearX.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerShearX_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerShearX_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ShearX, sltLayerShearX.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltLayerShearX_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ShearX, sltLayerShearX.Value
End Sub

Private Sub sltLayerShearY_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerShearY sltLayerShearY.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerShearY_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerShearY_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ShearY, sltLayerShearY.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltLayerShearY_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ShearY, sltLayerShearY.Value
End Sub

Private Sub tudLayerMove_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    Select Case Index
    
        'Layer position (x)
        Case 0
            pdImages(g_CurrentImage).GetActiveLayer.SetLayerOffsetX tudLayerMove(Index).Value
        
        'Layer position (y)
        Case 1
            pdImages(g_CurrentImage).GetActiveLayer.SetLayerOffsetY tudLayerMove(Index).Value
        
        'Layer width
        Case 2
            pdImages(g_CurrentImage).GetActiveLayer.SetLayerCanvasXModifier tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False)
            
        'Layer height
        Case 3
            pdImages(g_CurrentImage).GetActiveLayer.SetLayerCanvasYModifier tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False)
        
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)

End Sub

'Non-destructive resizing requires the synchronization of several menus, as well.  Because it's time-consuming to invoke
' SyncInterfaceToCurrentImage, we only call it when the user releases the mouse.
Private Sub tudLayerMove_FinalChange(Index As Integer)
    FinalChangeHandler True
End Sub

Private Sub tudLayerMove_GotFocusAPI(Index As Integer)
    If (g_OpenImageCount = 0) Then Exit Sub
    If (Index = 0) Then
        Processor.FlagInitialNDFXState_Generic pgp_OffsetX, tudLayerMove(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
    ElseIf (Index = 1) Then
        Processor.FlagInitialNDFXState_Generic pgp_OffsetY, tudLayerMove(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
    ElseIf (Index = 2) Then
        Processor.FlagInitialNDFXState_Generic pgp_CanvasXModifier, tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False), pdImages(g_CurrentImage).GetActiveLayerID
    ElseIf (Index = 3) Then
        Processor.FlagInitialNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False), pdImages(g_CurrentImage).GetActiveLayerID
    End If
End Sub

Private Sub tudLayerMove_LostFocusAPI(Index As Integer)
    If (Index = 0) Then
        Processor.FlagFinalNDFXState_Generic pgp_OffsetX, tudLayerMove(Index).Value
    ElseIf (Index = 1) Then
        Processor.FlagFinalNDFXState_Generic pgp_OffsetY, tudLayerMove(Index).Value
    ElseIf (Index = 2) Then
        Processor.FlagFinalNDFXState_Generic pgp_CanvasXModifier, tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False)
    ElseIf (Index = 3) Then
        Processor.FlagFinalNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(Index).Value / pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False)
    End If
End Sub

'When a layer property control (either a slider or spinner) has a "FinalChange" event, use this control to update any
' necessary UI elements that may need to be adjusted due to non-destructive changes.
Private Sub FinalChangeHandler(Optional ByVal performFullUISync As Boolean = False)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.GetToolBusyState
    'If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Redraw the viewport and any relevant UI elements
    'ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    Processor.NDFXUiUpdate
    
    If performFullUISync Then Interface.SyncInterfaceToCurrentImage
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'UI images must be updated against theme-specific colors
    Dim buttonSize As Long
    buttonSize = FixDPI(32)
    cmdLayerMove(0).AssignImage "generic_commit", , buttonSize, buttonSize
    cmdLayerAffinePermanent.AssignImage "generic_commit", , buttonSize, buttonSize
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub
