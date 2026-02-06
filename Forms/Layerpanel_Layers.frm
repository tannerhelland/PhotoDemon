VERSION 5.00
Begin VB.Form layerpanel_Layers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Layerpanel_Layers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdLayerList lstLayers 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4048
   End
   Begin PhotoDemon.pdContainer ctlGroupLayerButtons 
      Height          =   525
      Left            =   0
      Top             =   6720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   926
      Begin PhotoDemon.pdButtonToolbox cmdLayerAction 
         Height          =   510
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAction 
         Height          =   510
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAction 
         Height          =   510
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAction 
         Height          =   510
         Index           =   3
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAction 
         Height          =   510
         Index           =   4
         Left            =   2400
         TabIndex        =   9
         Top             =   0
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         AutoToggle      =   -1  'True
      End
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   360
      Left            =   945
      TabIndex        =   0
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdTextBox txtLayerName 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
   End
   Begin PhotoDemon.pdLabel lblLayerSettings 
      Height          =   240
      Index           =   0
      Left            =   0
      Top             =   120
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   423
      Caption         =   "opacity:"
      Layout          =   2
   End
   Begin PhotoDemon.pdSlider sltLayerOpacity 
      CausesValidation=   0   'False
      Height          =   405
      Left            =   960
      TabIndex        =   2
      Top             =   30
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   53
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdLabel lblLayerSettings 
      Height          =   240
      Index           =   1
      Left            =   0
      Top             =   540
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   423
      Caption         =   "blend:"
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblLayerSettings 
      Height          =   240
      Index           =   2
      Left            =   0
      Top             =   960
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   423
      Caption         =   "alpha:"
      Layout          =   2
   End
   Begin PhotoDemon.pdDropDown cboAlphaMode 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   900
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
   End
End
Attribute VB_Name = "layerpanel_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layer Tool Panel
'Copyright 2014-2026 by Tanner Helland
'Created: 25/March/14
'Last updated: 25/September/15
'Last update: split into its own subpanel, so we can stick more cool stuff on the right panel.
'
'As part of the 7.0 release, PD's right-side panel gained a lot of new functionality.  To simplify the code for
' the new panel, each chunk of related settings (e.g. layer, nav, color selector) was moved to its own subpanel.
'
'This form is the subpanel for layer settings.
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

'Layer buttons are more easily referenced by this enum rather than their actual indices
Private Enum LAYER_BUTTON_ID
    LYR_BTN_ADD = 0
    LYR_BTN_DELETE = 1
    LYR_BTN_MOVE_UP = 2
    LYR_BTN_MOVE_DOWN = 3
    LYR_BTN_DUPLICATE = 4
End Enum

#If False Then
    Private Const LYR_BTN_ADD = 0, LYR_BTN_DELETE = 1
    Private Const LYR_BTN_MOVE_UP = 2, LYR_BTN_MOVE_DOWN = 4, LYR_BTN_DUPLICATE = 4
#End If

'Sometimes we need to make changes that will raise redraw-causing events.  Set this variable to TRUE if you want
' such functions to ignore their automatic redrawing.
Private m_DisableRedraws As Boolean

'To prevent unnecessary redraws, we check for repeat calls and ignore accordingly
Private m_WidthAtLastResize As Long, m_HeightAtLastResize As Long

'External functions can force a full redraw by calling this sub.
' (This is necessary whenever layers are added, deleted, re-ordered, etc.)
Public Sub ForceRedraw(Optional ByVal refreshThumbnailCache As Boolean = True, Optional ByVal layerID As Long = -1)
    
    'Sync opacity, blend mode, and other controls to the currently active layer
    m_DisableRedraws = True
    If PDImages.IsImageActive() Then
        If (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            
            With PDImages.GetActiveImage.GetActiveLayer
                
                'Synchronize the opacity scroll bar to the active layer
                If (sltLayerOpacity.Value <> .GetLayerOpacity) Then sltLayerOpacity.Value = .GetLayerOpacity
                
                'Synchronize the blend and alpha modes to the active layer
                If (cboBlendMode.ListIndex <> .GetLayerBlendMode) Then cboBlendMode.ListIndex = .GetLayerBlendMode
                If (cboAlphaMode.ListIndex <> .GetLayerAlphaMode) Then cboAlphaMode.ListIndex = .GetLayerAlphaMode
            
            End With
        
        End If
    End If
    
    m_DisableRedraws = False
    
    'Notify the layer box of the redraw request
    lstLayers.RequestRedraw refreshThumbnailCache, layerID
    
    'Determine which buttons need to be activated.
    CheckButtonEnablement
    
End Sub

'Whenever a layer is activated, we must re-determine which buttons the user has access to.  Move up/down are disabled for
' entries at either end, and the last layer of an image cannot be deleted.
Private Sub CheckButtonEnablement()
    
    'Make sure at least one image has been loaded
    If PDImages.IsImageActive() Then

        'Add and Dupliate layer are always allowed
        cmdLayerAction(LYR_BTN_ADD).Enabled = True
        cmdLayerAction(LYR_BTN_DUPLICATE).Enabled = True
        
        'Merge down is only allowed for layer indexes > 0
        cmdLayerAction(LYR_BTN_MOVE_DOWN).Enabled = (PDImages.GetActiveImage.GetActiveLayerIndex > 0)
        
        'Merge up is only allowed for layer indexes < NUM_OF_LAYERS
        cmdLayerAction(LYR_BTN_MOVE_UP).Enabled = (PDImages.GetActiveImage.GetActiveLayerIndex < PDImages.GetActiveImage.GetNumOfLayers - 1)
        
        'Delete layer is only allowed if there are multiple layers present
        cmdLayerAction(LYR_BTN_DELETE).Enabled = (PDImages.GetActiveImage.GetNumOfLayers > 1)
        
    'If no images are loaded, disable all layer action buttons
    Else
    
        Dim i As Long
        For i = cmdLayerAction.lBound To cmdLayerAction.UBound
            cmdLayerAction(i).Enabled = False
        Next i
        
    End If
    
End Sub

'Change the alpha mode of the active layer
Private Sub cboAlphaMode_Click()

    'By default, changing the drop-down will automatically update the alpha mode of the selected layer, and the main viewport
    ' will be redrawn.  When changing the alpha mode programmatically, set m_DisableRedraws to TRUE to prevent cylical redraws.
    If m_DisableRedraws Then Exit Sub

    If PDImages.IsImageActive() Then
        If (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            PDImages.GetActiveImage.GetActiveLayer.SetLayerAlphaMode cboAlphaMode.ListIndex
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
    End If

End Sub

Private Sub cboAlphaMode_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_AlphaMode, cboAlphaMode.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboAlphaMode_LostFocusAPI()
    If Tools.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Generic pgp_AlphaMode, cboAlphaMode.ListIndex
End Sub

'Change the blend mode of the active layer
Private Sub cboBlendMode_Click()

    'By default, changing the drop-down will automatically update the blend mode of the selected layer, and the main viewport
    ' will be redrawn.  When changing the blend mode programmatically, set m_DisableRedraws to TRUE to prevent cylical redraws.
    If m_DisableRedraws Then Exit Sub

    If PDImages.IsImageActive() Then
        If (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            PDImages.GetActiveImage.GetActiveLayer.SetLayerBlendMode cboBlendMode.ListIndex
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
    End If

End Sub

Private Sub cboBlendMode_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_BlendMode, cboBlendMode.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboBlendMode_LostFocusAPI()
    If Tools.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Generic pgp_BlendMode, cboBlendMode.ListIndex
End Sub

'Layer action buttons - move layers up/down, delete layers, etc.
Private Sub cmdLayerAction_Click(Index As Integer, ByVal Shift As ShiftConstants)

    Select Case Index
    
        Case LYR_BTN_ADD
            If (Shift = vbShiftMask) Then
                Actions.LaunchAction_ByName "layer_addblank"
            Else
                Actions.LaunchAction_ByName "layer_addbasic"
            End If
            
        Case LYR_BTN_DELETE
            Actions.LaunchAction_ByName "layer_deletecurrent"
            
        Case LYR_BTN_MOVE_UP
            If (Shift = vbShiftMask) Then
                Actions.LaunchAction_ByName "layer_mergeup"
            Else
                Actions.LaunchAction_ByName "layer_moveup"
            End If
            
        Case LYR_BTN_MOVE_DOWN
            If (Shift = vbShiftMask) Then
                Actions.LaunchAction_ByName "layer_mergedown"
            Else
                Actions.LaunchAction_ByName "layer_movedown"
            End If
            
        Case LYR_BTN_DUPLICATE
            Actions.LaunchAction_ByName "layer_duplicate"
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    m_DisableRedraws = True
    
    'Populate the alpha and blend mode boxes
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    Interface.PopulateAlphaModeDropDown cboAlphaMode, AM_Normal
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    Me.UpdateAgainstCurrentTheme
    
    m_DisableRedraws = False
    
End Sub

Private Sub Form_Resize()
    If (Me.ScaleWidth <> m_WidthAtLastResize) Or (Me.ScaleHeight <> m_HeightAtLastResize) Then ReflowInterface
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If

End Sub

'Change the opacity of the current layer
Private Sub sltLayerOpacity_Change()

    'By default, changing the scroll bar will automatically update the opacity value of the selected layer, and
    ' the main viewport will be redrawn.  When changing the scrollbar programmatically, set m_DisableRedraws to TRUE
    ' to prevent cylical redraws.
    If m_DisableRedraws Then Exit Sub

    If PDImages.IsImageActive() Then
        If Not (PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            PDImages.GetActiveImage.GetActiveLayer.SetLayerOpacity sltLayerOpacity.Value
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
    End If

End Sub

Private Sub sltLayerOpacity_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_Opacity, sltLayerOpacity.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltLayerOpacity_LostFocusAPI()
    If Tools.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Generic pgp_Opacity, sltLayerOpacity.Value
End Sub

'Whenever the layer toolbox is resized, we must reflow all objects to fill the available space.  Note that we do not do
' specialized handling for the vertical direction; vertically, the only change we handle is resizing the layer box itself
' to fill whatever vertical space is available.
Private Sub ReflowInterface()
    
    Dim curFormWidth As Long, curFormHeight As Long
    If (g_WindowManager Is Nothing) Then
        curFormWidth = Me.ScaleWidth
        curFormHeight = Me.ScaleHeight
    Else
        curFormWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        curFormHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    End If
    
    m_WidthAtLastResize = curFormWidth
    m_HeightAtLastResize = curFormHeight
    
    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available horizontal and vertical space.
    
    'This value will be used to check for minimizing.  If the window is going down, we do not want to attempt a resize!
    Dim sizeCheck As Long
    
    'Start by moving the button box to the bottom of the available area
    Dim bottomPadding As Long
    bottomPadding = Interface.FixDPI(2)
    
    sizeCheck = curFormHeight - ctlGroupLayerButtons.GetHeight - bottomPadding
    If (sizeCheck > 0) Then ctlGroupLayerButtons.SetTop sizeCheck Else Exit Sub
    
    'Next, stretch the layer box to fill the available space
    sizeCheck = (ctlGroupLayerButtons.GetTop - lstLayers.GetTop) - bottomPadding
    If (sizeCheck > 0) Then
            
        If (lstLayers.GetHeight <> sizeCheck) Then lstLayers.SetHeight sizeCheck
        
        'Vertical resizing has now been covered successfully.  Time to handle horizontal resizing.
        
        'Left-align the opacity, blend and alpha mode controls against their respective labels.
        sltLayerOpacity.SetLeft lblLayerSettings(0).GetLeft + lblLayerSettings(0).GetWidth + Interface.FixDPI(4)
        cboBlendMode.SetLeft lblLayerSettings(1).GetLeft + lblLayerSettings(1).GetWidth + Interface.FixDPI(8)
        
        'So this is kind of funny, but in English, the "blend mode" and "alpha mode" layers are offset
        ' by 1 px due to the different pixel lengths of the "blend" and "alpha" labels.  To make them
        ' look a bit prettier, we manually pad the non-translated version.
        Dim alphaOffset As Long
        If (Not g_Language Is Nothing) Then
            If g_Language.TranslationActive Then alphaOffset = 8 Else alphaOffset = 9
        Else
            alphaOffset = 9
        End If
        
        cboAlphaMode.SetLeft lblLayerSettings(2).GetLeft + lblLayerSettings(2).GetWidth + Interface.FixDPI(alphaOffset)
        
        'Horizontally stretch the opacity, blend, and alpha mode UI inputs
        sltLayerOpacity.SetWidth curFormWidth - sltLayerOpacity.GetLeft
        cboBlendMode.SetWidth curFormWidth - cboBlendMode.GetLeft
        cboAlphaMode.SetWidth curFormWidth - cboAlphaMode.GetLeft
        
        'Resize the layer box and associated scrollbar
        If (lstLayers.GetWidth <> curFormWidth - lstLayers.GetLeft) Then lstLayers.SetWidth curFormWidth - lstLayers.GetLeft
        
        'Reflow the bottom button box; this is inevitably more complicated, owing to the spacing requirements of the buttons
        ctlGroupLayerButtons.SetLeft lstLayers.GetLeft
        ctlGroupLayerButtons.SetWidth lstLayers.GetWidth
        
        'The total size of the button region is [numButtons * buttonWidth + (numButtons - 1) * buttonPadding],
        ' for e.g. N buttons and N-1 spacers.
        Dim btnPadding As Long, btnWidth As Long
        btnPadding = Interface.FixDPI(4)
        btnWidth = cmdLayerAction(0).GetWidth()
        
        Dim numLayerButtons As Long, numLayerButtonsAllowed As Long
        numLayerButtons = cmdLayerAction.Count
        numLayerButtonsAllowed = numLayerButtons
        
        Dim buttonAreaWidth As Long, buttonAreaLeft As Long
        buttonAreaWidth = Interface.FixDPI(numLayerButtons * btnWidth + (numLayerButtons - 1) * btnPadding)
        
        Dim btnContainerWidth As Long
        btnContainerWidth = ctlGroupLayerButtons.GetWidth
        
        Do While (buttonAreaWidth > btnContainerWidth)
            numLayerButtonsAllowed = numLayerButtonsAllowed - 1
            buttonAreaWidth = Interface.FixDPI(numLayerButtonsAllowed * btnWidth + (numLayerButtonsAllowed - 1) * btnPadding)
        Loop
        
        buttonAreaLeft = (btnContainerWidth - buttonAreaWidth) \ 2
        
        sizeCheck = btnWidth + btnPadding
        
        Dim i As Long
        For i = 0 To cmdLayerAction.Count - 1
            If (i < numLayerButtonsAllowed) Then
                cmdLayerAction(i).SetLeft buttonAreaLeft + (i * sizeCheck)
                cmdLayerAction(i).Visible = True
            Else
                cmdLayerAction(i).Visible = False
            End If
        Next i
    
    End If
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Add images to the layer action buttons at the bottom of the toolbox
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(26)
    cmdLayerAction(LYR_BTN_ADD).AssignImage "layer_add", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    cmdLayerAction(LYR_BTN_DELETE).AssignImage "layer_delete", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    cmdLayerAction(LYR_BTN_MOVE_UP).AssignImage "layer_up", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    cmdLayerAction(LYR_BTN_MOVE_DOWN).AssignImage "layer_down", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    cmdLayerAction(LYR_BTN_DUPLICATE).AssignImage "layer_duplicate", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Recreate tooltips (necessary to support run-time language changes)
    
    'Tooltips for these controls are now multiline, because they describe multiple interactions
    Dim ttTextNew As String
    ttTextNew = g_Language.TranslateMessage("Click: show the ""New layer"" dialog")
    ttTextNew = ttTextNew & vbCrLf & g_Language.TranslateMessage("Shift + Click: add a blank layer")
    cmdLayerAction(LYR_BTN_ADD).AssignTooltip ttTextNew, "Add layer"
    
    cmdLayerAction(LYR_BTN_DELETE).AssignTooltip "Click: delete this layer", "Delete layer"
    
    ttTextNew = g_Language.TranslateMessage("Click: move this layer up the layer stack")
    ttTextNew = ttTextNew & vbCrLf & g_Language.TranslateMessage("Shift + Click: merge this layer with the layer above it")
    cmdLayerAction(LYR_BTN_MOVE_UP).AssignTooltip ttTextNew, "Move or merge layer up"
    
    ttTextNew = g_Language.TranslateMessage("Click: move this layer down the layer stack")
    ttTextNew = ttTextNew & vbCrLf & g_Language.TranslateMessage("Shift + Click: merge this layer with the layer beneath it")
    cmdLayerAction(LYR_BTN_MOVE_DOWN).AssignTooltip ttTextNew, "Move or merge layer down"
    
    cmdLayerAction(LYR_BTN_DUPLICATE).AssignTooltip "Click: add a duplicate of this layer", "Duplicate layer"
        
    'Reflow the interface, to account for any language changes.  (This will also trigger a redraw of the layer list box.)
    ReflowInterface
        
End Sub
