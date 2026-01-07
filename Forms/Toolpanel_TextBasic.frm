VERSION 5.00
Begin VB.Form toolpanel_TextBasic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14955
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
   Icon            =   "Toolpanel_TextBasic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   997
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdHyperlink hypEditText 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   "click to edit text"
      RaiseClickEvent =   -1  'True
   End
   Begin PhotoDemon.pdContainer picConvertLayer 
      Height          =   1695
      Left            =   120
      Top             =   3960
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   2990
      Begin PhotoDemon.pdButton cmdConvertLayer 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         Caption         =   "yes"
      End
      Begin PhotoDemon.pdLabel lblConvertLayer 
         Height          =   735
         Left            =   5280
         Top             =   120
         Width           =   5640
         _ExtentX        =   19050
         _ExtentY        =   1296
         Alignment       =   2
         Caption         =   "yes"
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      Caption         =   "edit text"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2310
      Index           =   0
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   4075
      Begin PhotoDemon.pdCheckBox chkAutoOpenText 
         Height          =   360
         Left            =   90
         TabIndex        =   18
         Top             =   1905
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   635
         Caption         =   "always open this panel for new text layers"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   7350
         TabIndex        =   1
         Top             =   1875
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtTextTool 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   30
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3201
         Multiline       =   -1  'True
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      Caption         =   "font"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdButtonStrip btsHAlignment 
      Height          =   435
      Left            =   7950
      TabIndex        =   4
      Top             =   345
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdButtonStrip btsVAlignment 
      Height          =   435
      Left            =   9510
      TabIndex        =   5
      Top             =   345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdDropDownFont cboTextFontFace 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdColorSelector csTextFontColor 
      Height          =   750
      Left            =   5640
      TabIndex        =   7
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1323
      Caption         =   "color"
      FontSize        =   10
      curColor        =   0
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1740
      Index           =   1
      Left            =   8400
      Top             =   840
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3069
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   5640
         TabIndex        =   8
         Top             =   1170
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sldTextFontSize 
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         Caption         =   "size"
         FontSizeCaption =   10
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   16
         NotchPosition   =   2
         NotchValueCustom=   16
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   300
         Index           =   2
         Left            =   120
         Top             =   840
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         Caption         =   "style"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   1
         Left            =   720
         TabIndex        =   11
         Top             =   1200
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   2
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   3
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltTextClarity 
         Height          =   765
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1349
         Caption         =   "clarity"
         FontSizeCaption =   10
         Value           =   5
         NotchPosition   =   2
         NotchValueCustom=   5
      End
      Begin PhotoDemon.pdDropDown cboTextRenderingHint 
         Height          =   735
         Left            =   3360
         TabIndex        =   15
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         Caption         =   "antialiasing"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   300
      Index           =   0
      Left            =   7950
      Top             =   30
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   529
      Caption         =   "alignment"
      ForeColor       =   0
   End
End
Attribute VB_Name = "toolpanel_TextBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Basic Text Tool Panel
'Copyright 2013-2026 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 09/March/22
'Last update: new checkbox for auto-dropping text entry field after creating a new text layer
'
'This form includes all user-editable settings for the Basic Text tool.
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

'While the dialog is loading, we need to suspend relaying changes to the active layer.
' (Otherwise, we may accidentally relay last-used settings from a previous image to the current one!)
Private m_suspendSettingRelay As Boolean

Private Sub btnFontStyles_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
            
        'Italic
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        'Underline
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        'Strikeout
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
    
    End Select
    
End Sub

Private Sub btnFontStyles_LostFocusAPI(Index As Integer)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Evaluate Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagFinalNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value
            
        'Italic
        Case 1
            Processor.FlagFinalNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            Processor.FlagFinalNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            Processor.FlagFinalNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
End Sub

Private Sub btnFontStyles_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = Me.sldTextFontSize.hWndSpinner
            Case Else
                newTargetHwnd = Me.btnFontStyles(Index - 1).hWnd
        End Select
    Else
        Select Case Index
            Case 0, 1, 2
                newTargetHwnd = Me.btnFontStyles(Index + 1).hWnd
            Case Else
                newTargetHwnd = Me.cboTextRenderingHint.hWnd
        End Select
    End If
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsHAlignment_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.csTextFontColor.hWnd
    Else
        newTargetHwnd = Me.btsVAlignment.hWnd
    End If
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub btsVAlignment_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsHAlignment.hWnd
    Else
        newTargetHwnd = Me.ttlPanel(0).hWnd
    End If
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextFontFace_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(1).hWnd
    Else
        newTargetHwnd = Me.sldTextFontSize.hWndSlider
    End If
End Sub

Private Sub cboTextRenderingHint_Click()
        
    'We show/hide the AA clarity option depending on this tool's setting.  (AA clarity doesn't make much sense
    ' if AA is disabled.)
    If (cboTextRenderingHint.ListIndex = 0) Then
        sltTextClarity.Visible = False
    Else
        sltTextClarity.Visible = True
    End If
        
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub cboTextRenderingHint_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btnFontStyles(3).hWnd
    Else
        newTargetHwnd = Me.sltTextClarity.hWndSlider
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
                newTargetHwnd = Me.txtTextTool.hWnd
            Else
                newTargetHwnd = Me.ttlPanel(1).hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sltTextClarity.hWndSpinner
            Else
                newTargetHwnd = Me.csTextFontColor.hWnd
            End If
    End Select
    
End Sub

Private Sub csTextFontColor_ColorChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontColor, csTextFontColor.Color
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub csTextFontColor_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontColor, csTextFontColor.Color, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub csTextFontColor_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontColor, csTextFontColor.Color
End Sub

Private Sub csTextFontColor_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
    Else
        newTargetHwnd = Me.btsHAlignment.hWnd
    End If
End Sub

Private Sub Form_Load()
    
    m_suspendSettingRelay = True
    
    'Disable any layer updates as a result of control changes during the load process
    Tools.SetToolBusyState True
    
    'Forcibly hide the "convert to text layer" panel.  (This appears when a typography layer
    ' is active, to allow the user to switch back-and-forth between typography and text layers.)
    toolpanel_TextBasic.picConvertLayer.Visible = False
    
    If PDMain.IsProgramRunning() Then
        
        'Generate a list of fonts
        cboTextFontFace.InitializeFontList
        cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(Fonts.GetUIFontName(), vbBinaryCompare)
        
        'Antialiasing options behave slightly differently from the advanced text tool
        cboTextRenderingHint.SetAutomaticRedraws False
        cboTextRenderingHint.Clear
        cboTextRenderingHint.AddItem "none", 0
        cboTextRenderingHint.AddItem "normal", 1
        cboTextRenderingHint.AddItem "crisp", 2
        cboTextRenderingHint.ListIndex = 1
        cboTextRenderingHint.SetAutomaticRedraws True
        
        'Add dummy entries to the various alignment buttons; we'll populate these with theme-specific
        ' images in the UpdateAgainstCurrentTheme() function.
        btsHAlignment.AddItem vbNullString, 0
        btsHAlignment.AddItem vbNullString, 1
        btsHAlignment.AddItem vbNullString, 2
        
        btsVAlignment.AddItem vbNullString, 0
        btsVAlignment.AddItem vbNullString, 1
        btsVAlignment.AddItem vbNullString, 2
        
        'Load any last-used settings for this form
        Set m_lastUsedSettings = New pdLastUsedSettings
        m_lastUsedSettings.SetParentForm Me
        m_lastUsedSettings.LoadAllControlValues
        
    End If
    
    Tools.SetToolBusyState False
    
    m_suspendSettingRelay = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    UpdateAgainstCurrentLayer
End Sub

Private Sub hypEditText_Click()
    UpdateFlyout 0, True
    Me.txtTextTool.SetFocusToEditBox False
    Me.txtTextTool.SelStart = Len(Me.txtTextTool.Text)
End Sub

Private Sub hypEditText_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub hypEditText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(0).hWnd
    Else
        newTargetHwnd = Me.txtTextTool.hWnd
    End If
End Sub

Private Sub cmdConvertLayer_Click()
        
    If (Not PDImages.IsImageActive()) Then Exit Sub
        
    'Because of the way this warning panel is constructed, this label will not be visible unless a click is valid.
    PDImages.GetActiveImage.GetActiveLayer.SetLayerType PDL_TextBasic
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
    
    'Hide the warning panel and redraw both the viewport, and the UI (as new UI options may now be available)
    Me.UpdateAgainstCurrentLayer
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    Interface.SyncInterfaceToCurrentImage
    
End Sub

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub sldTextFontSize_Change()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FontSize, sldTextFontSize.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sldTextFontSize_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontSize, sldTextFontSize.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sldTextFontSize_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontSize, sldTextFontSize.Value
End Sub

Private Sub sldTextFontSize_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboTextFontFace.hWnd
    Else
        newTargetHwnd = Me.btnFontStyles(0).hWnd
    End If
End Sub

Private Sub sltTextClarity_Change()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_TextContrast, sltTextClarity.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub sltTextClarity_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextContrast, sltTextClarity.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltTextClarity_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextContrast, sltTextClarity.Value
End Sub

Private Sub sltTextClarity_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboTextRenderingHint.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.btsVAlignment.hWnd
            Else
                newTargetHwnd = Me.hypEditText.hWnd
            End If
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
            Else
                newTargetHwnd = Me.cboTextFontFace.hWnd
            End If
    End Select
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub txtTextTool_GotFocusAPI()
    UpdateFlyout 0, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_Text, txtTextTool.Text, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub

Private Sub txtTextTool_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.hypEditText.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

'Outside functions can forcibly request an update against the current layer.  If the current layer is
' a non-basic text layer, an option will be displayed to convert the layer.
Public Sub UpdateAgainstCurrentLayer()
    
    If PDImages.IsImageActive() Then

        If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then
        
            'Check for non-basic-text layers.
            If (PDImages.GetActiveImage.GetActiveLayer.GetLayerType <> PDL_TextBasic) Then
            
                Select Case PDImages.GetActiveImage.GetActiveLayer.GetLayerType()
                
                    Case PDL_TextAdvanced
                        Dim newMessage As String
                        newMessage = g_Language.TranslateMessage("This is an advanced text layer.  To edit it with the basic text tool, you must first convert it to a basic text layer.")
                        newMessage = newMessage & Space$(2) & g_Language.TranslateMessage("(This action is non-destructive.)")
                        Me.lblConvertLayer.Caption = newMessage
                        
                    'In the future, other text layer types can be added here.
                
                End Select
            
                Me.cmdConvertLayer.Caption = g_Language.TranslateMessage("Click here to convert this layer to basic text.")
                
                'Make the prompt panel the size of the tool window
                Me.picConvertLayer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
                
                'Left-align the convert command
                Me.cmdConvertLayer.SetPositionAndSize 1, 1, Me.cmdConvertLayer.GetWidth, Me.picConvertLayer.GetHeight - 2
                
                'Align the conversion explanation next to the command button Center all labels on the panel.
                Dim lblPadding As Long, newLeft As Long
                lblPadding = Interface.FixDPI(16)
                newLeft = Me.cmdConvertLayer.GetLeft + Me.cmdConvertLayer.GetWidth + lblPadding
                Me.lblConvertLayer.SetPositionAndSize newLeft, 0, Me.picConvertLayer.GetWidth - (newLeft + lblPadding), Me.picConvertLayer.GetHeight
                
                'Display the panel
                Me.picConvertLayer.Visible = True
                Me.picConvertLayer.ZOrder 0
                
            Else
                Me.picConvertLayer.Visible = False
            End If
        
        Else
            Me.picConvertLayer.Visible = False
        End If
        
    Else
        Me.picConvertLayer.Visible = False
    End If

End Sub

'Most objects on this form can avoid doing any work if the current layer is not a text layer.
Private Function CurrentLayerIsText() As Boolean
    
    CurrentLayerIsText = False
    
    'Changing UI elements does nothing if no images are loaded
    If PDImages.IsImageActive() Then
        If (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            CurrentLayerIsText = PDImages.GetActiveImage.GetActiveLayer.IsLayerText
        End If
    End If
    
End Function

'When a new text layer is created, the user can choose to auto-drop the text entry panel.
Public Sub NotifyNewLayerCreated()
    If Me.chkAutoOpenText.Value Then
        UpdateFlyout 0, True
        Me.txtTextTool.SetFocusToEditBox True
    End If
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current UI theme settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update any UI images against the current theme
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(24)
    
    btnFontStyles(0).AssignImage "format_bold", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btnFontStyles(1).AssignImage "format_italic", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btnFontStyles(2).AssignImage "format_underline", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btnFontStyles(3).AssignImage "format_strikethrough", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    btsHAlignment.AssignImageToItem 0, "format_alignleft", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsHAlignment.AssignImageToItem 1, "format_aligncenter", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsHAlignment.AssignImageToItem 2, "format_alignright", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    btsVAlignment.AssignImageToItem 0, "format_aligntop", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsVAlignment.AssignImageToItem 1, "format_alignmiddle", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsVAlignment.AssignImageToItem 2, "format_alignbottom", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    'Finish redrawing the form according to current theme and translation settings
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
