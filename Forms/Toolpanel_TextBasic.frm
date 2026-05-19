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
      Top             =   4320
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
      Height          =   3135
      Index           =   0
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   4075
      Begin PhotoDemon.pdButtonToolbox cmdAddStyle 
         Height          =   435
         Left            =   7290
         TabIndex        =   20
         Top             =   2220
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdDropDown ddStyle 
         Height          =   720
         Left            =   90
         TabIndex        =   19
         Top             =   1890
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   1270
         Caption         =   "text style preset"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdCheckBox chkAutoOpenText 
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   2745
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   635
         Caption         =   "always open this panel for new text layers"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   7350
         TabIndex        =   1
         Top             =   2730
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
'Last updated: 01/May/25
'Last update: new text styles feature allows users to save all current text settings as a "style"
'             (a glorified text-specific preset)
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
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'While the dialog is loading, we need to suspend relaying changes to the active layer.
' (Otherwise, we may accidentally relay last-used settings from a previous image to the current one!)
Private m_suspendSettingRelay As Boolean

'Persistent text styles are handled by a pdToolPreset instance, while XML handling (used to save/load text styles)
' is handled through a specialized class.
Private m_Presets As pdToolPreset, m_Params As pdSerialize

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

Private Sub chkAutoOpenText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cmdAddStyle.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

'Add a new text style to the user's saved style collection.  This behaves similarly to the preset management
' in standalone PD windows; see the command bar UC for additional implementation details.
Private Sub cmdAddStyle_Click(ByVal Shift As ShiftConstants)
    
    'Opening a new dialog will auto-close the current flyout panel.
    ' To prevent this, lock it open *prior* to raising the dialog.
    Dim initFlyoutLockState As Boolean
    initFlyoutLockState = cmdFlyoutLock(0).Value
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(0).hWnd, True, cmdFlyoutLock(0)
    
    'Prompt the user for a style name
    Dim newNameReturn As VbMsgBoxResult, newPresetToSave As String
    newNameReturn = Dialogs.PromptNewPreset(m_Presets, newPresetToSave, Me)
    If (newNameReturn = vbOK) Then
    
        'The user added a new style, meaning we need to rebuild the style dropdown with the new entry.
        ' (They may also have *delete* existing styles; we'll deal with that possibility outside this branch.)
        
        'Start by disabling previews
        m_suspendSettingRelay = True
        
        'If we were given a new preset name to save, save it now
        If (LenB(newPresetToSave) > 0) Then StorePreset newPresetToSave
        
    End If
    
    'The user can remove presets and then *cancel* the dialog, so always re-load all presets
    ' regardless of OK/Cancel behavior.
    LoadAllPresets
    
    'If the user just added a preset, set the combo box index to match the preset they added
    If (newNameReturn = vbOK) And (LenB(newPresetToSave) <> 0) Then
    
        Dim i As Long
        For i = 0 To ddStyle.ListCount - 1
            If Strings.StringsEqual(newPresetToSave, Trim$(ddStyle.List(i)), True) Then
                ddStyle.ListIndex = i
                Exit For
            End If
        Next i
        
        'Re-enable previews
        m_suspendSettingRelay = False
        
    Else
        ddStyle.ListIndex = 0
    End If
    
    'Restore the original flyout lock state
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(0).hWnd, initFlyoutLockState, cmdFlyoutLock(0)
    
End Sub

Private Sub cmdAddStyle_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub cmdAddStyle_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ddStyle.hWnd
    Else
        newTargetHwnd = Me.chkAutoOpenText.hWnd
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
                newTargetHwnd = Me.chkAutoOpenText.hWnd
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

Private Sub ddStyle_Click()
    
    'Ignore the user selecting the top (blank) style settings
    If (ddStyle.ListIndex > 0) And (Not m_suspendSettingRelay) Then
        
        'Load the preset and refresh all UI elements accordingly
        LoadPreset ddStyle.List(ddStyle.ListIndex)
        
        'Use a special initialization command that basically copies all existing text properties into the newly created layer.
        Tools.SyncCurrentLayerToToolOptionsUI
        
        'Redraw the viewport immediately
        Dim tmpViewportParams As PD_ViewportParams
        tmpViewportParams = Viewport.GetDefaultParamObject()
        tmpViewportParams.curPOI = poi_CornerSE
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0), VarPtr(tmpViewportParams)
        
    End If
    
End Sub

Private Sub ddStyle_GotFocusAPI()
    UpdateFlyout 0, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_Style, ddStyle.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub ddStyle_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_Style, ddStyle.ListIndex
End Sub

Private Sub ddStyle_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.txtTextTool.hWnd
    Else
        newTargetHwnd = Me.cmdAddStyle.hWnd
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
        
        'Initialize a text styles object
        Set m_Presets = New pdToolPreset
        
        'Load any previously saved text styles
        Const PRESET_BASE_NAME As String = "text-basic-styles"
        m_Presets.SetPresetFilePath UserPrefs.GetPresetPath & PRESET_BASE_NAME & ".xml", PRESET_BASE_NAME, "text styles for the basic text tool"
        LoadAllPresets
        
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

Private Sub m_LastUsedSettings_ReadCustomPresetData()
    
    'We don't actually need to read anything here - we just want to always default the style dropdown
    ' to a "blank" value (so that last-used settings are used instead)
    ddStyle.ListIndex = 0
    
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
        newTargetHwnd = Me.ddStyle.hWnd
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

'Synchronize *all* UI elements on this page to reflect the current (basic text) layer's settings
Public Sub SyncSettingsToCurrentLayer()
    
    txtTextTool.Text = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_Text)
    cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontFace), vbTextCompare)
    sldTextFontSize.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontSize)
    csTextFontColor.Color = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontColor)
    cboTextRenderingHint.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextAntialiasing)
    sltTextClarity.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextContrast)
    btnFontStyles(0).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontBold))
    btnFontStyles(1).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontItalic))
    btnFontStyles(2).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontUnderline))
    btnFontStyles(3).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontStrikeout))
    btsHAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_HorizontalAlignment)
    btsVAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_VerticalAlignment)
    
    'For now, *unselect* any active styles.  (In the future, we will need to match these via some sort
    ' of checksum, since styles can be edited, so if the user previously applied a style, then changed the
    ' style contents, the layer wouldn't match correctly if we only tagged style *name* and not style *contents*.)
    ddStyle.ListIndex = 0
    
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
    
    cmdAddStyle.AssignImage "generic_savepreset", , buttonSize, buttonSize + 1
    cmdAddStyle.AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarSavePreset)
    ddStyle.AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarPresetList)
    
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

'Record the current value of all UI objects on our parent dialog, and return their combined value as an XML string.
' An optional preset name can be passed; note that this gets embedded in the XML, as well.
Private Function GetPresetParamString(Optional ByVal srcPresetName As String = "last-used settings") As String
    
    'Failsafe only; errors are not expected in this function
    On Error GoTo SkipPreset
    
    'Initialize a param handler and initialize it with the passed preset name
    If (m_Params Is Nothing) Then Set m_Params = New pdSerialize
    m_Params.Reset
    If (LenB(srcPresetName) <> 0) Then m_Params.AddParam "fullPresetName", srcPresetName, True
    
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    'Next, we're going to iterate through each control on the form.  For each control, we're going to assemble two things:
    ' a name (basically, the control name plus its index, if any), and its value.  These are forwarded to the preset manager,
    ' which handles the actual XML storage for each entry.
    Dim eControl As Object
    For Each eControl In Me.Controls
        
        'Retrieve the control name and index, if any
        controlName = eControl.Name
        If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
        
        'Reset our control value checker
        controlValue = vbNullString
        
        'Value retrieval must be handled uniquely for each possible control type (including custom PD-specific controls).
        controlType = TypeName(eControl)
        Select Case controlType
        
            'PD-specific sliders, checkboxes, option buttons, and text up/downs return a .Value property
            Case "pdSlider", "pdCheckBox", "pdRadioButton", "pdSpinner", "pdTitle", "pdScrollBar", "pdButtonToolbox"
                controlValue = Str$(eControl.Value)
            
            'List-type objects have a .ListIndex property
            Case "pdButtonStrip", "pdButtonStripVertical"
                controlValue = Str$(eControl.ListIndex)
            
            'Note that we don't store presets for the preset combo box itself!
            Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown"
                If (eControl.hWnd <> ddStyle.hWnd) Then controlValue = Str$(eControl.ListIndex)
            
            'Font dropdowns store last-used font by name.  (Font list size is *not* guaranteed to be consistent between sessions,
            ' unlike internal listboxes.)
            Case "pdDropDownFont"
                controlValue = eControl.List(eControl.ListIndex)
                    
            'Various PD controls have their own custom "value"-type properties.
            Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                controlValue = Str$(eControl.Color)
            
            Case "pdBrushSelector"
                controlValue = eControl.Brush
                
            Case "pdPenSelector"
                controlValue = eControl.Pen
                
            Case "pdGradientSelector"
                controlValue = eControl.Gradient
                
            'Text boxes will store a copy of their current text
            Case "pdTextBox"
                If (eControl.hWnd <> txtTextTool.hWnd) Then controlValue = eControl.Text
                
            Case "pdRandomizeUI"
                controlValue = eControl.Value
                
            'PD supports a number of other user controls, but they are not exposed on this form.
            ' (See the command bar UC for details on their implementation.)
                
        End Select
        
        'Remove VB's default padding from the generated string.  (Str() prepends positive numbers with a space)
        If (LenB(controlValue) <> 0) Then controlValue = Trim$(controlValue)
        
        'If the control value still has a non-zero length, add it now
        If (LenB(controlValue) <> 0) Then
            If (controlIndex >= 0) Then
                m_Params.AddParam controlName & ":" & controlIndex, controlValue
            Else
                m_Params.AddParam controlName, controlValue
            End If
        End If
        
    'Continue with the next control on the parent dialog
    Next eControl
    
    GetPresetParamString = m_Params.GetParamString()

SkipPreset:

End Function

'Search the preset file for all valid text style presets.  This sub doesn't actually load any of the presets -
' it just adds their names to the text styles combo box.
Private Sub LoadAllPresets(Optional ByVal newListIndex As Long = 0)

    ddStyle.Clear
    ddStyle.SetAutomaticRedraws False
    
    'We always add one blank entry to the preset combo box, which is selected by default
    ddStyle.AddItem " ", 0, True

    'Query the preset manager for any available presets.  If found, it will return the number of available presets
    Dim listOfPresets As pdStringStack
    If (m_Presets.GetListOfPresets(listOfPresets) > 0) Then
        
        'Add all discovered presets to the combo box.  Note that we do not use a traditional stack pop here,
        ' as that would reverse the preset order!
        Dim i As Long
        For i = 0 To listOfPresets.GetNumOfStrings() - 1
            ddStyle.AddItem listOfPresets.GetString(i), i + 1, False
        Next i
        
    End If
    
    'When finished, set the requested list index
    ddStyle.SetAutomaticRedraws True
    ddStyle.ListIndex = newListIndex

End Sub

'This sub will set the values of all controls on this form, using the values stored in the tool's XML file under the
' "presetName" section.  By default, it will look for the last-used settings, as this is its most common request.
Private Function LoadPreset(Optional ByVal srcPresetName As String = "last-used settings", Optional ByVal loadEverything As Boolean = True) As Boolean
    
    'Start by asking the preset engine if the requested preset even exists in the file
    Dim presetExists As Boolean
    presetExists = m_Presets.DoesPresetExist(srcPresetName)
    
    'If the preset exists, continue with the load process
    If presetExists Then
        LoadPreset = LoadPresetFromString(m_Presets.GetPresetXML(srcPresetName), loadEverything)
                
    'If the preset does *not* exist, exit without further processing
    Else
        LoadPreset = False
        Exit Function
    End If
    
End Function

Private Function LoadPresetFromString(ByRef srcString As String, Optional ByVal loadEverything As Boolean = True) As Boolean

    'Copy this preset's XML into a local param evaluator
    If (m_Params Is Nothing) Then Set m_Params = New pdSerialize
    m_Params.SetParamString srcString
    
    'Loading preset values involves (potentially) changing the value of every single object on this form.  To prevent each
    ' of these changes from triggering a full preview redraw, we forcibly suspend previews now.
    m_suspendSettingRelay = True
    
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    'If parameters allow, iterate through each control on the form and attempt to retrieve its last-used value
    Dim eControl As Object
    
    If loadEverything Then
    
        For Each eControl In Me.Controls
            
            'Control values are saved by control name, and if it exists, control index.  We start by generating a matching preset
            ' name for this control.
            controlName = eControl.Name
            If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
            If (controlIndex >= 0) Then controlName = controlName & ":" & controlIndex
            
            Dim okToLoad As Boolean: okToLoad = True
            'If (Not m_NoLoadList Is Nothing) Then okToLoad = (m_NoLoadList.ContainsString(controlName, True) < 0)
            
            'See if a preset exists for this control and this particular preset
            If (okToLoad And m_Params.GetStringEx(controlName, controlValue)) Then
                
                'A value for this control exists, and it has been retrieved into controlValue.  We sort handling of this value
                ' by control type, as different controls require different input values (bool, int, etc).
                controlType = TypeName(eControl)
            
                Select Case controlType
                
                    'Sliders and text up/downs allow for floating-point values, so we always cast these returns as doubles
                    Case "pdSlider", "pdSpinner"
                        eControl.Value = CDblCustom(controlValue)
                    
                    'Check boxes use a long (technically a boolean, as PD's custom check box doesn't support a gray state, but for
                    ' backward compatibility with VB check box constants, we cast to a Long)
                    Case "pdCheckBox"
                        eControl.Value = CBool(controlValue)
                    
                    'Option buttons use booleans
                    Case "pdRadioButton"
                        If CBool(controlValue) Then eControl.Value = CBool(controlValue)
                    
                    'Toolbox-style buttons should only be saved if they use the sticky-toggle feature
                    Case "pdButtonToolbox"
                        If eControl.StickyToggle Then eControl.Value = CBool(controlValue)
                        
                    'Button strips are similar to list boxes, so they use a .ListIndex property
                    Case "pdButtonStrip", "pdButtonStripVertical"
                    
                        'To protect against future changes that modify the number of available entries in a button strip, we always
                        ' validate the list index against the current list count prior to setting it.
                        If (CLng(controlValue) < eControl.ListCount) Then
                            eControl.ListIndex = CLng(controlValue)
                        Else
                            If (eControl.ListCount > 0) Then eControl.ListIndex = eControl.ListCount - 1
                        End If
                    
                    'Various PD controls have their own custom "value"-type properties.
                    Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                        eControl.Color = CLng(controlValue)
                               
                    Case "pdBrushSelector"
                        eControl.Brush = controlValue
                    
                    Case "pdPenSelector"
                        eControl.Pen = controlValue
                    
                    Case "pdGradientSelector"
                        eControl.Gradient = controlValue
                    
                    'Traditional scroll bar values are cast as Longs, despite them only having Int ranges
                    ' (hopefully the original caller planned for this!)
                    Case "pdScrollBar"
                        eControl.Value = CLng(controlValue)
                    
                    'List boxes and dropdowns all use a Long-type .ListIndex property
                    Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown"
                    
                        'Validate range before setting
                        If (CLng(controlValue) < eControl.ListCount) Then
                            eControl.ListIndex = CLng(controlValue)
                        Else
                            If (eControl.ListCount > 0) Then eControl.ListIndex = eControl.ListCount - 1
                        End If
                    
                    'Font dropdowns store last-used font by name.  (Font list size is *not* guaranteed to be consistent between sessions,
                    ' unlike internal listboxes.)
                    Case "pdDropDownFont"
                        Dim fontListIndex As Long
                        fontListIndex = eControl.ListIndexByString(controlValue, vbTextCompare)
                        If (fontListIndex >= 0) Then eControl.ListIndex = fontListIndex Else eControl.ListIndex = eControl.ListIndexByString(Fonts.GetUIFontName())
                    
                    'Text boxes just take the stored string as-is
                    Case "TextBox", "pdTextBox"
                        eControl.Text = controlValue
                    
                    'PD supports a number of other user controls, but they are not exposed on this form.
                    ' (See the command bar UC for details on their implementation.)
                
                End Select
    
            End If
        
        'Iterate through the next control
        Next eControl
        
    End If
    
    'Re-enable previews
    m_suspendSettingRelay = False
    
    'If the parent dialog is active (e.g. this function is not occurring during the parent dialog's Load process),
    ' request a preview update as the preview has likely changed due to the new control values.
    'If m_controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
    'TODO: here or elsewhere?  relay *all* changes to base layer
    
    'This function's return isn't meaningful at present
    LoadPresetFromString = True
        
End Function

'This sub will fill the class's pdXML class (xmlEngine) with the values of all controls on this form, and it will store
' those values in the section titled "presetName".
Private Sub StorePreset(Optional ByVal srcPresetName As String = "last-used settings")
    
    'Make sure PD's built-in "last-used settings" text is properly translated
    If (Not g_Language Is Nothing) And Strings.StringsEqual(srcPresetName, "last-used settings", True) Then srcPresetName = g_Language.TranslateMessage("last-used settings")
    srcPresetName = Trim$(srcPresetName)
    
    'An external function handles the actual XML assembly.
    m_Presets.AddPreset srcPresetName, GetPresetParamString(srcPresetName)
    
    'Because the user may still cancel the dialog, we want to request an XML file dump immediately,
    ' so the recently added preset won't be lost.
    m_Presets.WritePresetFile
    
End Sub
