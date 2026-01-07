VERSION 5.00
Begin VB.Form toolpanel_TextAdvanced 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18435
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
   Icon            =   "Toolpanel_Typography.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1229
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2310
      Index           =   0
      Left            =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   4075
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   7350
         TabIndex        =   2
         Top             =   1875
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtTextTool 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   30
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3201
         FontSize        =   9
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkAutoOpenText 
         Height          =   360
         Left            =   90
         TabIndex        =   43
         Top             =   1905
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   635
         Caption         =   "always open this panel for new text layers"
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   3540
      Index           =   1
      Left            =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   11033
      _ExtentY        =   7726
      Begin PhotoDemon.pdButtonStrip btsHinting 
         Height          =   855
         Left            =   3330
         TabIndex        =   20
         Top             =   840
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1508
         Caption         =   "hinting"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   7920
         TabIndex        =   4
         Top             =   3000
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sldTextFontSize 
         Height          =   735
         Left            =   120
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   2145
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   300
         Index           =   0
         Left            =   120
         Top             =   1785
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
         TabIndex        =   7
         Top             =   2145
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Top             =   2145
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnFontStyles 
         Height          =   435
         Index           =   3
         Left            =   1680
         TabIndex        =   9
         Top             =   2145
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboTextRenderingHint 
         Height          =   735
         Left            =   3330
         TabIndex        =   10
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1296
         Caption         =   "antialiasing"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltCharSpacing 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   2655
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         Caption         =   "spacing"
         FontSizeCaption =   10
         Min             =   -1
         Max             =   1
         SigDigits       =   3
      End
      Begin PhotoDemon.pdSlider sltCharOrientation 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   3330
         TabIndex        =   14
         Top             =   2655
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1296
         Caption         =   "orientation"
         FontSizeCaption =   10
         Min             =   -360
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.pdDropDown cboCharCase 
         Height          =   735
         Left            =   3330
         TabIndex        =   15
         Top             =   1785
         Width           =   2355
         _ExtentX        =   4577
         _ExtentY        =   1296
         Caption         =   "remap"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboCharMirror 
         Height          =   735
         Left            =   5880
         TabIndex        =   16
         Top             =   1785
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1296
         Caption         =   "mirror"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   270
         Index           =   3
         Left            =   5880
         Top             =   0
         Width           =   2355
         _ExtentX        =   5371
         _ExtentY        =   476
         Caption         =   "jitter (x, y)"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdSpinner tudJitter 
         Height          =   345
         Index           =   1
         Left            =   7160
         TabIndex        =   17
         Top             =   390
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         Max             =   100
         SigDigits       =   1
      End
      Begin PhotoDemon.pdSpinner tudJitter 
         Height          =   345
         Index           =   0
         Left            =   5940
         TabIndex        =   18
         Top             =   390
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         Max             =   100
         SigDigits       =   1
      End
      Begin PhotoDemon.pdSlider sltCharInflation 
         CausesValidation=   0   'False
         Height          =   855
         Left            =   5880
         TabIndex        =   19
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1508
         Caption         =   "inflate"
         CaptionPadding  =   3
         FontSizeCaption =   10
         Max             =   20
         SigDigits       =   1
      End
      Begin PhotoDemon.pdButtonStrip btsStretch 
         Height          =   855
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1508
         Caption         =   "automatic fit"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2625
      Index           =   3
      Left            =   8640
      Top             =   4200
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4630
      Begin PhotoDemon.pdSlider sldLineSpacing 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         Caption         =   "line spacing"
         FontSizeCaption =   10
         Min             =   -100
         Max             =   1000
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   3
         Left            =   5820
         TabIndex        =   34
         Top             =   2160
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboWordWrap 
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1296
         Caption         =   "line wrap"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   1
         Left            =   3240
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   423
         Caption         =   "horizontal padding"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   2
         Left            =   3240
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   423
         Caption         =   "vertical padding"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdSpinner tudMargin 
         Height          =   345
         Index           =   0
         Left            =   3360
         TabIndex        =   37
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.pdSpinner tudMargin 
         Height          =   345
         Index           =   1
         Left            =   4800
         TabIndex        =   38
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.pdSpinner tudMargin 
         Height          =   345
         Index           =   2
         Left            =   3360
         TabIndex        =   39
         Top             =   1320
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.pdSpinner tudMargin 
         Height          =   345
         Index           =   3
         Left            =   4800
         TabIndex        =   40
         Top             =   1320
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.pdButtonStrip btsHAlignJustify 
         Height          =   435
         Left            =   150
         TabIndex        =   45
         Top             =   450
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   4
         Left            =   150
         Top             =   120
         Width           =   2940
         _ExtentX        =   5106
         _ExtentY        =   423
         Caption         =   "last line justify"
         ForeColor       =   0
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   3090
      Index           =   2
      Left            =   8400
      Top             =   840
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5450
      Begin PhotoDemon.pdCheckBox chkFillFirst 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         Caption         =   "outline on top"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   5880
         TabIndex        =   22
         Top             =   2640
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdBrushSelector bsText 
         Height          =   855
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   2775
         _ExtentX        =   3625
         _ExtentY        =   1508
      End
      Begin PhotoDemon.pdCheckBox chkOutlineText 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         Caption         =   "outline text"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdPenSelector psText 
         Height          =   855
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4260
         _ExtentY        =   1508
      End
      Begin PhotoDemon.pdCheckBox chkBackground 
         Height          =   330
         Left            =   3240
         TabIndex        =   26
         Top             =   0
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         Caption         =   "fill background"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdBrushSelector bsTextBackground 
         Height          =   855
         Left            =   3480
         TabIndex        =   27
         Top             =   360
         Width           =   2775
         _ExtentX        =   3625
         _ExtentY        =   1508
      End
      Begin PhotoDemon.pdPenSelector psTextBackground 
         Height          =   855
         Left            =   3480
         TabIndex        =   28
         Top             =   1680
         Width           =   2775
         _ExtentX        =   3625
         _ExtentY        =   1508
      End
      Begin PhotoDemon.pdCheckBox chkBackgroundBorder 
         Height          =   330
         Left            =   3240
         TabIndex        =   29
         Top             =   1320
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         Caption         =   "outline background"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkFillText 
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         Caption         =   "fill text"
      End
   End
   Begin PhotoDemon.pdHyperlink hypEditText 
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   405
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   "click to edit text"
      RaiseClickEvent =   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      Caption         =   "edit text"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      Caption         =   "font"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdDropDownFont cboTextFontFace 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   2
      Left            =   5640
      TabIndex        =   21
      Top             =   0
      Width           =   2055
      _ExtentX        =   5318
      _ExtentY        =   635
      Caption         =   "fill and outline"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdHyperlink hypEditStyles 
      Height          =   375
      Left            =   5640
      TabIndex        =   41
      Top             =   405
      Width           =   2055
      _ExtentX        =   4048
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   "click to edit"
      RaiseClickEvent =   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsHAlignment 
      Height          =   435
      Left            =   7950
      TabIndex        =   42
      Top             =   345
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdButtonStrip btsVAlignment 
      Height          =   435
      Left            =   9990
      TabIndex        =   32
      Top             =   345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      ColorScheme     =   1
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   3
      Left            =   7920
      TabIndex        =   33
      Top             =   0
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   635
      Caption         =   "alignment"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer picConvertLayer 
      Height          =   1695
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   23945
      _ExtentY        =   2990
      Begin PhotoDemon.pdButton cmdConvertLayer 
         Height          =   615
         Left            =   120
         TabIndex        =   0
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
End
Attribute VB_Name = "toolpanel_TextAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Advanced Typography Tool Panel
'Copyright 2013-2026 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 17/May/22
'Last update: new stretch-to-fit option
'
'This form includes all user-editable settings for PD's Advanced Typography text tool.
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
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'While the dialog is loading, we need to suspend relaying changes to the active layer.
' (Otherwise, we may accidentally relay last-used settings from a previous image to the current one!)
Private m_suspendSettingRelay As Boolean

Private Sub bsText_BrushChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FillBrush, bsText.Brush
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub bsText_GotFocusAPI()
    UpdateFlyout 2, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FillBrush, bsText.Brush, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub bsText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FillBrush, bsText.Brush
End Sub

Private Sub bsText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkFillText.hWnd
    Else
        newTargetHwnd = Me.chkOutlineText.hWnd
    End If
End Sub

Private Sub bsTextBackground_BrushChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_BackgroundBrush, bsTextBackground.Brush
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub bsTextBackground_GotFocusAPI()
    UpdateFlyout 2, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackgroundBrush, bsTextBackground.Brush, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub bsTextBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackgroundBrush, bsTextBackground.Brush
End Sub

Private Sub bsTextBackground_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkBackground.hWnd
    Else
        newTargetHwnd = Me.chkBackgroundBorder.hWnd
    End If
End Sub

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
    
    'Non-destructive effects are obviously not tracked if no images are loaded
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
        If (Index = 0) Then
            newTargetHwnd = Me.btsStretch.hWnd
        Else
            newTargetHwnd = Me.btnFontStyles(Index - 1).hWnd
        End If
    Else
        If (Index = 3) Then
            newTargetHwnd = Me.sltCharSpacing.hWndSlider
        Else
            newTargetHwnd = Me.btnFontStyles(Index + 1).hWnd
        End If
    End If
End Sub

Private Sub btsHAlignJustify_Click(ByVal buttonIndex As Long)

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_AlignLastLine, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsHAlignJustify_GotFocusAPI()
    UpdateFlyout 3, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_AlignLastLine, btsHAlignJustify.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsHAlignJustify_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_AlignLastLine, btsHAlignJustify.ListIndex
End Sub

Private Sub btsHAlignJustify_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsVAlignment.hWnd
    Else
        newTargetHwnd = Me.sldLineSpacing.hWndSlider
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
    UpdateFlyout 3, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsHAlignment_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(3).hWnd
    Else
        newTargetHwnd = Me.btsVAlignment.hWnd
    End If
End Sub

Private Sub btsHinting_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboTextRenderingHint.hWnd
    Else
        newTargetHwnd = Me.cboCharCase.hWnd
    End If
End Sub

Private Sub btsStretch_Click(ByVal buttonIndex As Long)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    Tools.SetToolBusyState True
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_StretchToFit, btsStretch.ListIndex
    Tools.SetToolBusyState False
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub btsStretch_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_StretchToFit, btsStretch.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsStretch_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_StretchToFit, btsStretch.ListIndex
End Sub

Private Sub btsStretch_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sldTextFontSize.hWndSpinner
    Else
        newTargetHwnd = Me.btnFontStyles(0).hWnd
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
    UpdateFlyout 3, True
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
        newTargetHwnd = Me.btsHAlignJustify.hWnd
    End If
End Sub

Private Sub cboCharCase_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharRemap, cboCharCase.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboCharCase_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharRemap, cboCharCase.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboCharCase_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharRemap, cboCharCase.ListIndex
End Sub

Private Sub cboCharCase_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsHinting.hWnd
    Else
        newTargetHwnd = Me.sltCharOrientation.hWnd
    End If
End Sub

Private Sub cboCharMirror_Click()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharMirror, cboCharMirror.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub cboCharMirror_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharMirror, cboCharMirror.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboCharMirror_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharMirror, cboCharMirror.ListIndex
End Sub

Private Sub cboCharMirror_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sltCharInflation.hWndSpinner
    Else
        newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
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
        newTargetHwnd = Me.sltCharSpacing.hWndSpinner
    Else
        newTargetHwnd = Me.btsHinting.hWnd
    End If
End Sub

Private Sub cboWordWrap_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_WordWrap, cboWordWrap.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboWordWrap_GotFocusAPI()
    UpdateFlyout 3, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboWordWrap_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex
End Sub

Private Sub cboWordWrap_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sldLineSpacing.hWndSpinner
    Else
        newTargetHwnd = Me.tudMargin(0).hWnd
    End If
End Sub

Private Sub chkBackground_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_BackgroundActive, chkBackground.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub chkBackground_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackgroundActive, chkBackground.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackgroundActive, chkBackground.Value
End Sub

Private Sub chkBackground_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkFillFirst.hWnd
    Else
        newTargetHwnd = Me.bsTextBackground.hWnd
    End If
End Sub

Private Sub chkBackgroundBorder_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_BackBorderActive, chkBackgroundBorder.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub chkBackgroundBorder_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackBorderActive, chkBackgroundBorder.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkBackgroundBorder_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackBorderActive, chkBackgroundBorder.Value
End Sub

Private Sub chkBackgroundBorder_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.bsTextBackground.hWnd
    Else
        newTargetHwnd = Me.psTextBackground.hWnd
    End If
End Sub

Private Sub chkFillFirst_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_OutlineAboveFill, chkFillFirst.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub chkFillFirst_GotFocusAPI()
    UpdateFlyout 2, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_OutlineAboveFill, chkFillFirst.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkFillFirst_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_OutlineAboveFill, chkFillFirst.Value
End Sub

Private Sub chkFillFirst_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.psText.hWnd
    Else
        newTargetHwnd = Me.chkBackground.hWnd
    End If
End Sub

Private Sub chkFillText_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_FillActive, chkFillText.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub chkFillText_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FillActive, chkFillText.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkFillText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FillActive, chkFillText.Value
End Sub

Private Sub btsHinting_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_TextHinting, (btsHinting.ListIndex = 1)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub btsHinting_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextHinting, (btsHinting.ListIndex = 1), PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub btsHinting_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextHinting, (btsHinting.ListIndex = 1)
End Sub

Private Sub chkFillText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.hypEditStyles.hWnd
    Else
        newTargetHwnd = Me.bsText.hWnd
    End If
End Sub

Private Sub chkOutlineText_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_OutlineActive, chkOutlineText.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub chkOutlineText_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_OutlineActive, chkOutlineText.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkOutlineText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_OutlineActive, chkOutlineText.Value
End Sub

Private Sub chkOutlineText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.bsText.hWnd
    Else
        newTargetHwnd = Me.psText.hWnd
    End If
End Sub

Private Sub cmdConvertLayer_Click()
        
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Because of the way this warning panel is constructed, this label will not be visible unless a click is valid.
    PDImages.GetActiveImage.GetActiveLayer.SetLayerType PDL_TextAdvanced
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
    
    'Hide the warning panel and redraw both the viewport, and the UI (as new UI options may now be available)
    Me.UpdateAgainstCurrentLayer
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    Interface.SyncInterfaceToCurrentImage
    
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
                newTargetHwnd = Me.txtTextTool.hWnd
            Case 1
                newTargetHwnd = Me.cboCharMirror.hWnd
            Case 2
                newTargetHwnd = Me.psTextBackground.hWnd
            Case 3
                newTargetHwnd = Me.tudMargin(3).hWnd
        End Select
        
    Else
        
        Dim newIndex As Long
        newIndex = Index + 1
        If (newIndex > 3) Then newIndex = newIndex - 4
        
        newTargetHwnd = Me.ttlPanel(newIndex).hWnd
        
    End If
End Sub

Private Sub Form_Load()
    
    m_suspendSettingRelay = True
    
    'Disable any layer updates as a result of control changes during the load process
    Tools.SetToolBusyState True
    
    'Forcibly hide the "convert to text layer" panel.  (This appears when a text layer
    ' is active, to allow the user to switch back-and-forth between typography and text layers.)
    toolpanel_TextBasic.picConvertLayer.Visible = False
    
    'Generate a list of fonts
    If PDMain.IsProgramRunning() Then
        
        'Initialize the font list
        cboTextFontFace.InitializeFontList
        cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(Fonts.GetUIFontName(), vbBinaryCompare)
        
        'OpenType-specific features are a big investment, so I've postponed them to a later date
        'If OS.IsVistaOrLater Then btsCharCategory.AddItem "OpenType", 2
        
        'Fill AA options
        cboTextRenderingHint.SetAutomaticRedraws False
        cboTextRenderingHint.Clear
        cboTextRenderingHint.AddItem "none", 0
        cboTextRenderingHint.AddItem "normal", 1
        cboTextRenderingHint.AddItem "crisp", 2
        cboTextRenderingHint.AddItem "smooth", 3
        cboTextRenderingHint.ListIndex = 1
        cboTextRenderingHint.SetAutomaticRedraws True
        
        'Add dummy entries to the various alignment buttons; we'll populate these with theme-specific
        ' images in the UpdateAgainstCurrentTheme() function.
        btsHAlignment.AddItem vbNullString, 0
        btsHAlignment.AddItem vbNullString, 1
        btsHAlignment.AddItem vbNullString, 2
        btsHAlignment.AddItem vbNullString, 3
        
        btsVAlignment.AddItem vbNullString, 0
        btsVAlignment.AddItem vbNullString, 1
        btsVAlignment.AddItem vbNullString, 2
        
        btsHAlignJustify.AddItem vbNullString, 0
        btsHAlignJustify.AddItem vbNullString, 1
        btsHAlignJustify.AddItem vbNullString, 2
        btsHAlignJustify.AddItem vbNullString, 3
        
        'Fill various character positioning settings
        btsStretch.AddItem "none", 0
        btsStretch.AddItem "box", 1
        btsStretch.ListIndex = 0
        
        cboCharMirror.SetAutomaticRedraws False
        cboCharMirror.Clear
        cboCharMirror.AddItem "none", 0
        cboCharMirror.AddItem "horizontal", 1
        cboCharMirror.AddItem "vertical", 2
        cboCharMirror.AddItem "both", 3
        cboCharMirror.ListIndex = 0
        cboCharMirror.SetAutomaticRedraws True
        
        cboCharCase.SetAutomaticRedraws False
        cboCharCase.Clear
        cboCharCase.AddItem "none", 0
        cboCharCase.AddItem "lowercase", 1
        cboCharCase.AddItem "UPPERCASE", 2
        cboCharCase.AddItem "hiragana", 3
        cboCharCase.AddItem "katakana", 4
        cboCharCase.AddItem "simplified Chinese", 5
        cboCharCase.AddItem "traditional Chinese", 6
        If OS.IsWin7OrLater Then cboCharCase.AddItem "Titlecase", 7
        cboCharCase.ListIndex = 0
        cboCharCase.SetAutomaticRedraws True
        
        btsHinting.AddItem "off", 0
        btsHinting.AddItem "on", 1
        btsHinting.ListIndex = 1
        
        'Fill wordwrap options
        cboWordWrap.SetAutomaticRedraws False
        cboWordWrap.Clear
        cboWordWrap.AddItem "none", 0
        cboWordWrap.AddItem "manual only", 1
        cboWordWrap.AddItem "characters", 2
        cboWordWrap.AddItem "words", 3
        cboWordWrap.ListIndex = 3
        cboWordWrap.SetAutomaticRedraws True
        
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

Private Sub hypEditStyles_Click()
    UpdateFlyout 2, True
End Sub

Private Sub hypEditStyles_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub hypEditStyles_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(2).hWnd
    Else
        newTargetHwnd = Me.chkFillText.hWnd
    End If
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

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub psText_PenChanged(ByVal isFinalChange As Boolean)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_OutlinePen, psText.Pen
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub psText_GotFocusAPI()
    UpdateFlyout 2, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_OutlinePen, psText.Pen, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub psText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_OutlinePen, psText.Pen
End Sub

Private Sub psText_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkOutlineText.hWnd
    Else
        newTargetHwnd = Me.chkFillFirst.hWnd
    End If
End Sub

Private Sub psTextBackground_PenChanged(ByVal isFinalChange As Boolean)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_BackBorderPen, psTextBackground.Pen
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub psTextBackground_GotFocusAPI()
    UpdateFlyout 2, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackBorderPen, psTextBackground.Pen, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub psTextBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackBorderPen, psTextBackground.Pen
End Sub

Private Sub psTextBackground_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkBackgroundBorder.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    End If
End Sub

Private Sub sldLineSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsHAlignJustify.hWnd
    Else
        newTargetHwnd = Me.cboWordWrap.hWnd
    End If
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
        newTargetHwnd = Me.btsStretch.hWnd
    End If
End Sub

Private Sub sltCharInflation_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharInflation, sltCharInflation.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sltCharInflation_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharInflation, sltCharInflation.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltCharInflation_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharInflation, sltCharInflation.Value
End Sub

Private Sub sltCharInflation_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.tudJitter(1).hWnd
    Else
        newTargetHwnd = Me.cboCharMirror.hWnd
    End If
End Sub

Private Sub sltCharOrientation_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharOrientation, sltCharOrientation.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sltCharOrientation_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharOrientation, sltCharOrientation.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltCharOrientation_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharOrientation, sltCharOrientation.Value
End Sub

Private Sub sltCharOrientation_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboCharCase.hWnd
    Else
        newTargetHwnd = Me.tudJitter(0).hWnd
    End If
End Sub

Private Sub sltCharSpacing_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharSpacing, sltCharSpacing.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sltCharSpacing_GotFocusAPI()
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharSpacing, sltCharSpacing.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltCharSpacing_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharSpacing, sltCharSpacing.Value
End Sub

Private Sub sltCharSpacing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btnFontStyles(3).hWnd
    Else
        newTargetHwnd = Me.cboTextRenderingHint.hWnd
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        
        Dim newIndex As Long
        newIndex = Index - 1
        If (newIndex < 0) Then newIndex = newIndex + 4
        
        newTargetHwnd = Me.cmdFlyoutLock(newIndex).hWnd
    
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.hypEditText.hWnd
            Case 1
                newTargetHwnd = Me.cboTextFontFace.hWnd
            Case 2
                newTargetHwnd = Me.hypEditStyles.hWnd
            Case 3
                newTargetHwnd = Me.btsHAlignment.hWnd
        End Select
    End If
End Sub

Private Sub tudJitter_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_CharJitterX + Index, tudJitter(Index).Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub tudJitter_GotFocusAPI(Index As Integer)
    UpdateFlyout 1, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharJitterX + Index, tudJitter(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub tudJitter_LostFocusAPI(Index As Integer)
    Processor.FlagFinalNDFXState_Text ptp_CharJitterX + Index, tudJitter(Index).Value
End Sub

Private Sub sldLineSpacing_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the setting
    PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_LineSpacing, sldLineSpacing.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub sldLineSpacing_GotFocusAPI()
    UpdateFlyout 3, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_LineSpacing, sldLineSpacing.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sldLineSpacing_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_LineSpacing, sldLineSpacing.Value
End Sub

Private Sub tudJitter_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then
            newTargetHwnd = Me.sltCharOrientation.hWndSpinner
        Else
            newTargetHwnd = Me.tudJitter(1).hWnd
        End If
    Else
        If shiftTabWasPressed Then
            newTargetHwnd = Me.tudJitter(0).hWnd
        Else
            newTargetHwnd = Me.sltCharInflation.hWnd
        End If
    End If
End Sub

Private Sub tudMargin_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Or m_suspendSettingRelay Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current setting
    Select Case Index
    
        Case 0
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_MarginLeft, tudMargin(Index).Value
        
        Case 1
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_MarginRight, tudMargin(Index).Value
        
        Case 2
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_MarginTop, tudMargin(Index).Value
        
        Case 3
            PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_MarginBottom, tudMargin(Index).Value
    
    End Select
        
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub tudMargin_GotFocusAPI(Index As Integer)
    
    UpdateFlyout 3, True
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    Select Case Index
    
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_MarginLeft, tudMargin(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_MarginRight, tudMargin(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_MarginTop, tudMargin(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_MarginBottom, tudMargin(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
        
    End Select
    
End Sub

Private Sub tudMargin_LostFocusAPI(Index As Integer)
    
    Select Case Index
    
        Case 0
            Processor.FlagFinalNDFXState_Text ptp_MarginLeft, tudMargin(Index).Value
        
        Case 1
            Processor.FlagFinalNDFXState_Text ptp_MarginRight, tudMargin(Index).Value
        
        Case 2
            Processor.FlagFinalNDFXState_Text ptp_MarginTop, tudMargin(Index).Value
        
        Case 3
            Processor.FlagFinalNDFXState_Text ptp_MarginBottom, tudMargin(Index).Value
    
    End Select
        
End Sub

Private Sub tudMargin_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If (Index = 0) Then
            newTargetHwnd = Me.cboWordWrap.hWnd
        Else
            newTargetHwnd = Me.tudMargin(Index - 1).hWnd
        End If
    Else
        If (Index = 3) Then
            newTargetHwnd = Me.cmdFlyoutLock(3).hWnd
        Else
            newTargetHwnd = Me.tudMargin(Index + 1).hWnd
        End If
    End If
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

'Outside functions can forcibly request an update against the current layer.  If the current layer is
' a basic text layer, an option will be displayed to convert the layer to advanced text.
Public Sub UpdateAgainstCurrentLayer()
    
    If PDImages.IsImageActive() Then
    
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then
        
            'Check for non-advanced-text layers.
            If (PDImages.GetActiveImage.GetActiveLayer.GetLayerType <> PDL_TextAdvanced) Then
            
                Select Case PDImages.GetActiveImage.GetActiveLayer.GetLayerType
                
                    Case PDL_TextBasic
                        Dim newMessage As String
                        newMessage = g_Language.TranslateMessage("This is a basic text layer.  To edit it with the advanced text tool, you must first convert it to an advanced text layer.")
                        newMessage = newMessage & Space$(2) & g_Language.TranslateMessage("(This action is non-destructive.)")
                        Me.lblConvertLayer.Caption = newMessage
                        
                    'In the future, other text layer types can be added here.
                
                End Select
            
                Me.cmdConvertLayer.Caption = g_Language.TranslateMessage("Click here to convert this layer to advanced text.")
                
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

'When a new text layer is created, the user can choose to auto-drop the text entry panel.
Public Sub NotifyNewLayerCreated()
    If Me.chkAutoOpenText.Value Then
        UpdateFlyout 0, True
        Me.txtTextTool.SetFocusToEditBox True
    End If
End Sub

Public Sub SyncSettingsToCurrentLayer()

    txtTextTool.Text = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_Text)
    cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontFace), vbTextCompare)
    sldTextFontSize.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontSize)
    btsStretch.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_StretchToFit)
    cboTextRenderingHint.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextAntialiasing)
    If PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextHinting) Then btsHinting.ListIndex = 1 Else btsHinting.ListIndex = 0
    btnFontStyles(0).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontBold))
    btnFontStyles(1).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontItalic))
    btnFontStyles(2).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontUnderline))
    btnFontStyles(3).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontStrikeout))
    btsHAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_HorizontalAlignment)
    btsVAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_VerticalAlignment)
    btsHAlignJustify.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_AlignLastLine)
    cboWordWrap.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_WordWrap)
    chkFillText.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FillActive)
    bsText.Brush = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FillBrush)
    chkOutlineText.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_OutlineActive)
    psText.Pen = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_OutlinePen)
    chkBackground.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_BackgroundActive)
    bsTextBackground.Brush = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_BackgroundBrush)
    chkBackgroundBorder.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_BackBorderActive)
    psTextBackground.Pen = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_BackBorderPen)
    tudMargin(0).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_MarginLeft)
    tudMargin(1).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_MarginRight)
    tudMargin(2).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_MarginTop)
    tudMargin(3).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_MarginBottom)
    sldLineSpacing.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_LineSpacing)
    sltCharInflation.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharInflation)
    tudJitter(0).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharJitterX)
    tudJitter(1).Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharJitterY)
    cboCharMirror.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharMirror)
    sltCharOrientation.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharOrientation)
    cboCharCase.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharRemap)
    sltCharSpacing.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_CharSpacing)
    chkFillFirst.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_OutlineAboveFill)

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
    btsHAlignment.AssignImageToItem 3, "format_alignjustify", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    btsVAlignment.AssignImageToItem 0, "format_aligntop", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsVAlignment.AssignImageToItem 1, "format_alignmiddle", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsVAlignment.AssignImageToItem 2, "format_alignbottom", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    btsHAlignJustify.AssignImageToItem 0, "format_alignleft", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsHAlignJustify.AssignImageToItem 1, "format_aligncenter", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsHAlignJustify.AssignImageToItem 2, "format_alignright", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    btsHAlignJustify.AssignImageToItem 3, "format_alignjustify", , buttonSize, buttonSize, usePDResamplerInstead:=rf_Box
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
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
