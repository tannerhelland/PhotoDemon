VERSION 5.00
Begin VB.Form toolpanel_FancyText 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18435
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
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1229
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer pdcMain 
      Height          =   1500
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   2646
      Begin PhotoDemon.pdTextBox txtTextTool 
         Height          =   1380
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2434
         FontSize        =   9
         Multiline       =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupConvertLayer 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdHyperlink lblConvertLayerConfirm 
         Height          =   240
         Left            =   120
         Top             =   900
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   423
         Alignment       =   2
         Caption         =   "yes"
         Layout          =   2
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblConvertLayer 
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   1296
         Alignment       =   2
         Caption         =   "yes"
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdButtonStripVertical btsMain 
      Height          =   1380
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   2434
   End
   Begin PhotoDemon.pdContainer pdcMain 
      Height          =   1500
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   2646
      Begin PhotoDemon.pdButtonStripVertical btsCategory 
         Height          =   1380
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   2175
         _ExtentX        =   4048
         _ExtentY        =   2434
      End
      Begin PhotoDemon.pdContainer ctlGroupCategory 
         Height          =   1500
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   0
         Width           =   10935
         _ExtentX        =   0
         _ExtentY        =   0
         Begin PhotoDemon.pdButtonStripVertical btsCharCategory 
            Height          =   1380
            Left            =   0
            TabIndex        =   7
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2434
         End
         Begin PhotoDemon.pdContainer ctlGroupCharCategory 
            Height          =   1500
            Index           =   0
            Left            =   1920
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   0
            _ExtentY        =   0
            Begin PhotoDemon.pdDropDownFont cboTextFontFace 
               Height          =   375
               Left            =   1320
               TabIndex        =   9
               Top             =   0
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
            End
            Begin PhotoDemon.pdSpinner tudTextFontSize 
               Height          =   345
               Left            =   1320
               TabIndex        =   10
               Top             =   450
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   609
               DefaultValue    =   16
               Min             =   1
               Max             =   1000
               Value           =   16
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   3
               Left            =   0
               Top             =   60
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "font face:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   4
               Left            =   0
               Top             =   510
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "font size:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   2
               Left            =   0
               Top             =   960
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "font style:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdButtonToolbox btnFontStyles 
               Height          =   435
               Index           =   1
               Left            =   1800
               TabIndex        =   11
               Top             =   870
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   767
               StickyToggle    =   -1  'True
            End
            Begin PhotoDemon.pdButtonToolbox btnFontStyles 
               Height          =   435
               Index           =   2
               Left            =   2280
               TabIndex        =   12
               Top             =   870
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   767
               StickyToggle    =   -1  'True
            End
            Begin PhotoDemon.pdButtonToolbox btnFontStyles 
               Height          =   435
               Index           =   3
               Left            =   2760
               TabIndex        =   13
               Top             =   870
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   767
               StickyToggle    =   -1  'True
            End
            Begin PhotoDemon.pdCheckBox chkHinting 
               Height          =   330
               Left            =   4200
               TabIndex        =   14
               Top             =   450
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   582
               Caption         =   "hinting"
               Value           =   0
            End
            Begin PhotoDemon.pdDropDown cboTextRenderingHint 
               Height          =   375
               Left            =   5400
               TabIndex        =   15
               Top             =   0
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   635
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   5
               Left            =   3840
               Top             =   60
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "antialiasing:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdButtonToolbox btnFontStyles 
               Height          =   435
               Index           =   0
               Left            =   1320
               TabIndex        =   16
               Top             =   870
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   767
               StickyToggle    =   -1  'True
            End
         End
         Begin PhotoDemon.pdContainer ctlGroupCharCategory 
            Height          =   1500
            Index           =   1
            Left            =   1920
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   0
            _ExtentY        =   0
            Begin PhotoDemon.pdSpinner tudJitter 
               Height          =   345
               Index           =   0
               Left            =   5280
               TabIndex        =   18
               Top             =   0
               Width           =   1215
               _ExtentX        =   1720
               _ExtentY        =   609
               Max             =   100
               SigDigits       =   1
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   26
               Left            =   0
               Top             =   60
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "remap:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdDropDown cboCharCase 
               Height          =   375
               Left            =   1320
               TabIndex        =   19
               Top             =   0
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   661
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   30
               Left            =   0
               Top             =   540
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "spacing:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdSlider sltCharSpacing 
               CausesValidation=   0   'False
               Height          =   405
               Left            =   1200
               TabIndex        =   20
               Top             =   420
               Width           =   2760
               _ExtentX        =   4868
               _ExtentY        =   873
               Min             =   -1
               Max             =   1
               SigDigits       =   3
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   31
               Left            =   0
               Top             =   1020
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "orientation:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdSlider sltCharOrientation 
               CausesValidation=   0   'False
               Height          =   405
               Left            =   1200
               TabIndex        =   21
               Top             =   900
               Width           =   2760
               _ExtentX        =   4868
               _ExtentY        =   873
               Min             =   -360
               Max             =   360
               SigDigits       =   1
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   32
               Left            =   3960
               Top             =   60
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "jitter:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   33
               Left            =   4125
               Top             =   1020
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "mirror:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdDropDown cboCharMirror 
               Height          =   375
               Left            =   5280
               TabIndex        =   22
               Top             =   945
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   661
            End
            Begin PhotoDemon.pdSpinner tudJitter 
               Height          =   345
               Index           =   1
               Left            =   6675
               TabIndex        =   23
               Top             =   0
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   609
               Max             =   100
               SigDigits       =   1
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   34
               Left            =   3960
               Top             =   540
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "inflation:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdSlider sltCharInflation 
               CausesValidation=   0   'False
               Height          =   405
               Left            =   5160
               TabIndex        =   24
               Top             =   420
               Width           =   2760
               _ExtentX        =   4868
               _ExtentY        =   873
               Max             =   20
               SigDigits       =   1
            End
         End
      End
      Begin PhotoDemon.pdContainer ctlGroupCategory 
         Height          =   1500
         Index           =   3
         Left            =   2280
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdContainer ctlGroupCategory 
         Height          =   1500
         Index           =   2
         Left            =   2280
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   0
         _ExtentY        =   0
         Begin PhotoDemon.pdButtonStripVertical btsAppearanceCategory 
            Height          =   1380
            Left            =   0
            TabIndex        =   27
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2434
         End
         Begin PhotoDemon.pdContainer ctlGroupAppearanceCategory 
            Height          =   1500
            Index           =   1
            Left            =   1920
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   0
            _ExtentY        =   0
            Begin PhotoDemon.pdPenSelector psTextBackground 
               Height          =   855
               Left            =   4680
               TabIndex        =   29
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   1508
            End
            Begin PhotoDemon.pdBrushSelector bsTextBackground 
               Height          =   855
               Left            =   1200
               TabIndex        =   30
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   1508
            End
            Begin PhotoDemon.pdCheckBox chkBackground 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   60
               Width           =   3240
               _ExtentX        =   4445
               _ExtentY        =   582
               Caption         =   "fill background"
               Value           =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   720
               Index           =   15
               Left            =   0
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   1270
               Alignment       =   1
               Caption         =   "fill style:"
               ForeColor       =   0
               Layout          =   1
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   720
               Index           =   28
               Left            =   3360
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1270
               Alignment       =   1
               Caption         =   "border style:"
               ForeColor       =   0
               Layout          =   1
            End
            Begin PhotoDemon.pdCheckBox chkBackgroundBorder 
               Height          =   330
               Left            =   3480
               TabIndex        =   32
               Top             =   60
               Width           =   3240
               _ExtentX        =   4445
               _ExtentY        =   582
               Caption         =   "background border"
               Value           =   0
            End
         End
         Begin PhotoDemon.pdContainer ctlGroupAppearanceCategory 
            Height          =   1500
            Index           =   0
            Left            =   1920
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   0
            _ExtentY        =   0
            Begin PhotoDemon.pdPenSelector psText 
               Height          =   855
               Left            =   4680
               TabIndex        =   34
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   1508
            End
            Begin PhotoDemon.pdBrushSelector bsText 
               Height          =   855
               Left            =   1200
               TabIndex        =   35
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   1508
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   855
               Index           =   6
               Left            =   0
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   1508
               Alignment       =   1
               Caption         =   "fill style:"
               Layout          =   1
            End
            Begin PhotoDemon.pdCheckBox chkFillText 
               Height          =   330
               Left            =   120
               TabIndex        =   36
               Top             =   60
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   582
               Caption         =   "fill text"
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   855
               Index           =   7
               Left            =   3360
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1508
               Alignment       =   1
               Caption         =   "outline style:"
               Layout          =   1
            End
            Begin PhotoDemon.pdCheckBox chkOutlineText 
               Height          =   330
               Left            =   3480
               TabIndex        =   37
               Top             =   60
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   582
               Caption         =   "outline text"
               Value           =   0
            End
         End
      End
      Begin PhotoDemon.pdContainer ctlGroupCategory 
         Height          =   1500
         Index           =   1
         Left            =   2280
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   0
         _ExtentY        =   0
         Begin PhotoDemon.pdSpinner tudLineSpacing 
            Height          =   345
            Left            =   5160
            TabIndex        =   39
            Top             =   1020
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Min             =   -10
            SigDigits       =   2
         End
         Begin PhotoDemon.pdSpinner tudMargin 
            Height          =   345
            Index           =   0
            Left            =   5160
            TabIndex        =   40
            Top             =   90
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            Min             =   -1000
            Max             =   1000
         End
         Begin PhotoDemon.pdButtonStrip btsHAlignment 
            Height          =   435
            Left            =   1320
            TabIndex        =   41
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   767
            ColorScheme     =   1
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   8
            Left            =   0
            Top             =   150
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "alignment:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdButtonStrip btsVAlignment 
            Height          =   435
            Left            =   1320
            TabIndex        =   42
            Top             =   510
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   767
            ColorScheme     =   1
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   0
            Left            =   0
            Top             =   1080
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "line wrap:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdDropDown cboWordWrap 
            Height          =   375
            Left            =   1320
            TabIndex        =   43
            Top             =   1020
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   661
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   23
            Left            =   3360
            Top             =   150
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "h. padding:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdSpinner tudMargin 
            Height          =   345
            Index           =   1
            Left            =   6540
            TabIndex        =   44
            Top             =   90
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            Min             =   -1000
            Max             =   1000
         End
         Begin PhotoDemon.pdSpinner tudMargin 
            Height          =   345
            Index           =   2
            Left            =   5160
            TabIndex        =   45
            Top             =   570
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            Min             =   -1000
            Max             =   1000
         End
         Begin PhotoDemon.pdSpinner tudMargin 
            Height          =   345
            Index           =   3
            Left            =   6540
            TabIndex        =   46
            Top             =   570
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            Min             =   -1000
            Max             =   1000
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   24
            Left            =   3360
            Top             =   630
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "v. padding:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   25
            Left            =   3480
            Top             =   1080
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "line spacing:"
            ForeColor       =   0
         End
      End
   End
End
Attribute VB_Name = "toolpanel_FancyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Advanced Typography Tool Panel
'Copyright 2013-2017 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 22/June/17
'Last update: large improvements to the way non-destructive actions interact with the Undo/Redo engine
'
'This form includes all user-editable settings for PD's Advanced Typography text tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub bsText_BrushChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FillBrush, bsText.Brush
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub bsText_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FillBrush, bsText.Brush, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub bsText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FillBrush, bsText.Brush
End Sub

Private Sub bsTextBackground_BrushChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_BackgroundBrush, bsTextBackground.Brush
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub bsTextBackground_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackgroundBrush, bsTextBackground.Brush, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub bsTextBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackgroundBrush, bsTextBackground.Brush
End Sub

Private Sub btnFontStyles_Click(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    'Non-destructive effects are obviously not tracked if no images are loaded
    If (g_OpenImageCount = 0) Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
            
        'Italic
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        'Underline
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        'Strikeout
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
    
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

Private Sub btsAppearanceCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsAppearanceCategory.ListCount - 1
        ctlGroupAppearanceCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        ctlGroupCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsCharCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsCharCategory.ListCount - 1
        ctlGroupCharCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsMain_Click(ByVal buttonIndex As Long)
    ChangeMainPanel
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub cboCharCase_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharRemap, cboCharCase.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboCharCase_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharRemap, cboCharCase.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboCharCase_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharRemap, cboCharCase.ListIndex
End Sub

Private Sub cboCharMirror_Click()

    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharMirror, cboCharMirror.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub cboCharMirror_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharMirror, cboCharMirror.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboCharMirror_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharMirror, cboCharMirror.ListIndex
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextRenderingHint_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub cboWordWrap_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_WordWrap, cboWordWrap.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboWordWrap_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub cboWordWrap_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex
End Sub

Private Sub chkBackground_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_BackgroundActive, CBool(chkBackground.Value)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkBackground_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackgroundActive, CBool(chkBackground.Value), pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub chkBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackgroundActive, CBool(chkBackground.Value)
End Sub

Private Sub chkBackgroundBorder_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_BackBorderActive, CBool(chkBackgroundBorder.Value)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkBackgroundBorder_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackBorderActive, CBool(chkBackgroundBorder.Value), pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub chkBackgroundBorder_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackBorderActive, CBool(chkBackgroundBorder.Value)
End Sub

Private Sub chkFillText_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FillActive, CBool(chkFillText.Value)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkFillText_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FillActive, chkFillText.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub chkFillText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FillActive, chkFillText.Value
End Sub

Private Sub chkHinting_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_TextHinting, CBool(chkHinting.Value)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkHinting_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value), pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub chkHinting_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value)
End Sub

Private Sub chkOutlineText_Click()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_OutlineActive, CBool(chkOutlineText.Value)
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkOutlineText_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_OutlineActive, chkOutlineText.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub chkOutlineText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_OutlineActive, chkOutlineText.Value
End Sub

Private Sub Form_Load()
    
    'Disable any layer updates as a result of control changes during the load process
    Tools.SetToolBusyState True
    
    'Generate a list of fonts
    If MainModule.IsProgramRunning() Then
        
        'This tool is separated into two panels: text entry, and text settings
        btsMain.AddItem "text", 0
        btsMain.AddItem "settings", 1
        btsMain.ListIndex = 0
        ChangeMainPanel
        
        'Initialize the font list
        cboTextFontFace.InitializeFontList
        cboTextFontFace.ListIndex = cboTextFontFace.ListIndexByString(g_InterfaceFont, vbBinaryCompare)
        
        'Draw the primary category selector
        btsCategory.AddItem "character", 0
        btsCategory.AddItem "paragraph", 1
        btsCategory.AddItem "visual", 2
        
        'I've already stubbed out a 4th options panel, but the vertical button list is *really* cramped, so another solution might be necessary
        
        'Draw the character sub-category selector
        btsCharCategory.AddItem "font", 0
        btsCharCategory.AddItem "special", 1
        btsCharCategory.ListIndex = 0
        
        'OpenType-specific features are a big investment, so I've postponed them to a later date
        'If OS.IsVistaOrLater Then btsCharCategory.AddItem "OpenType", 2
        
        'Fill AA options
        cboTextRenderingHint.Clear
        cboTextRenderingHint.AddItem "none", 0
        cboTextRenderingHint.AddItem "normal", 1
        cboTextRenderingHint.AddItem "crisp", 2
        cboTextRenderingHint.ListIndex = 1
        
        'Add dummy entries to the various alignment buttons; we'll populate these with theme-specific
        ' images in the UpdateAgainstCurrentTheme() function.
        btsHAlignment.AddItem vbNullString, 0
        btsHAlignment.AddItem vbNullString, 1
        btsHAlignment.AddItem vbNullString, 2
        
        btsVAlignment.AddItem vbNullString, 0
        btsVAlignment.AddItem vbNullString, 1
        btsVAlignment.AddItem vbNullString, 2
        
        'Fill various character positioning settings
        cboCharMirror.Clear
        cboCharMirror.AddItem "none", 0
        cboCharMirror.AddItem "horizontal", 1
        cboCharMirror.AddItem "vertical", 2
        cboCharMirror.AddItem "both", 3
        cboCharMirror.ListIndex = 0
        
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
        
        'Fill wordwrap options
        cboWordWrap.Clear
        cboWordWrap.AddItem "none", 0
        cboWordWrap.AddItem "manual only", 1
        cboWordWrap.AddItem "characters", 2
        cboWordWrap.AddItem "words", 3
        cboWordWrap.ListIndex = 3
        
        'Draw the appearance sub-category selector
        btsAppearanceCategory.AddItem "text", 0
        btsAppearanceCategory.AddItem "background", 1
        btsAppearanceCategory.ListIndex = 0
        
        'Load any last-used settings for this form
        Set lastUsedSettings = New pdLastUsedSettings
        lastUsedSettings.SetParentForm Me
        lastUsedSettings.LoadAllControlValues
        
        'Update everything against the current theme.  This will also set tooltips for various controls.
        UpdateAgainstCurrentTheme
        
    End If
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    UpdateAgainstCurrentLayer
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Make sure the correct panels are shown
    btsCategory_Click btsCategory.ListIndex
    btsAppearanceCategory_Click btsAppearanceCategory.ListIndex
    btsCharCategory_Click btsCharCategory.ListIndex

End Sub

Private Sub lblConvertLayerConfirm_Click()
    
    'Because of the way this warning panel is constructed, this label will not be visible unless a click is valid.
    pdImages(g_CurrentImage).GetActiveLayer.SetLayerType PDL_TYPOGRAPHY
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).GetActiveLayerIndex
    
    'Hide the warning panel and redraw both the viewport, and the UI (as new UI options may now be available)
    Me.UpdateAgainstCurrentLayer
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    Interface.SyncInterfaceToCurrentImage
    
End Sub

Private Sub pdcMain_SizeChanged(Index As Integer)

    'The "text" panel auto-resizes the text entry area to match the size of the container
    If (Index = 0) Then
        txtTextTool.SetSize (pdcMain(Index).GetWidth - txtTextTool.GetLeft) - FixDPI(4), txtTextTool.GetHeight
    End If
    
End Sub

Private Sub psText_PenChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_OutlinePen, psText.Pen
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub psText_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_OutlinePen, psText.Pen, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub psText_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_OutlinePen, psText.Pen
End Sub

Private Sub psTextBackground_PenChanged()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_BackBorderPen, psTextBackground.Pen
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub psTextBackground_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_BackBorderPen, psTextBackground.Pen, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub psTextBackground_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_BackBorderPen, psTextBackground.Pen
End Sub

Private Sub sltCharInflation_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharInflation, sltCharInflation.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub sltCharInflation_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharInflation, sltCharInflation.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltCharInflation_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharInflation, sltCharInflation.Value
End Sub

Private Sub sltCharOrientation_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharOrientation, sltCharOrientation.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub sltCharOrientation_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharOrientation, sltCharOrientation.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltCharOrientation_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharOrientation, sltCharOrientation.Value
End Sub

Private Sub sltCharSpacing_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharSpacing, sltCharSpacing.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub sltCharSpacing_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharSpacing, sltCharSpacing.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub sltCharSpacing_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_CharSpacing, sltCharSpacing.Value
End Sub

Private Sub tudJitter_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_CharJitterX + Index, tudJitter(Index).Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudJitter_GotFocusAPI(Index As Integer)
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_CharJitterX + Index, tudJitter(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub tudJitter_LostFocusAPI(Index As Integer)
    Processor.FlagFinalNDFXState_Text ptp_CharJitterX + Index, tudJitter(Index).Value
End Sub

Private Sub tudLineSpacing_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the setting
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_LineSpacing, tudLineSpacing.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudLineSpacing_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_LineSpacing, tudLineSpacing.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub tudLineSpacing_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_LineSpacing, tudLineSpacing.Value
End Sub

Private Sub tudMargin_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current setting
    Select Case Index
    
        Case 0
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_MarginLeft, tudMargin(Index).Value
        
        Case 1
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_MarginRight, tudMargin(Index).Value
        
        Case 2
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_MarginTop, tudMargin(Index).Value
        
        Case 3
            pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_MarginBottom, tudMargin(Index).Value
    
    End Select
        
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudMargin_GotFocusAPI(Index As Integer)

    If (g_OpenImageCount = 0) Then Exit Sub
    
    Select Case Index
    
        Case 0
            Processor.FlagInitialNDFXState_Text ptp_MarginLeft, tudMargin(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        Case 1
            Processor.FlagInitialNDFXState_Text ptp_MarginRight, tudMargin(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        Case 2
            Processor.FlagInitialNDFXState_Text ptp_MarginTop, tudMargin(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
        Case 3
            Processor.FlagInitialNDFXState_Text ptp_MarginBottom, tudMargin(Index).Value, pdImages(g_CurrentImage).GetActiveLayerID
        
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

Private Sub tudTextFontSize_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_FontSize, tudTextFontSize.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudTextFontSize_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_FontSize, tudTextFontSize.Value, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub tudTextFontSize_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_FontSize, tudTextFontSize.Value
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.  (Note that this also queries Tools.GetToolBusyState)
    If (Not Tools.CanvasToolsAllowed) Or (Not CurrentLayerIsText) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
End Sub

Private Sub txtTextTool_GotFocusAPI()
    If (g_OpenImageCount = 0) Then Exit Sub
    Processor.FlagInitialNDFXState_Text ptp_Text, txtTextTool.Text, pdImages(g_CurrentImage).GetActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    Processor.FlagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub

'Most objects on this form can avoid doing any work if the current layer is not a text layer.
Private Function CurrentLayerIsText() As Boolean
    
    CurrentLayerIsText = False
    
    'Changing UI elements does nothing if no images are loaded
    If (g_OpenImageCount = 0) Then Exit Function
    
    If (Not pdImages(g_CurrentImage) Is Nothing) Then
        If (Not pdImages(g_CurrentImage).GetActiveLayer Is Nothing) Then
            CurrentLayerIsText = pdImages(g_CurrentImage).GetActiveLayer.IsLayerText
        End If
    End If
    
End Function

Private Sub ChangeMainPanel()
    
    Dim i As Long
    For i = pdcMain.lBound To pdcMain.UBound
        pdcMain(i).Visible = (i = btsMain.ListIndex)
    Next i
    
End Sub

'Outside functions can forcibly request an update against the current layer.  If the current layer is a non-typography text layer of
' some type (e.g. basic text), an option will be displayed to convert the layer.
Public Sub UpdateAgainstCurrentLayer()
    
    'Regardless of layer type, resize our containers to match the current window width.
    Dim winSize As winRect
    If (Not g_WindowManager Is Nothing) Then
        
        g_WindowManager.GetClientWinRect Me.hWnd, winSize
        
        Dim i As Long
        For i = pdcMain.lBound To pdcMain.UBound
            pdcMain(i).SetSize (winSize.x2 - winSize.x1) - pdcMain(i).GetLeft, pdcMain(i).GetHeight
        Next i
        
    End If
    
    If (g_OpenImageCount > 0) Then
    
        If pdImages(g_CurrentImage).GetActiveLayer.IsLayerText Then
        
            'Check for non-typography layers.
            If pdImages(g_CurrentImage).GetActiveLayer.GetLayerType <> PDL_TYPOGRAPHY Then
            
                Select Case pdImages(g_CurrentImage).GetActiveLayer.GetLayerType
                
                    Case PDL_TEXT
                        Dim newMessage As String
                        newMessage = g_Language.TranslateMessage("This layer is a basic text layer.  To edit it with the typography tool, you must first convert it to a typography layer.")
                        newMessage = newMessage & vbCrLf & g_Language.TranslateMessage("(This action is non-destructive.)")
                        Me.lblConvertLayer.Caption = newMessage
                        
                    'In the future, other text layer types can be added here.
                
                End Select
            
                Me.lblConvertLayerConfirm.Caption = g_Language.TranslateMessage("Click here to convert this layer to a typography layer.")
                
                'Make the prompt panel the size of the tool window
                Me.ctlGroupConvertLayer.SetPositionAndSize 0, 0, Me.ScaleWidth, Me.ScaleHeight
                
                'Center all labels on the panel
                Me.lblConvertLayer.SetLeft (Me.ScaleWidth - lblConvertLayer.GetWidth) / 2
                Me.lblConvertLayerConfirm.SetLeft (Me.ScaleWidth - lblConvertLayerConfirm.GetWidth) / 2
                
                'Display the panel
                Me.ctlGroupConvertLayer.Visible = True
                Me.ctlGroupConvertLayer.ZOrder 0
                Me.ctlGroupConvertLayer.Refresh
                
            Else
                Me.ctlGroupConvertLayer.Visible = False
            End If
        
        Else
            Me.ctlGroupConvertLayer.Visible = False
        End If
        
    Else
        Me.ctlGroupConvertLayer.Visible = False
    End If

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update any UI images against the current theme
    Dim buttonSize As Long
    buttonSize = FixDPI(24)
    
    btnFontStyles(0).AssignImage "format_bold", , buttonSize, buttonSize
    btnFontStyles(1).AssignImage "format_italic", , buttonSize, buttonSize
    btnFontStyles(2).AssignImage "format_underline", , buttonSize, buttonSize
    btnFontStyles(3).AssignImage "format_strikethrough", , buttonSize, buttonSize
    
    btsHAlignment.AssignImageToItem 0, "format_alignleft", , buttonSize, buttonSize
    btsHAlignment.AssignImageToItem 1, "format_aligncenter", , buttonSize, buttonSize
    btsHAlignment.AssignImageToItem 2, "format_alignright", , buttonSize, buttonSize
    
    btsVAlignment.AssignImageToItem 0, "format_aligntop", , buttonSize, buttonSize
    btsVAlignment.AssignImageToItem 1, "format_alignmiddle", , buttonSize, buttonSize
    btsVAlignment.AssignImageToItem 2, "format_alignbottom", , buttonSize, buttonSize
        
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    Interface.ApplyThemeAndTranslations Me

End Sub
