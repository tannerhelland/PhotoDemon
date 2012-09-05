VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13845
   Icon            =   "VBP_FormMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProgBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   923
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7905
      Width           =   13845
   End
   Begin PhotoDemon.vbalHookControl ctlAccelerator 
      Left            =   12240
      Top             =   6960
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7905
      Left            =   0
      ScaleHeight     =   525
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2235
      Begin VB.ComboBox CmbZoom 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "VBP_FormMain.frx":058A
         Left            =   840
         List            =   "VBP_FormMain.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Click to adjust image zoom"
         Top             =   2610
         Width           =   1215
      End
      Begin PhotoDemon.jcbutton cmdOpen 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Open Image"
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":058E
         DisabledPictureMode=   1
         CaptionEffects  =   0
      End
      Begin PhotoDemon.jcbutton cmdSave 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save Image"
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":15E0
         DisabledPictureMode=   1
         CaptionEffects  =   0
      End
      Begin PhotoDemon.jcbutton cmdUndo 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":2632
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Undo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdRedo 
         Height          =   615
         Left            =   1155
         TabIndex        =   4
         Top             =   1680
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":3684
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Redo"
         TooltipBackColor=   -2147483643
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   232
         Y2              =   232
      End
      Begin VB.Label lblRecording 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RECORDING IN PROGRESS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   3720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblCoordinates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(X, Y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   7560
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblImgSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Size: WidthxHeight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D1B499&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   1845
      End
      Begin VB.Label lblZoom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00544E43&
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   2640
         Width           =   555
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   104
         Y2              =   104
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuRecent 
         Caption         =   "Open &Recent..."
         Begin VB.Menu mnuRecDocs 
            Caption         =   "Empty"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MnuRecentSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuClearMRU 
            Caption         =   "Clear recent image list"
         End
      End
      Begin VB.Menu MnuAcquire 
         Caption         =   "&Import..."
         Begin VB.Menu MnuScanImage 
            Caption         =   "From Scanner/Camera..."
            Shortcut        =   ^I
         End
         Begin VB.Menu MnuSelectScanner 
            Caption         =   "Select Scanner/Camera Source..."
         End
         Begin VB.Menu MnuImportSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFromInternet 
            Caption         =   "From Internet..."
         End
         Begin VB.Menu MnuImportSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScreenCapture 
            Caption         =   "Capture the Screen..."
         End
         Begin VB.Menu MnuImportSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFrx 
            Caption         =   "From Visual Basic Binary File..."
         End
      End
      Begin VB.Menu MnuFileSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu MnuFileSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MnuFileSepBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBatchConvert 
         Caption         =   "&Batch Convert..."
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuFileSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileSepBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu MnuRepeatLast 
         Caption         =   "Repeat &Last Action"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuEditSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "&Copy Image to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste as New Image"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuEmptyClipboard 
         Caption         =   "&Empty Clipboard"
      End
      Begin VB.Menu MnuEditSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPreferences 
         Caption         =   "Program Preferences..."
      End
   End
   Begin VB.Menu MnuImage 
      Caption         =   "&Image"
      Begin VB.Menu MnuDuplicate 
         Caption         =   "&Duplicate"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuImageSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuResample 
         Caption         =   "Resize..."
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuImageSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFlip 
         Caption         =   "Flip (Vertical)"
      End
      Begin VB.Menu MnuMirror 
         Caption         =   "Mirror (Horizontal)"
      End
      Begin VB.Menu MnuRotate 
         Caption         =   "Rotate"
         Begin VB.Menu MnuRotateClockwise 
            Caption         =   "90° Clockwise"
         End
         Begin VB.Menu MnuRotate270Clockwise 
            Caption         =   "90° Counter-clockwise"
         End
         Begin VB.Menu MnuRotate180 
            Caption         =   "180°"
         End
         Begin VB.Menu MnuRotateArbitrary 
            Caption         =   "Arbitrary..."
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnuImageSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIsometric 
         Caption         =   "Convert to Isometric View"
      End
   End
   Begin VB.Menu MnuColor 
      Caption         =   "&Color"
      Begin VB.Menu MnuBrightness 
         Caption         =   "Brightness/Contrast..."
      End
      Begin VB.Menu MnuGamma 
         Caption         =   "Gamma Correction..."
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuImageLevels 
         Caption         =   "Image Levels..."
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuWhiteBalance 
         Caption         =   "White Balance..."
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuSepBarColor2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHistogramTop 
         Caption         =   "Histogram"
         Begin VB.Menu MnuHistogram 
            Caption         =   "Display Histogram"
            Shortcut        =   ^H
         End
         Begin VB.Menu MnuHistogramEqualize 
            Caption         =   "Equalize"
            Begin VB.Menu MnuEqualizeLuminance 
               Caption         =   "Luminance"
            End
            Begin VB.Menu MnuHistogramSepBar1 
               Caption         =   "-"
            End
            Begin VB.Menu MnuEqualizeRed 
               Caption         =   "Red"
            End
            Begin VB.Menu MnuEqualizeGreen 
               Caption         =   "Green"
            End
            Begin VB.Menu MnuEqualizeBlue 
               Caption         =   "Blue"
            End
            Begin VB.Menu MnuEqualizeAll 
               Caption         =   "All Channels"
            End
         End
         Begin VB.Menu MnuHistogramStretch 
            Caption         =   "Stretch"
         End
      End
      Begin VB.Menu MnuColorSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuColorShift 
         Caption         =   "Color Shift"
         Begin VB.Menu MnuCShiftR 
            Caption         =   "Shift Right (r -> g -> b -> r)"
         End
         Begin VB.Menu MnuCShiftL 
            Caption         =   "Shift Left (r -> b -> g -> r)"
         End
      End
      Begin VB.Menu MnuRechannel 
         Caption         =   "Rechannel"
         Begin VB.Menu MnuRR 
            Caption         =   "Red"
         End
         Begin VB.Menu MnuRG 
            Caption         =   "Green"
         End
         Begin VB.Menu MnuRB 
            Caption         =   "Blue"
         End
      End
      Begin VB.Menu MnuSepBarColor1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBlackAndWhite 
         Caption         =   "Black and White..."
      End
      Begin VB.Menu MnuColorize 
         Caption         =   "Colorize..."
      End
      Begin VB.Menu MnuAutoEnhanceTop 
         Caption         =   "Enhance"
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Contrast"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu MnuAutoEnhanceHighlights 
            Caption         =   "Highlights"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MnuAutoEnhanceMidtones 
            Caption         =   "Midtones"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu MnuAutoEnhanceShadows 
            Caption         =   "Shadows"
            Shortcut        =   +{F4}
         End
      End
      Begin VB.Menu MnuFade 
         Caption         =   "Fade"
         Begin VB.Menu MnuFadeLow 
            Caption         =   "Low Fade"
         End
         Begin VB.Menu MnuFadeMedium 
            Caption         =   "Medium Fade"
         End
         Begin VB.Menu MnuFadeHigh 
            Caption         =   "High Fade"
         End
         Begin VB.Menu MnuCustomFade 
            Caption         =   "Custom Fade..."
         End
         Begin VB.Menu MnuFadeSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuUnfade 
            Caption         =   "Unfade"
         End
      End
      Begin VB.Menu MnuGrayscale 
         Caption         =   "Grayscale..."
      End
      Begin VB.Menu MnuInvertTop 
         Caption         =   "Invert"
         Begin VB.Menu MnuCompoundInvert 
            Caption         =   "Compound Invert"
         End
         Begin VB.Menu MnuNegative 
            Caption         =   "Invert CMYK (Film Negative)"
         End
         Begin VB.Menu MnuInvertHue 
            Caption         =   "Invert Hue"
         End
         Begin VB.Menu MnuInvert 
            Caption         =   "Invert RGB"
         End
      End
      Begin VB.Menu MnuPosterize 
         Caption         =   "Posterize..."
      End
      Begin VB.Menu MnuR255 
         Caption         =   "Reduce Image &Colors..."
      End
      Begin VB.Menu MnuColorSepBarPreCountColors 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCountColors 
         Caption         =   "Count image colors"
      End
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "Filte&rs"
      Begin VB.Menu MnuFadeLastEffect 
         Caption         =   "Fade last effect"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuFilterSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArtisticUpper 
         Caption         =   "Artistic"
         Begin VB.Menu MnuAntique 
            Caption         =   "Antique (Sepia)"
         End
         Begin VB.Menu MnuComicBook 
            Caption         =   "Comic Book"
         End
         Begin VB.Menu MnuMosaic 
            Caption         =   "Mosaic..."
         End
         Begin VB.Menu MnuPencil 
            Caption         =   "Pencil Drawing"
         End
         Begin VB.Menu MnuRelief 
            Caption         =   "Relief"
         End
      End
      Begin VB.Menu MnuBlurUpper 
         Caption         =   "Blur"
         Begin VB.Menu MnuAntialias 
            Caption         =   "Antialias"
         End
         Begin VB.Menu BlurSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSoften 
            Caption         =   "Soften"
         End
         Begin VB.Menu MnuSoftenMore 
            Caption         =   "Soften More"
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu MnuBlurMore 
            Caption         =   "Blur More"
         End
         Begin VB.Menu MnuGaussianBlur 
            Caption         =   "Gaussian Blur"
         End
         Begin VB.Menu MnuGaussianBlurMore 
            Caption         =   "Gaussian Blur More"
         End
         Begin VB.Menu BlurSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuGridBlur 
            Caption         =   "Grid Blur"
         End
      End
      Begin VB.Menu MnuDiffuseUpper 
         Caption         =   "Diffuse"
         Begin VB.Menu MnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu MnuDiffuseMore 
            Caption         =   "Diffuse More"
         End
         Begin VB.Menu MnuDiffuseSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomDiffuse 
            Caption         =   "Custom Diffuse..."
         End
      End
      Begin VB.Menu MnuEdge 
         Caption         =   "Edge"
         Begin VB.Menu MnuEmbossEngrave 
            Caption         =   "Emboss/Engrave..."
         End
         Begin VB.Menu MnuEdgeEnhance 
            Caption         =   "Enhance Edges"
         End
         Begin VB.Menu MnuFindEdges 
            Caption         =   "Find Edges..."
         End
      End
      Begin VB.Menu MnuNaturalFilters 
         Caption         =   "Natural"
         Begin VB.Menu MnuAtmosperic 
            Caption         =   "Atmosphere"
         End
         Begin VB.Menu MnuBurn 
            Caption         =   "Burn"
         End
         Begin VB.Menu MnuFogEffect 
            Caption         =   "Fog"
         End
         Begin VB.Menu MnuFrozen 
            Caption         =   "Freeze"
         End
         Begin VB.Menu MnuLava 
            Caption         =   "Lava"
         End
         Begin VB.Menu MnuOcean 
            Caption         =   "Ocean"
         End
         Begin VB.Menu MnuRainbow 
            Caption         =   "Rainbow"
         End
         Begin VB.Menu MnuSteel 
            Caption         =   "Steel"
         End
         Begin VB.Menu MnuWater 
            Caption         =   "Water"
         End
      End
      Begin VB.Menu MnuNoiseFilters 
         Caption         =   "Noise"
         Begin VB.Menu MnuNoise 
            Caption         =   "Add Noise..."
         End
         Begin VB.Menu MnuNoiseSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomDespeckle 
            Caption         =   "Despeckle..."
         End
         Begin VB.Menu MnuDespeckle 
            Caption         =   "Remove Orphan Pixels"
         End
      End
      Begin VB.Menu MnuOtherFilters 
         Caption         =   "Other filters"
         Begin VB.Menu MnuAlien 
            Caption         =   "Alien"
         End
         Begin VB.Menu MnuBlackLight 
            Caption         =   "Black Light..."
         End
         Begin VB.Menu MnuDream 
            Caption         =   "Dream"
         End
         Begin VB.Menu MnuRadioactive 
            Caption         =   "Radioactive"
         End
         Begin VB.Menu MnuSolarize 
            Caption         =   "Solarize..."
         End
         Begin VB.Menu MnuSynthesize 
            Caption         =   "Synthesize"
         End
         Begin VB.Menu MnuTile 
            Caption         =   "Tile..."
         End
         Begin VB.Menu MnuTwins 
            Caption         =   "Twins..."
         End
         Begin VB.Menu MnuVibrate 
            Caption         =   "Vibrate"
         End
      End
      Begin VB.Menu MnuRank 
         Caption         =   "Rank"
         Begin VB.Menu MnuMaximum 
            Caption         =   "Dilate (Maximum)"
         End
         Begin VB.Menu MnuMinimum 
            Caption         =   "Erode (Minimum)"
         End
         Begin VB.Menu MnuExtreme 
            Caption         =   "Extreme (Furthest value)"
         End
         Begin VB.Menu MnuRankSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomRank 
            Caption         =   "Custom Rank..."
         End
      End
      Begin VB.Menu MnuSharpenUpper 
         Caption         =   "Sharpen"
         Begin VB.Menu MnuUnsharp 
            Caption         =   "Remove Blur (Unsharp)"
         End
         Begin VB.Menu MnuSharpenSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu MnuSharpenMore 
            Caption         =   "Sharpen More"
         End
      End
      Begin VB.Menu MnuFilterSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCustomFilter 
         Caption         =   "Custom Filter..."
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuMacro 
      Caption         =   "&Macro"
      Begin VB.Menu MnuPlayMacroRecording 
         Caption         =   "Play Saved Macro..."
      End
      Begin VB.Menu MnuMacroSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStartMacroRecording 
         Caption         =   "&Start Recording"
      End
      Begin VB.Menu MnuStopMacroRecording 
         Caption         =   "Sto&p Recording..."
      End
   End
   Begin VB.Menu MnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuFitOnScreen 
         Caption         =   "&Fit Image On Screen"
      End
      Begin VB.Menu MnuFitWindowToImage 
         Caption         =   "Fit Window to &Image"
      End
      Begin VB.Menu MnuWindowSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu MnuCascadeWindows 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu MnuTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu MnuTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu MnuWindowSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMinimizeAllWindows 
         Caption         =   "&Minimize All Windows"
      End
      Begin VB.Menu MnuRestoreAllWindows 
         Caption         =   "&Restore All Windows"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuDonate 
         Caption         =   "Support PhotoDemon with a small donation (thank you!)"
      End
      Begin VB.Menu MnuHelpSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheckUpdates 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu MnuVisitWebsite 
         Caption         =   "&Visit the PhotoDemon Website"
      End
      Begin VB.Menu MnuEmailAuthor 
         Caption         =   "Submit Feedback..."
      End
      Begin VB.Menu MnuBugReport 
         Caption         =   "Submit Bug Report..."
      End
      Begin VB.Menu MnuHelpSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About PhotoDemon"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please see the included README.txt file for additional information regarding licensing and redistribution.

'PhotoDemon is Copyright ©1999-2012 by Tanner Helland, www.tannerhelland.com

'***************************************************************************
'Main Program MDI Form
'Copyright ©2000-2012 by Tanner Helland
'Created: 9/15/02
'Last updated: 30/July/12
'Last update: new accelerators added (page up, page down for previous, next image respectively)
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its
' primary purpose is sending parameters to other, more interesting sections of the program.
'
'***************************************************************************

Option Explicit

'These functions are used to scroll through consecutive MDI windows without flickering
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, ByVal wCmd As Long) As Long

'When the zoom combo box is changed, redraw the image using the new zoom value
Private Sub CmbZoom_Click()
    
    'Track the current zoom value
    If NumOfWindows > 0 Then pdImages(FormMain.ActiveForm.Tag).CurrentZoomValue = FormMain.CmbZoom.ListIndex
    
    PrepareViewport FormMain.ActiveForm, "Zoom changed"
    
End Sub

Private Sub cmdOpen_Click()
    Process FileOpen
End Sub

Private Sub cmdRedo_Click()
    Process Redo
End Sub

Private Sub cmdSave_Click()
    Process FileSave
End Sub

Private Sub cmdUndo_Click()
    Process Undo
End Sub

'THE BEGINNING OF EVERYTHING
'This form is actually loaded first!  Everything starts here!!
Private Sub MDIForm_Load()

    'Use a global variable to store the command-line parameters
    CommandLine = Command$
    
    'Temporarily exit from here and run the load program subroutine (Loading module)
    LoadTheProgram
    
    'After the program has been successfully loaded, change the focus to the Open Image button
    Me.Visible = True
    cmdOpen.SetFocus
    
    'If the user wants us to check for updates, now's the time to do it
    Dim tmpString As String
    
    'Start by seeing if we're allowed to check for software updates
    tmpString = GetFromIni("General Preferences", "CheckForUpdates")
    Dim allowedToUpdate As Boolean
    If val(tmpString) = 0 Then allowedToUpdate = False Else allowedToUpdate = True
    
    'If updates are allowed, now is the time to check
    If allowedToUpdate = True Then
    
        Message "Checking for software updates (this feature can be disabled from the Edit -> Preferences menu)..."
    
        Dim updateNeeded As Boolean
        updateNeeded = CheckForSoftwareUpdate
        
        If updateNeeded = True Then
            Message "Update found!  Launching update notifier..."
            FormSoftwareUpdate.Show 1, Me
        Else
            Message "Software is up-to-date."
        End If
    
    End If
    
    'Last but not least, if any core plugin files were marked as "missing," offer to download them
    ' (NOTE: this check is superceded by the update check - since a full program update will include the missing plugins -
    '        so ignore this request if the user was already notified of an update.)
    If (updateNeeded = False) And ((zLibEnabled = False) Or (ScanEnabled = False) Or (FreeImageEnabled = False)) Then
    
        Message "Some core plugins could not be found. Preparing updater..."
        
        'As a courtesy, if the user has asked us to stop bugging them about downloading plugins, obey their request
        tmpString = GetFromIni("General Preferences", "PromptForPluginDownload")
        Dim promptToDownload As Boolean
        If val(tmpString) = 0 Then promptToDownload = False Else promptToDownload = True
        
        'Finally, if allowed, we can prompt the user to download the recommended plugin set
        If promptToDownload = True Then
            FormPluginDownloader.Show 1, FormMain
        Else
            Message "Ignoring plugin update request per user's INI settings"
        End If
    
    End If
    
    Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Allow the user to drag-and-drop files from Windows Explorer onto the main MDI form
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        Dim sFile() As String
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        Dim tmpString As String
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Because the OLE drop may include blank strings, verify the size of the array against countFiles
        ReDim Preserve sFile(0 To countFiles - 1) As String
        
        'Pass the list of filenames to PreLoadImage, which will load the images one-at-a-time
        PreLoadImage sFile
        
    End If
    
End Sub

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'If the user is attempting to close the program, run some checks
' Note: in VB6, the order of events for program closing is MDI Parent QueryUnload, MDI children QueryUnload, MDI children Unload, MDI Unload
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'If the histogram form is open, close it
    Unload FormHistogram
    
End Sub

'UNLOAD EVERYTHING
Private Sub MDIForm_Unload(Cancel As Integer)
        
    'By this point, all the child forms should have taken care of their Undo clearing-out.
    ' Just in case, however, prompt a final cleaning.
    ClearALLUndo
    
    'Release GDIPlus (if applicable)
    If GDIPlusEnabled Then releaseGDIPlus
    
    'Release the scanner (if applicable)
    If ScanEnabled Then UnloadScanner
    
    'Destroy all custom-created form icons
    destroyAllIcons
    
    'Release the hand cursor we use for all clickable objects
    destroyHandCursor

    'Save the MRU list to the INI file.  (I've considered doing this as files are loaded, but the
    ' only time that would be an improvement is if the program crashes, and if it does crash, the user
    ' wouldn't want to re-load the problematic image anyway.)
    MRU_SaveToINI
    
End Sub

Private Sub MnuAbout_Click()
    'Show the "about" form
    FormAbout.Show 1, FormMain
End Sub

'Private Sub MnuAnimate_Click()
'    Process Animate
'End Sub

Private Sub MnuAntialias_Click()
    Process Antialias
End Sub

Private Sub MnuAntique_Click()
    Process Antique
End Sub

Private Sub MnuArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub MnuAtmosperic_Click()
    Process Atmospheric
End Sub

Private Sub MnuAutoEnhanceHighlights_Click()
    Process AutoHighlights
End Sub

Private Sub MnuAutoEnhanceMidtones_Click()
    Process AutoMidtones
End Sub

Private Sub MnuAutoEnhanceShadows_Click()
    Process AutoShadows
End Sub

Private Sub MnuBatchConvert_Click()
    FormBatchConvert.Show 1, FormMain
End Sub

Private Sub MnuBlackAndWhite_Click()
    Process BWImpressionist, , , , , , , , , , True
End Sub

Private Sub MnuBlackLight_Click()
    Process BlackLight, , , , , , , , , , True
End Sub

Private Sub MnuBlur_Click()
    Process Blur
End Sub

Private Sub MnuBlurMore_Click()
    Process BlurMore
End Sub

Public Sub MnuBrightness_Click()
    Process BrightnessAndContrast, , , , , , , , , , True
End Sub

Private Sub MnuBugReport_Click()
    
    'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & PROGRAMNAME & " better")
    
    'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
    If msgReturn = vbYes Then
        'Shell a browser window with the GitHub issue report form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/tannerhelland/PhotoDemon/issues/new", "", 0, SW_SHOWNORMAL
    Else
        'Shell a browser window with the tannerhelland.com PhotoDemon contact form
        ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon-contact/", "", 0, SW_SHOWNORMAL
    End If

End Sub

Private Sub MnuBurn_Click()
    Process Burn
End Sub

Private Sub MnuCascadeWindows_Click()
    Me.Arrange vbCascade
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To CurrentImage
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Cascade"
    Next i
    
End Sub

'This allows the user to manually check for updates.
Private Sub MnuCheckUpdates_Click()
        
    Message "Checking for software updates..."
    
    Dim updateNeeded As Boolean
    updateNeeded = CheckForSoftwareUpdate
    
    If updateNeeded = True Then
        Message "Update found!  Launching update notifier..."
        FormSoftwareUpdate.Show 1, Me
    Else
        Message "This copy of PhotoDemon is up-to-date.  (Version " & App.Major & "." & App.Minor & ")"
    End If

End Sub

Private Sub MnuClearMRU_Click()
    MRU_ClearList
End Sub

Private Sub MnuClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub MnuColorize_Click()
    Process Colorize, , , , , , , , , , True
End Sub

Private Sub MnuComicBook_Click()
    Process ComicBook
End Sub

Private Sub MnuCompoundInvert_Click()
    Process CompoundInvert, 128
End Sub

Private Sub MnuCountColors_Click()
    Process CountColors
End Sub

Private Sub MnuCShiftL_Click()
    Process ColorShiftLeft, 1
End Sub

Private Sub MnuCShiftR_Click()
    Process ColorShiftRight, 0
End Sub

Private Sub MnuCustomDespeckle_Click()
    Process CustomDespeckle, , , , , , , , , , True
End Sub

Private Sub MnuCustomDiffuse_Click()
    Process CustomDiffuse, , , , , , , , , , True
End Sub

Private Sub MnuCustomFade_Click()
    Process Fade, , , , , , , , , , True
End Sub

Private Sub MnuCustomFilter_Click()
    Process CustomFilter, , , , , , , , , , True
End Sub

Private Sub MnuCustomRank_Click()
    Process CustomRank, , , , , , , , , , True
End Sub

Private Sub MnuDespeckle_Click()
    Process Despeckle
End Sub

Private Sub MnuDiffuse_Click()
    Process Diffuse
End Sub

Private Sub MnuDiffuseMore_Click()
    Process DiffuseMore
End Sub

Private Sub MnuDonate_Click()
    'Launch the default web browser with the tannerhelland.com donation page
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/donate", "", 0, SW_SHOWNORMAL
End Sub

Private Sub MnuDream_Click()
    Process Dream
End Sub

'Duplicate the current image
Private Sub MnuDuplicate_Click()
    
    'This sub can be found in the "Loading" module
    DuplicateCurrentImage
    
End Sub

Private Sub MnuEdgeEnhance_Click()
    Process EdgeEnhance
End Sub

Private Sub MnuEmailAuthor_Click()
    
    'Shell a browser window with the tannerhelland.com contact form
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon-contact/", "", 0, SW_SHOWNORMAL

End Sub

Private Sub MnuEmbossEngrave_Click()
    Process EmbossToColor, , , , , , , , , , True
End Sub

Private Sub MnuEqualizeAll_Click()
    Process Equalize, True, True, True
End Sub

Private Sub MnuEqualizeBlue_Click()
    Process Equalize, 0, 0, True
End Sub

Private Sub MnuEqualizeGreen_Click()
    Process Equalize, 0, True, 0
End Sub

Private Sub MnuEqualizeLuminance_Click()
    Process EqualizeLuminance
End Sub

Private Sub MnuEqualizeRed_Click()
    Process Equalize, True, 0, 0
End Sub

Private Sub MnuExtreme_Click()
    Process RankExtreme
End Sub

Private Sub MnuFadeHigh_Click()
    Process Fade, 75
End Sub

Private Sub MnuFadeLastEffect_Click()
    Process FadeLastEffect
End Sub

Private Sub MnuFadeLow_Click()
    Process Fade, 25
End Sub

Private Sub MnuFadeMedium_Click()
    Process Fade, 50
End Sub

Private Sub MnuFindEdges_Click()
    Process Laplacian, , , , , , , , , , True
End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuFitWindowToImage_Click()
    FitWindowToImage
End Sub

Private Sub MnuFogEffect_Click()
    Process FogEffect
End Sub

Private Sub MnuFrozen_Click()
    Process Frozen
End Sub

Private Sub MnuGamma_Click()
    Process GammaCorrection, , , , , , , , , , True
End Sub

Private Sub MnuGaussianBlur_Click()
    Process GaussianBlur
End Sub

Private Sub MnuGaussianBlurMore_Click()
    Process GaussianBlurMore
End Sub

Private Sub MnuGrayscale_Click()
    Process GrayScale, , , , , , , , , , True
End Sub

Private Sub MnuGridBlur_Click()
    Process GridBlur
End Sub

Private Sub MnuHistogram_Click()
    Process ViewHistogram, , , , , , , , , , True
End Sub

Private Sub MnuHistogramStretch_Click()
    Process StretchHistogram
End Sub

Private Sub MnuImageLevels_Click()
    Process ImageLevels, , , , , , , , , , True
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    If FormInternetImport.Visible = False Then FormInternetImport.Show 1, FormMain
End Sub

Private Sub MnuImportFrx_Click()
    On Error Resume Next
    If FormImportFrx.Visible = False Then FormImportFrx.Show 1, FormMain
End Sub

Private Sub MnuAlien_Click()
    Process Alien
End Sub

Private Sub MnuInvertHue_Click()
    Process InvertHue
End Sub

Private Sub MnuIsometric_Click()
    Process Isometric
End Sub

Private Sub MnuLava_Click()
    Process Lava
End Sub

Private Sub MnuMaximum_Click()
    Process RankMaximum
End Sub

Private Sub MnuMinimizeAllWindows_Click()
    'Run a loop through every child form and minimize it
    Dim tForm As Form
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then tForm.WindowState = vbMinimized
    Next
End Sub

Private Sub MnuMinimum_Click()
    Process RankMinimum
End Sub

Private Sub MnuMosaic_Click()
    Process Mosaic, , , , , , , , , , True
End Sub

Private Sub MnuNegative_Click()
    Process Negative
End Sub

Private Sub MnuNoise_Click()
    Process Noise, , , , , , , , , , True
End Sub

Private Sub MnuOcean_Click()
    Process Ocean
End Sub

Private Sub MnuPencil_Click()
    Process Pencil
End Sub

Private Sub MnuCopy_Click()
    Process cCopy
End Sub

Private Sub MnuEmptyClipboard_Click()
    Process cEmpty
End Sub

Private Sub MnuExit_Click()
    Unload FormMain
End Sub

Private Sub MnuFlip_Click()
    Process Flip
End Sub

'Once I get this working (working WELL :), it will be re-enabled
Private Sub MnuFreeRotate_Click()
    Process FreeRotate, , , , , , , , , , True
End Sub

Private Sub MnuInvert_Click()
    Process Invert
End Sub

Private Sub MnuMirror_Click()
    Process Mirror
End Sub

Private Sub MnuOpen_Click()
    Process FileOpen
End Sub

Private Sub MnuPaste_Click()
    Process cPaste
End Sub

Private Sub MnuPlayMacroRecording_Click()
    Process MacroPlayRecording
End Sub

Private Sub MnuPosterize_Click()
    Process Posterize, , , , , , , , , , True
End Sub

Private Sub MnuPreferences_Click()
    If FormPreferences.Visible = False Then FormPreferences.Show 1, FormMain
End Sub

Private Sub MnuPrint_Click()
    If FormPrint.Visible = False Then FormPrint.Show 1, FormMain
End Sub

Private Sub MnuR255_Click()
    Process ReduceColors, , , , , , , , , , True
End Sub

Private Sub MnuRadioactive_Click()
    Process Radioactive
End Sub

Private Sub MnuRainbow_Click()
    Process Rainbow
End Sub

Private Sub MnuRB_Click()
    Process RechannelBlue, 2
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Check - just in case - to make sure the path isn't empty
    If mnuRecDocs(Index).Caption <> "" Then
        Message "Preparing to load MRU entry..."
        
        'Strip the accelerator caption from this recent menu entry
        Dim tmpString As String
        tmpString = mnuRecDocs(Index).Caption
        StripAcceleratorFromCaption tmpString
        
        'Because PreLoadImage requires a string array, create an array to pass it
        Dim sFile(0) As String
        sFile(0) = tmpString
        
        PreLoadImage sFile
    End If
    
End Sub

Private Sub MnuRedo_Click()
    Process Redo
End Sub

Private Sub MnuRepeatLast_Click()
    Process LastCommand
End Sub

Private Sub MnuRelief_Click()
    Process Relief
End Sub

Private Sub MnuResample_Click()
    Process ImageSize, , , , , , , , , , True
End Sub

Private Sub MnuAutoEnhance_Click()
    Process AutoEnhance
End Sub

Private Sub MnuRestoreAllWindows_Click()
    'Run a loop through every child form and un-minimize it
    Dim tForm As Form
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then
            tForm.WindowState = vbNormal
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - VB is quirky about triggering that event reliably)
            If pdImages(tForm.Tag).IsActive = True Then PrepareViewport tForm, "Restore all windows"
        End If
    Next
End Sub

Private Sub MnuRG_Click()
    Process RechannelGreen, 1
End Sub

Private Sub MnuRotate180_Click()
    Process Rotate180
End Sub

Private Sub MnuRotate270Clockwise_Click()
    Process Rotate270Clockwise
End Sub

Private Sub MnuRotateArbitrary_Click()
    Process FreeRotate, , , , , , , , , , True
End Sub

Private Sub MnuRotateClockwise_Click()
    Process Rotate90Clockwise
End Sub

Private Sub MnuRR_Click()
    Process RechannelRed, 0
End Sub

Private Sub MnuSave_Click()
    Process FileSave
End Sub

Private Sub MnuSaveAs_Click()
    Process FileSaveAs
End Sub

Private Sub MnuScanImage_Click()
    Process ScanImage
End Sub

Private Sub MnuScreenCapture_Click()
    Process capScreen
End Sub

Private Sub MnuSelectScanner_Click()
    Process SelectScanner
End Sub

Private Sub MnuSharpen_Click()
    Process Sharpen
End Sub

Private Sub MnuSharpenMore_Click()
    Process SharpenMore
End Sub

Private Sub MnuSoften_Click()
    Process Soften
End Sub

Private Sub MnuSoftenMore_Click()
    Process SoftenMore
End Sub

Private Sub MnuSolarize_Click()
    Process Solarize, , , , , , , , , , True
End Sub

Private Sub MnuStartMacroRecording_Click()
    Process MacroStartRecording
End Sub

Private Sub MnuSteel_Click()
    Process Steel
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process MacroStopRecording
End Sub

Private Sub MnuSynthesize_Click()
    Process Synthesize
End Sub

Private Sub MnuTest_Click()
    MenuTest
End Sub

Private Sub MnuTile_Click()
    Process Tile, , , , , , , , , , True
End Sub

Private Sub MnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To CurrentImage
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Tile horizontally"
    Next i
    
End Sub

Private Sub MnuTileVertically_Click()
    Me.Arrange vbTileVertical
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To CurrentImage
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Tile vertically"
    Next i
    
End Sub

Private Sub MnuTwins_Click()
    Process Twins, , , , , , , , , , True
End Sub

Private Sub MnuUndo_Click()
    Process Undo
End Sub

Private Sub MnuUnfade_Click()
    Process Unfade
End Sub

Private Sub MnuUnsharp_Click()
    Process Unsharp
End Sub

Private Sub MnuVibrate_Click()
    Process Vibrate
End Sub

Private Sub MnuVisitWebsite_Click()
    'Nothing special here - just launch the default web browser with PhotoDemon's page on tannerhelland.com
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon", "", 0, SW_SHOWNORMAL
End Sub

Private Sub MnuWater_Click()
    Process Water
End Sub

Private Sub MnuWhiteBalance_Click()
    Process WhiteBalance, , , , , , , , , , True
End Sub

'Because VB doesn't allow key tracking in MDIForms, we have to hook keypresses via this method.
' Many thanks to Steve McMahon for the usercontrol that helps implement this
Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than one quarter-second apart.
    Static lastAccelerator As Single
    
    If Timer - lastAccelerator < 0.25 Then Exit Sub

    'Import from Internet
    If ctlAccelerator.Key(nIndex) = "Internet_Import" Then
        If FormInternetImport.Visible = False Then FormInternetImport.Show 1, FormMain
    End If
    
    'Save As...
    If ctlAccelerator.Key(nIndex) = "Save_As" Then
        If FormMain.MnuSaveAs.Enabled = True Then Process FileSaveAs
    End If
    
    'Capture the screen
    If ctlAccelerator.Key(nIndex) = "Screen_Capture" Then Process capScreen
    
    'Import from FRX
    If ctlAccelerator.Key(nIndex) = "Import_FRX" Then
        On Error Resume Next
        If FormImportFrx.Visible = False Then FormImportFrx.Show 1, FormMain
    End If

    'Open program preferences
    If ctlAccelerator.Key(nIndex) = "Preferences" Then
        If FormPreferences.Visible = False Then FormPreferences.Show 1, FormMain
    End If
    
    'Redo
    If ctlAccelerator.Key(nIndex) = "Redo" Then
        If FormMain.MnuRedo.Enabled = True Then Process Redo
    End If
    
    'Repeat last action
    If ctlAccelerator.Key(nIndex) = "Repeat_Last" Then
        If FormMain.MnuRepeatLast.Enabled = True Then Process LastCommand
    End If
    
    'Empty clipboard
    If ctlAccelerator.Key(nIndex) = "Empty_Clipboard" Then Process cEmpty
    
    'Zoom in
    If ctlAccelerator.Key(nIndex) = "Zoom_In" Then
        If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
    End If
    
    'Zoom out
    If ctlAccelerator.Key(nIndex) = "Zoom_Out" Then
        If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
    End If
    
    'Escape - right now it's only used to cancel batch conversions, but it could be applied elsewhere
    If ctlAccelerator.Key(nIndex) = "Escape" Then
        If MacroStatus = MacroBATCH Then MacroStatus = MacroCANCEL
    End If
    
    'Brightness/Contrast
    If ctlAccelerator.Key(nIndex) = "Bright_Contrast" Then
        Process BrightnessAndContrast, , , , , , , , , , True
    End If
    
    'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
    If ctlAccelerator.Key(nIndex) = "Prev_Image" Or ctlAccelerator.Key(nIndex) = "Next_Image" Then
    
        'If one (or zero) images are loaded, ignore this accelerator
        If NumOfWindows <= 1 Then Exit Sub
    
        'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
        Dim MDIClient As Long
        MDIClient = GetWindow(FormMain.HWnd, 5)
        
        'Use the API to instruct the MDI window to move one window forward or back
        If ctlAccelerator.Key(nIndex) = "Prev_Image" Then
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 0&
        Else
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 1&
        End If
    
    End If
    
    'MRU files
    Dim i As Integer
    For i = 0 To 9
        If ctlAccelerator.Key(nIndex) = ("MRU_" & i) Then
            If FormMain.mnuRecDocs.Count > i Then
                If FormMain.mnuRecDocs(i).Enabled = True Then
                    FormMain.mnuRecDocs_Click i
                End If
            End If
        End If
    Next i
    
    lastAccelerator = Timer
    
End Sub

'When the form is resized, the progress bar at bottom needs to be manually redrawn.  Unfortunately, VB doesn't trigger
' the Resize() event properly for MDI parent forms, so we use the picProgBar resize event instead.
Private Sub picProgBar_Resize()
    picProgBar.Refresh
    cProgBar.Draw
End Sub
