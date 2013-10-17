VERSION 5.00
Begin VB.Form FormMain 
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   10455
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17205
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10455
   ScaleWidth      =   17205
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProgBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1147
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9960
      Width           =   17205
   End
   Begin PhotoDemon.vbalHookControl ctlAccelerator 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin VB.Menu MnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu MnuFile 
         Caption         =   "&Open..."
         Index           =   0
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Open &recent"
         Index           =   1
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
      Begin VB.Menu MnuFile 
         Caption         =   "&Import"
         Index           =   2
         Begin VB.Menu MnuImportClipboard 
            Caption         =   "From clipboard"
         End
         Begin VB.Menu MnuImportSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScanImage 
            Caption         =   "From scanner or camera..."
         End
         Begin VB.Menu MnuSelectScanner 
            Caption         =   "Select which scanner or camera to use"
         End
         Begin VB.Menu MnuImportSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFromInternet 
            Caption         =   "Online image..."
         End
         Begin VB.Menu MnuImportSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScreenCapture 
            Caption         =   "Screen capture..."
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Save"
         Index           =   4
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save &as..."
         Index           =   5
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Close"
         Index           =   7
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Close all"
         Index           =   8
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Batch process..."
         Index           =   10
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Print..."
         Index           =   12
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuFile 
         Caption         =   "E&xit"
         Index           =   14
      End
   End
   Begin VB.Menu MnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu MnuEdit 
         Caption         =   "&Undo"
         Index           =   0
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Redo"
         Index           =   1
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Repeat &last action"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Copy to clipboard"
         Index           =   4
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Paste as new image"
         Index           =   5
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Empty clipboard"
         Index           =   6
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuFitOnScreen 
         Caption         =   "&Fit image on screen"
      End
      Begin VB.Menu MnuFitWindowToImage 
         Caption         =   "Fit viewport around &image"
      End
      Begin VB.Menu MnuViewSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuZoomIn 
         Caption         =   "Zoom &in"
      End
      Begin VB.Menu MnuZoomOut 
         Caption         =   "Zoom &out"
      End
      Begin VB.Menu MnuViewSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "16:1 (1600%)"
         Index           =   0
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "8:1 (800%)"
         Index           =   1
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "4:1 (400%)"
         Index           =   2
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "2:1 (200%)"
         Index           =   3
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:1 (actual size, 100%)"
         Index           =   4
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:2 (50%)"
         Index           =   5
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:4 (25%)"
         Index           =   6
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:8 (12.5%)"
         Index           =   7
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:16 (6.25%)"
         Index           =   8
      End
   End
   Begin VB.Menu MnuImageTop 
      Caption         =   "&Image"
      Begin VB.Menu MnuImage 
         Caption         =   "&Duplicate"
         Index           =   0
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Transparency"
         Index           =   2
         Begin VB.Menu MnuTransparency 
            Caption         =   "Add basic transparency..."
            Index           =   0
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "Make color transparent..."
            Index           =   1
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "Remove transparency..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Resize..."
         Index           =   4
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Canvas size..."
         Index           =   5
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Crop to selection"
         Index           =   7
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Autocrop"
         Index           =   8
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Rotate"
         Index           =   10
         Begin VB.Menu MnuRotate 
            Caption         =   "90° clockwise"
            Index           =   0
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90° counter-clockwise"
            Index           =   1
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "180°"
            Index           =   2
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Arbitrary..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip horizontal"
         Index           =   11
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip vertical"
         Index           =   12
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Convert to isometric view"
         Index           =   13
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Indexed color..."
         Index           =   15
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Tile..."
         Index           =   16
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Metadata"
         Index           =   18
         Begin VB.Menu MnuMetadata 
            Caption         =   "Browse image metadata..."
            Index           =   0
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Count unique colors"
            Index           =   2
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Map photo location..."
            Index           =   3
         End
      End
   End
   Begin VB.Menu MnuSelectTop 
      Caption         =   "&Select"
      Begin VB.Menu MnuSelect 
         Caption         =   "All"
         Index           =   0
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "None"
         Index           =   1
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Invert"
         Index           =   2
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Grow..."
         Index           =   4
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Shrink..."
         Index           =   5
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Border..."
         Index           =   6
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Feather..."
         Index           =   7
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Sharpen..."
         Index           =   8
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Load selection..."
         Index           =   10
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Save current selection..."
         Index           =   11
      End
   End
   Begin VB.Menu MnuAdjustmentsTop 
      Caption         =   "&Adjustments"
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Black and white..."
         Index           =   0
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Brightness and contrast..."
         Index           =   1
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Color balance..."
         Index           =   2
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Curves..."
         Index           =   3
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Levels..."
         Index           =   4
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Vibrance..."
         Index           =   5
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "White balance..."
         Index           =   6
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Channels"
         Index           =   8
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Channel mixer..."
            Index           =   0
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Rechannel..."
            Index           =   1
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Maximum channel"
            Index           =   3
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Minimum channel"
            Index           =   4
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift channels left"
            Index           =   6
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift channels right"
            Index           =   7
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Color"
         Index           =   9
         Begin VB.Menu MnuColor 
            Caption         =   "Color balance..."
            Index           =   0
         End
         Begin VB.Menu MnuColor 
            Caption         =   "White balance..."
            Index           =   1
         End
         Begin VB.Menu MnuColor 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Hue and saturation..."
            Index           =   3
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Photo filters..."
            Index           =   4
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Vibrance..."
            Index           =   5
         End
         Begin VB.Menu MnuColor 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Black and white..."
            Index           =   7
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Colorize..."
            Index           =   8
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Sepia"
            Index           =   9
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Histogram"
         Index           =   10
         Begin VB.Menu MnuHistogram 
            Caption         =   "Display histogram"
         End
         Begin VB.Menu mnuHistogramSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuHistogramEqualize 
            Caption         =   "Equalize..."
         End
         Begin VB.Menu MnuHistogramStretch 
            Caption         =   "Stretch"
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Invert"
         Index           =   11
         Begin VB.Menu MnuNegative 
            Caption         =   "Invert CMYK (film negative)"
         End
         Begin VB.Menu MnuInvertHue 
            Caption         =   "Invert hue"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert RGB"
         End
         Begin VB.Menu mnuInvertSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCompoundInvert 
            Caption         =   "Compound invert"
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Lighting"
         Index           =   12
         Begin VB.Menu MnuLighting 
            Caption         =   "Brightness and contrast..."
            Index           =   0
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Curves..."
            Index           =   1
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Exposure..."
            Index           =   2
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Gamma..."
            Index           =   3
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Levels..."
            Index           =   4
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Shadows and highlights..."
            Index           =   5
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Temperature..."
            Index           =   6
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Monochrome"
         Index           =   13
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Color to monochrome..."
            Index           =   0
         End
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Monochrome to grayscale..."
            Index           =   1
         End
      End
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "Effe&cts"
      Begin VB.Menu MnuFadeLastEffect 
         Caption         =   "Fade last effect"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuFilterSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Artistic"
         Index           =   0
         Begin VB.Menu MnuArtistic 
            Caption         =   "Comic book"
            Index           =   0
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Figured glass (dents)..."
            Index           =   1
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Film noir"
            Index           =   2
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Kaleiodoscope..."
            Index           =   3
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Modern art..."
            Index           =   4
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Oil painting..."
            Index           =   5
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Pencil drawing"
            Index           =   6
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Posterize..."
            Index           =   7
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Relief"
            Index           =   8
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Blur"
         Index           =   1
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Box blur..."
            Index           =   0
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Gaussian blur..."
            Index           =   1
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Grid blur"
            Index           =   2
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Motion blur..."
            Index           =   3
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Pixelate..."
            Index           =   4
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Radial blur..."
            Index           =   5
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Smart blur..."
            Index           =   6
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Zoom blur..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Distort"
         Index           =   2
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Apply lens distortion..."
            Index           =   0
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Correct existing lens distortion..."
            Index           =   1
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Miscellaneous..."
            Index           =   2
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Pan and zoom..."
            Index           =   3
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Perspective..."
            Index           =   4
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Pinch and whirl..."
            Index           =   5
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Poke..."
            Index           =   6
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Polar conversion..."
            Index           =   7
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Ripple..."
            Index           =   8
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Rotate..."
            Index           =   9
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Shear..."
            Index           =   10
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Spherize..."
            Index           =   11
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Squish..."
            Index           =   12
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Swirl..."
            Index           =   13
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Waves..."
            Index           =   14
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Edge"
         Index           =   3
         Begin VB.Menu MnuEdge 
            Caption         =   "Emboss or engrave..."
            Index           =   0
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Enhance edges"
            Index           =   1
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Find edges..."
            Index           =   2
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Trace contour..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Experimental"
         Index           =   4
         Begin VB.Menu MnuAlien 
            Caption         =   "Alien"
         End
         Begin VB.Menu MnuBlackLight 
            Caption         =   "Black light..."
         End
         Begin VB.Menu MnuDream 
            Caption         =   "Dream"
         End
         Begin VB.Menu MnuRadioactive 
            Caption         =   "Radioactive"
         End
         Begin VB.Menu MnuSynthesize 
            Caption         =   "Synthesize"
         End
         Begin VB.Menu MnuHeatmap 
            Caption         =   "Thermograph (heat map)"
         End
         Begin VB.Menu MnuVibrate 
            Caption         =   "Vibrate"
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Natural"
         Index           =   5
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Atmosphere"
            Index           =   0
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Burn"
            Index           =   1
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Fog"
            Index           =   2
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Freeze"
            Index           =   3
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Lava"
            Index           =   4
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Rainbow"
            Index           =   5
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Steel"
            Index           =   6
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Underwater"
            Index           =   7
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Noise"
         Index           =   6
         Begin VB.Menu MnuNoise 
            Caption         =   "Add film grain..."
            Index           =   0
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Add RGB noise..."
            Index           =   1
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Median..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Sharpen"
         Index           =   8
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen..."
            Index           =   0
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Unsharp masking..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Stylize"
         Index           =   9
         Begin VB.Menu MnuStylize 
            Caption         =   "Antique"
            Index           =   0
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Diffuse..."
            Index           =   1
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Dilate..."
            Index           =   2
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Erode..."
            Index           =   3
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Solarize..."
            Index           =   4
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Twins..."
            Index           =   5
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Vignetting..."
            Index           =   6
         End
      End
      Begin VB.Menu MnuFilterSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCustomFilter 
         Caption         =   "Custom filter..."
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTool 
         Caption         =   "Language"
         Index           =   0
         Begin VB.Menu mnuLanguages 
            Caption         =   "English (US)"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Language editor..."
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Macros"
         Index           =   3
         Begin VB.Menu MnuPlayMacroRecording 
            Caption         =   "Play saved macro..."
         End
         Begin VB.Menu MnuMacroSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuStartMacroRecording 
            Caption         =   "&Record new macro"
         End
         Begin VB.Menu MnuStopMacroRecording 
            Caption         =   "Sto&p recording..."
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuTool 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Options"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Plugin manager"
         Index           =   6
      End
   End
   Begin VB.Menu MnuWindowTop 
      Caption         =   "&Window"
      Begin VB.Menu MnuWindow 
         Caption         =   "File toolbar"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Selection toolbar"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Floating toolboxes"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Floating image windows"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Next image"
         Index           =   6
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Previous image"
         Index           =   7
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Cascade"
         Index           =   9
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tile &horizontally"
         Index           =   10
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tile &vertically"
         Index           =   11
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Minimize all windows"
         Index           =   13
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Restore all windows"
         Index           =   14
      End
   End
   Begin VB.Menu MnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu MnuHelp 
         Caption         =   "Support PhotoDemon with a small donation (thank you!)"
         Index           =   0
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Check for &updates"
         Index           =   2
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit feedback..."
         Index           =   3
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit bug report..."
         Index           =   4
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&Visit the PhotoDemon website"
         Index           =   6
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Download PhotoDemon's source code"
         Index           =   7
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Read PhotoDemon's license and terms of use"
         Index           =   8
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&About PhotoDemon"
         Index           =   10
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please see the included README.txt file for additional information regarding licensing and redistribution.

'PhotoDemon is Copyright ©1999-2013 by Tanner Helland, tannerhelland.com

'Please visit photodemon.org for updates and additional downloads

'***************************************************************************
'Main Program Form
'Copyright ©2002-2013 by Tanner Helland
'Created: 15/September/02
'Last updated: 16/October/13
'Last update: use the system setting for keyboard delay when processing back-to-back accelerators
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its
' primary purpose is sending parameters to other, more interesting sections of the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Private m_ToolTip As clsToolTip

'A weird situation arises on the main form if the language is changed at run-time.  All tooltips will have already been moved to
' a custom tooltip class, so the controls themselves will not contain any tooltips.  Thus they will also not be re-translated to
' a new language.  To remedy this, we make a backup of all tooltips when the program is first run.  We then re-apply this backup
' collection whenever we need tooltips replaced.
Private tooltipBackup As Collection


'THE BEGINNING OF EVERYTHING
' Actually, Sub "Main" in the module "modMain" is loaded first, but all it does is set up native theming.  Once it has done that, FormMain is loaded.
Private Sub Form_Load()

    'Use a global variable to store any command line parameters we may have been passed
    g_CommandLine = Command$
    
    'Instantiate the themed tooltip class
    Set m_ToolTip = New clsToolTip
    
    'The bulk of the loading code actually takes place inside the LoadTheprogram subroutine (which can be found in the "Loading" module)
    LoadTheProgram
        
    'Hide the selection tools
    metaToggle tSelection, False
    
    'We can now display the main form and any visible toolbars.  (There is currently a flicker if toolbars are hiden, and I'm working
    ' on a solution to that.)
    Me.Visible = True
    
    g_WindowManager.registerChildForm toolbar_File, TOOLBAR_WINDOW, 1, FILE_TOOLBOX
    g_WindowManager.registerChildForm toolbar_Selections, TOOLBAR_WINDOW, 3, SELECTION_TOOLBOX
    g_WindowManager.registerChildForm toolbar_ImageTabs, IMAGE_TABSTRIP
    
    toolbar_File.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_File.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show File Toolbox", True)
    toolbar_Selections.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_Selections.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True)
    toolbar_ImageTabs.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, False
                
    'Before continuing with the last few steps of interface initialization, we need to make sure the user is being presented
    ' with an interface they can understand - thus we need to evaluate the current language and make changes as necessary.
    
    'Start by asking the translation engine if it thinks we should display a language dialog.  (The conditions that trigger
    ' this are described in great detail in the pdTranslate class.)
    Dim lDialogReason As Long
    If g_Language.isLanguageDialogNeeded(lDialogReason) Then
    
        'If we are inside this block, the translation engine thinks we should ask the user to pick a language.  The reason
        ' for this is stored in the lDialogReason variable, and the values correspond to the following:
        ' 0) User-initiated dialog (irrelevant in this case; the return should never be 0)
        ' 1) First-time user, and an approximate (but not exact) language match was found.  Ask them to clarify.
        ' 2) First-time user, and no language match found.  Give them a language dialog in English.
        ' 3) Not a first-time user, but the preferred language file couldn't be located.  Ask them to pick a new one.
        
    
    End If
    
    'Start by seeing if we're allowed to check for software updates
    Dim allowedToUpdate As Boolean
    allowedToUpdate = g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True)
        
    'If updates ARE allowed, see when we last checked.  To be polite, only check once every 10 days.
    If allowedToUpdate Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPref_String("Updates", "Last Update Check", "")
        
        'If the last update check date was not found, request an update check now
        If lastCheckDate = "" Then
        
            allowedToUpdate = True
        
        'If a last update check date was found, check to see how much time has elapsed since that check
        Else
        
            Dim currentDate As Date
            currentDate = Format$(Now, "Medium Date")
            
            'If 10 days have elapsed, allow an update check
            If CLng(DateDiff("d", CDate(lastCheckDate), currentDate)) >= 10 Then
                allowedToUpdate = True
            
            'If 10 days haven't passed, prevent an update
            Else
                Message "Update check postponed (a check has been performed in the last 10 days)"
                allowedToUpdate = False
            End If
                    
        End If
    
    End If
    
    'If we're STILL allowed to update, do so now (unless this is the first time the user has run the program; in that case, suspend updates)
    If allowedToUpdate And (Not g_IsFirstRun) Then
    
        Message "Checking for software updates (this feature can be disabled from the Tools -> Options menu)..."
    
        Dim updateNeeded As UpdateCheck
        updateNeeded = CheckForSoftwareUpdate
        
        'CheckForSoftwareUpdate can return one of three values:
        ' UPDATE_ERROR - something went wrong (no Internet connection, etc)
        ' UPDATE_NOT_NEEDED - the check was successful, but this version is up-to-date
        ' UPDATE_AVAILABLE - the check was successful, and an update is available
        
        Select Case updateNeeded
        
            Case UPDATE_ERROR
                Message "An error occurred while checking for updates.  Please make sure you have an active Internet connection."
            
            Case UPDATE_NOT_NEEDED
                Message "Software is up-to-date."
                
                'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                
            Case UPDATE_AVAILABLE
                Message "Software update found!  Launching update notifier..."
                FormSoftwareUpdate.Show vbModal, Me
            
        End Select
            
    End If
    
    'Last but not least, if any core plugin files were marked as "missing," offer to download them
    ' (NOTE: this check is superceded by the update check - since a full program update will include the missing plugins -
    '        so ignore this request if the user was already notified of an update.)
    If (updateNeeded <> UPDATE_AVAILABLE) And ((Not isZLibAvailable) Or (Not isEZTwainAvailable) Or (Not isFreeImageAvailable) Or (Not isPngnqAvailable) Or (Not isExifToolAvailable)) Then
    
        Message "Some core plugins could not be found. Preparing updater..."
        
        'As a courtesy, if the user has asked us to stop bugging them about downloading plugins, obey their request
        Dim promptToDownload As Boolean
        promptToDownload = g_UserPreferences.GetPref_Boolean("Updates", "Prompt For Plugin Download", True)
                
        'Finally, if allowed, we can prompt the user to download the recommended plugin set
        If promptToDownload Then
            FormPluginDownloader.Show vbModal, FormMain
            
            'Since plugins may have been downloaded, update the interface to match any new features that may be available.
            LoadPlugins
            applyAllMenuIcons
            resetMenuIcons
            g_ImageFormats.generateInputFormats
            g_ImageFormats.generateOutputFormats
            
        Else
            Message "Ignoring plugin update request per user's saved preference"
        End If
    
    End If
        
    Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    
    'Render the main form with any extra visual styles we've decided to apply
    RedrawMainForm
    
    'TODO: As of 17 Oct '13, I am removing the interface warning.  I think things are now "stable enough" for people to once again
    '       play with nightly builds.
    'MsgBox "WARNING!  PhotoDemon's current interface is undergoing a huge overhaul.  As long as this message remains, the program may not work as expected.  I've suspended nightly builds for now, but if you've downloaded this from GitHub, consider yourself warned." & vbCrLf & vbCrLf & "(Seriously: please do any serious editing with with the 6.0 stable release, available from photodemon.org)", vbExclamation + vbOKOnly + vbApplicationModal, "6.2 Development Warning"
    
    'Because people may be using this code in the IDE, warn them about the consequences of doing so
    If (Not g_IsProgramCompiled) And (g_UserPreferences.GetPref_Boolean("Core", "Display IDE Warning", True)) Then displayIDEWarning
    
    'Finally, return focus to the main form
    FormMain.SetFocus
     
End Sub

'Allow the user to drag-and-drop files from Windows Explorer onto the main form
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

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

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'If the user is attempting to close the program, run some checks.  Specifically, we want to make sure all child forms have been saved.
' Note: in VB6, the order of events for program closing is MDI Parent QueryUnload, MDI children QueryUnload, MDI children Unload, MDI Unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    'If the histogram form is open, close it
    Unload FormHistogram
    
    'If the user wants us to remember the program's last-used location, store those values to file now
    If g_UserPreferences.GetPref_Boolean("Core", "Remember Window Location", True) Then
    
        g_UserPreferences.SetPref_Long "Core", "Last Window State", Me.WindowState
        g_UserPreferences.SetPref_Long "Core", "Last Window Left", Me.Left / Screen.TwipsPerPixelX
        g_UserPreferences.SetPref_Long "Core", "Last Window Top", Me.Top / Screen.TwipsPerPixelY
        g_UserPreferences.SetPref_Long "Core", "Last Window Width", Me.Width / Screen.TwipsPerPixelX
        g_UserPreferences.SetPref_Long "Core", "Last Window Height", Me.Height / Screen.TwipsPerPixelY
    
    End If
    
    'Set a public variable to let other functions know that the user has initiated a program-wide shutdown
    g_ProgramShuttingDown = True
    
    'Before exiting QueryUnload, attempt to unload all children forms.  If any of them cancel shutdown, postpone the program-wide
    ' shutdown as well
    Dim i As Long
    If g_NumOfImagesLoaded > 0 Then
    
        For i = 0 To g_NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive Then
                
                    'This image is active and so is its parent form.  Unload both now.
                    Unload pdImages(i).containingForm
                    
                    'If the child form canceled shut down, it will have reset the g_ProgramShuttingDown variable
                    If Not g_ProgramShuttingDown Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                End If
            End If
        Next i
        
    End If
    
End Sub

'UNLOAD EVERYTHING
Private Sub Form_Unload(Cancel As Integer)
        
    'By this point, all the child forms should have taken care of their Undo clearing-out.
    ' Just in case, however, prompt a final cleaning.
    ClearALLUndo
    
    'Release GDIPlus (if applicable)
    If g_ImageFormats.GDIPlusEnabled Then releaseGDIPlus
    
    'Stop tracking hotkeys
    ctlAccelerator.Enabled = False
    
    'Destroy all custom-created form icons
    destroyAllIcons
    
    'Release the hand cursor we use for all clickable objects
    unloadAllCursors

    'Save the MRU list to the preferences file.  (I've considered doing this as files are loaded, but the only time
    ' that would be an improvement is if the program crashes, and if it does crash, the user wouldn't want to re-load
    ' the problematic image anyway.)
    MRU_SaveToFile
        
    'Restore the user's font smoothing setting as necessary
    handleClearType False
    
    ReleaseFormTheming Me
    
    'Release the window manager
    g_WindowManager.unregisterForm Me
    g_WindowManager.saveAllWindowLocations
    Set g_WindowManager = Nothing
    
    'As a final failsafe, forcibly unload any remaining forms
    Dim tmpForm As Form
    For Each tmpForm In Forms
        
        'Note that there is no need to unload FormMain, as we're about to unload it anyway!
        If tmpForm.Name <> "FormMain" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
        
    Next tmpForm
    
End Sub

'The top-level adjustments menu provides some shortcuts to most-used items.
Private Sub MnuAdjustments_Click(Index As Integer)

    Select Case Index
    
        'Black and white
        Case 0
            Process "Black and white", True
        
        'Brightness and contrast
        Case 1
            Process "Brightness and contrast", True
        
        'Color balance
        Case 2
            Process "Color balance", True
        
        'Curves
        Case 3
            Process "Curves", True
        
        'Levels
        Case 4
            Process "Levels", True
        
        'Vibrance
        Case 5
            Process "Vibrance", True
        
        'White balance
        Case 6
            Process "White balance", True
    
    End Select

End Sub

'All artistic filters are launched here
Private Sub MnuArtistic_Click(Index As Integer)

    Select Case Index
            
        'Comic book
        Case 0
            Process "Comic book"
            
        'Figured glass
        Case 1
            Process "Figured glass", True
            
        'Film noir
        Case 2
            Process "Film noir"
        
        'Kaleidoscope
        Case 3
            Process "Kaleidoscope", True
        
        'Modern art
        Case 4
            Process "Modern art", True
        
        'Oil painting
        Case 5
            Process "Oil painting", True
            
        'Pencil drawing
        Case 6
            Process "Pencil drawing"
                
        'Posterize
        Case 7
            Process "Posterize", True
            
        'Relief
        Case 8
            Process "Relief"
    
    End Select

End Sub


Private Sub MnuBlackLight_Click()
    Process "Black light", True
End Sub

'All blur filters are handled here
Private Sub MnuBlurFilter_Click(Index As Integer)

    Select Case Index
        
        'Box blur
        Case 0
            Process "Box blur", True
            
        'Gaussian blur
        Case 1
            Process "Gaussian blur", True
                
        'Grid blur
        Case 2
            Process "Grid blur"
            
        'Motion blur
        Case 3
            Process "Motion blur", True
            
        'Pixelate (mosaic)
        Case 4
            Process "Pixelate", True
        
        'Radial blur
        Case 5
            Process "Radial blur", True
            
        'Smart Blur
        Case 6
            Process "Smart blur", True
        
        'Zoom Blur
        Case 7
            Process "Zoom blur", True
            
    End Select

End Sub

Private Sub MnuClearMRU_Click()
    MRU_ClearList
End Sub

'All Color sub-menu entries are handled here.
Private Sub MnuColor_Click(Index As Integer)

    Select Case Index
    
        'Color balance
        Case 0
            Process "Color balance", True
        
        'White balance
        Case 1
            Process "White balance", True
        
        '<separator>
        Case 2
        
        'HSL
        Case 3
            Process "Hue and saturation", True
        
        'Photo filters
        Case 4
            Process "Photo filter", True

        'Vibrance
        Case 5
            Process "Vibrance", True
        
        '<separator>
        Case 6
        
        'Grayscale (black and white)
        Case 7
            Process "Black and white", True
        
        'Colorize
        Case 8
            Process "Colorize", True
                
        'Sepia
        Case 9
            Process "Sepia"

    End Select

End Sub

'All entries in the "Color -> Components sub-menu are handled here"
Private Sub MnuColorComponents_Click(Index As Integer)

    Select Case Index
    
        'Channel mixer
        Case 0
            Process "Channel mixer", True
        
        'Rechannel
        Case 1
            Process "Rechannel", True
            
        '<separator>
        Case 2
        
        'Max channel
        Case 3
            Process "Maximum channel"
        
        'Min channel
        Case 4
            Process "Minimum channel"
            
        '<separator>
        Case 5
        
        'Shift colors left
        Case 6
            Process "Shift colors (left)"
            
        'Shift colors right
        Case 7
            Process "Shift colors (right)"
        
    End Select
    
End Sub

Private Sub MnuCompoundInvert_Click()
    Process "Compound invert", False, "128"
End Sub

Private Sub MnuCustomFilter_Click()
    Process "Custom filter", True
End Sub

'All distortion filters happen here
Private Sub MnuDistortFilter_Click(Index As Integer)

    Select Case Index
    
        'Apply lens distort
        Case 0
            Process "Apply lens distortion", True
        
        'Remove lens distort
        Case 1
            Process "Correct lens distortion", True
                
        'Miscellaneous
        Case 2
            Process "Miscellaneous distort", True
            
        'Pan and zoom
        Case 3
            Process "Pan and zoom", True
            
        'Perspective (free)
        Case 4
            Process "Perspective", True
            
        'Pinch and whirl
        Case 5
            Process "Pinch and whirl", True
        
        'Poke
        Case 6
            Process "Poke", True
            
        'Polar conversion
        Case 7
            Process "Polar conversion", True
        
        'Ripple
        Case 8
            Process "Ripple", True
        
        'Rotate
        Case 9
            Process "Rotate", True
        
        'Shear
        Case 10
            Process "Shear", True
            
        'Spherize
        Case 11
            Process "Spherize", True
            
        'Squish (formerly Fixed Perspective)
        Case 12
            Process "Squish", True
        
        'Swirl
        Case 13
            Process "Swirl", True
        
        'Waves
        Case 14
            Process "Waves", True
    
    End Select

End Sub

Private Sub MnuDream_Click()
    Process "Dream"
End Sub

Private Sub MnuEdge_Click(Index As Integer)

    Select Case Index
        
        'Emboss/engrave
        Case 0
            Process "Emboss or engrave", True
         
        'Enhance edges
        Case 1
            Process "Edge enhance"
        
        'Find edges
        Case 2
            Process "Find edges", True
        
        'Trace contour
        Case 3
            Process "Trace contour", True
    
    End Select

End Sub

Private Sub MnuEdit_Click(Index As Integer)

    Select Case Index
    
        'Undo
        Case 0
            Process "Undo", False, , 0
        
        'Redo
        Case 1
            Process "Redo", False, , 0
        
        'Repeat last
        Case 2
            Process "Repeat last action", False, , 1
        
        '<separator>
        Case 3
        
        'Copy to clipboard
        Case 4
            Process "Copy to clipboard", False, , 0, , False
        
        'Paste as new image
        Case 5
            Process "Paste as new image", False, , 0, , False
        
        'Empty clipboard
        Case 6
            Process "Empty clipboard", False, , 0, , False
                
    
    End Select
    
End Sub

Private Sub MnuFadeLastEffect_Click()
    Process "Fade last effect"
End Sub

'All file menu actions are launched from here
Private Sub MnuFile_Click(Index As Integer)

    Select Case Index
    
        'Open
        Case 0
            Process "Open", True
        
        '<Open Recent top-level>
        Case 1
        
        '<Import top-level>
        Case 2
        
        '<separator>
        Case 3
        
        'Save
        Case 4
            Process "Save", True
            
        'Save as
        Case 5
            Process "Save as", True
            
        '<separator>
        Case 6
        
        'Close
        Case 7
            Process "Close", True
        
        'Close all
        Case 8
            Process "Close all", True
        
        '<separator>
        Case 9
        
        'Batch wizard
        Case 10
            Process "Batch wizard", True
        
        '<separator>
        Case 11
        
        'Print
        Case 12
            Process "Print", True
        
        '<separator>
        Case 13
        
        'Exit
        Case 14
            Process "Exit program", True
        
    
    End Select

End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuFitWindowToImage_Click()
    If (pdImages(g_CurrentImage).containingForm.WindowState = vbMaximized) Or (pdImages(g_CurrentImage).containingForm.WindowState = vbMinimized) Then pdImages(g_CurrentImage).containingForm.WindowState = vbNormal
    FitWindowToImage
End Sub

Private Sub MnuHeatmap_Click()
    Process "Thermograph (heat map)"
End Sub

'All help menu entries are launched from here
Private Sub MnuHelp_Click(Index As Integer)

    Select Case Index
        
        'Donations are so very, very welcome!
        Case 0
            OpenURL "http://photodemon.org/donate"
            
        'Check for updates
        Case 2
            Message "Checking for software updates..."
    
            Dim updateNeeded As Long
            updateNeeded = CheckForSoftwareUpdate
    
            'CheckForSoftwareUpdate can return one of three values:
            ' 0 - something went wrong (no Internet connection, etc)
            ' 1 - the check was successful, but this version is up-to-date
            ' 2 - the check was successful, and an update is available
            Select Case updateNeeded
        
                Case 0
                    Message "An error occurred while checking for updates.  Please try again later."
                    
                Case 1
                    Message "This copy of PhotoDemon is the newest available.  (Version %1.%2.%3)", App.Major, App.Minor, App.Revision
                        
                    'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                    g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                        
                Case 2
                    Message "Software update found!  Launching update notifier..."
                    FormSoftwareUpdate.Show vbModal, Me
                
            End Select
        
        'Submit feedback
        Case 3
            OpenURL "http://photodemon.org/about/contact/"
        
        'Submit bug report
        Case 4
            'GitHub requires a login for submitting Issues; check for that first
            Dim msgReturn As VbMsgBoxResult
            
            'If the user has previously been prompted about having a GitHub account, use their previous answer
            If g_UserPreferences.doesValueExist("Core ", "Has GitHub Account") Then
            
                Dim hasGitHub As Boolean
                hasGitHub = g_UserPreferences.GetPref_Boolean("Core", "Has GitHub Account", False)
                
                If hasGitHub Then msgReturn = vbYes Else msgReturn = vbNo
            
            'If this is the first time they are submitting feedback, ask them if they have a GitHub account
            Else
            
                msgReturn = pdMsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNoCancel, "Thanks for fixing PhotoDemon")
                
                'If their answer was anything but "Cancel", store that answer to file
                If msgReturn = vbYes Then g_UserPreferences.SetPref_Boolean "Core", "Has GitHub Account", True
                If msgReturn = vbNo Then g_UserPreferences.SetPref_Boolean "Core", "Has GitHub Account", False
                
            End If
            
            'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the photodemon.org contact form
            If msgReturn = vbYes Then
                'Shell a browser window with the GitHub issue report form
                OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            ElseIf msgReturn = vbNo Then
                'Shell a browser window with the photodemon.org contact form
                OpenURL "http://photodemon.org/about/contact/"
            End If
            
        'PhotoDemon's homepage
        Case 6
            OpenURL "http://www.photodemon.org"
            
        'Download source code
        Case 7
            OpenURL "https://github.com/tannerhelland/PhotoDemon"
        
        'Read terms and license agreement
        Case 8
            OpenURL "http://photodemon.org/about/license/"
            
        'Display About page
        Case 10
            FormAbout.Show vbModal, FormMain
        
    End Select

End Sub

Private Sub MnuHistogram_Click()
    'Process "Display histogram", True
    FormHistogram.Show vbModal, Me
End Sub

Private Sub MnuHistogramEqualize_Click()
    Process "Equalize", True
End Sub

Private Sub MnuHistogramStretch_Click()
    Process "Stretch histogram"
End Sub

'All top-level Image menu actions are handled here
Private Sub MnuImage_Click(Index As Integer)

    Select Case Index
    
        'Duplicate
        Case 0
            Process "Duplicate image", , , False
        
        'Separator
        Case 1
        
        'Transparency top-level
        Case 2
        
        'Separator
        Case 3
        
        'Resize
        Case 4
            Process "Resize", True
        
        'Canvas resize
        Case 5
            Process "Canvas size", True
                
        'Separator
        Case 6
            
        'Crop to selection
        Case 7
            Process "Crop"
        
        'Autocrop
        Case 8
            Process "Autocrop"
        
        'Separator
        Case 9
        
        'Top-level Rotate
        Case 10
        
        'Flip horizontal (mirror)
        Case 11
            Process "Flip horizontal"
        
        'Flip vertical
        Case 12
            Process "Flip vertical"
        
        'Isometric view
        Case 13
            Process "Isometric conversion"
            
        'Separator
        Case 14
        
        'Indexed color
        Case 15
            Process "Reduce colors", True
        
        'Tile
        Case 16
            Process "Tile", True
            
        'Separator
        Case 17
        
        'Metadata top-level
        Case 18
    
    End Select

End Sub

'This is the exact same thing as "Paste as New Image".  It is provided in two locations for convenience.
Private Sub MnuImportClipboard_Click()
    Process "Paste as new image", , , False
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    Process "Internet import", True
End Sub

Private Sub MnuAlien_Click()
    Process "Alien"
End Sub

Private Sub MnuInvertHue_Click()
    Process "Invert hue"
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    
    'Because loading a language can take some time, display a wait screen to discourage attempted interaction
    displayWaitScreen g_Language.TranslateMessage("Please wait while the new language is applied..."), Me
    
    'Remove the existing translation
    Message "Removing existing translation..."
    g_Language.undoTranslations FormMain, True
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.activateNewLanguage Index, True
    
    Message "Language changed successfully."
    
    hideWaitScreen
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub MnuLighting_Click(Index As Integer)

    Select Case Index
            
        'Brightness/Contrast
        Case 0
            Process "Brightness and contrast", True
        
        'Curves
        Case 1
            Process "Curves", True
        
        'Exposure
        Case 2
            Process "Exposure", True
            
        'Gamma correction
        Case 3
            Process "Gamma", True
            
        'Levels
        Case 4
            Process "Levels", True

        'Shadows/Midtones/Highlights
        Case 5
            Process "Shadows and highlights", True
            
        'Temperature
        Case 6
            Process "Temperature", True
    
    End Select

End Sub

'All metadata sub-menu options are handled here
Private Sub MnuMetadata_Click(Index As Integer)

    Select Case Index
    
        'Browse metadata
        Case 0
        
            'Before doing anything else, see if we've already loaded metadata.  If we haven't, do so now.
            If Not pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata Then
                pdImages(g_CurrentImage).imgMetadata.loadAllMetadata pdImages(g_CurrentImage).LocationOnDisk, pdImages(g_CurrentImage).OriginalFileFormat
                
                'If the image contains GPS metadata, enable that option now
                metaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata()
            End If
            
            'If the image STILL doesn't have metadata, warn the user and exit.
            If Not pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata Then
                Message "No metadata available."
                pdMsgBox "This image does not contain any metadata.", vbInformation + vbOKOnly + vbApplicationModal, "No metadata available"
                Exit Sub
            End If
            
            FormMetadata.Show vbModal, Me
        
        'Separator
        Case 1
        
        'Count colors
        Case 2
            Process "Count image colors", True
        
        'Map photo location
        Case 3
            
            'Note that mapping can only be performed if GPS metadata exists for this image.  If the user clicks this option while
            ' using the on-demand model for metadata caching, we must now attempt to load metadata.
            If Not pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata Then
            
                'Attempt to load it now...
                Message "Loading metadata for this image..."
                pdImages(g_CurrentImage).imgMetadata.loadAllMetadata pdImages(g_CurrentImage).LocationOnDisk, pdImages(g_CurrentImage).OriginalFileFormat
                
                'Determine whether metadata is present, and dis/enable metadata menu items accordingly
                metaToggle tMetadata, pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata
                metaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata()
            
            End If
            
            If Not pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata Then
                pdMsgBox "This image does not contain any GPS metadata.", vbOKOnly + vbApplicationModal + vbInformation, "No GPS data found"
                Exit Sub
            End If
            
            Dim gMapsURL As String, latString As String, lonString As String
            If pdImages(g_CurrentImage).imgMetadata.fillLatitudeLongitude(latString, lonString) Then
                
                'Build a valid Google maps URL (you can use Google to see what the various parameters mean)
                
                'Note: I find a zoom of 18 ideal, as that is a common level for switching to an "aerial"
                ' view instead of a satellite view.  Much higher than that and you run the risk of not
                ' having data available at that high of zoom.
                gMapsURL = "https://maps.google.com/maps?f=q&z=18&t=h&q=" & latString & "%2c+" & lonString
                
                'As a convenience, request Google Maps in the current language
                If g_Language.translationActive Then
                    gMapsURL = gMapsURL & "&hl=" & g_Language.getCurrentLanguage()
                Else
                    gMapsURL = gMapsURL & "&hl=en"
                End If
                
                'Launch Google maps in the user's browser
                OpenURL gMapsURL
                
            End If
            
    End Select
    
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    
    Select Case Index
        
        'Convert color to monochrome
        Case 0
            Process "Color to monochrome", True
        
        'Convert monochrome to grayscale
        Case 1
            Process "Monochrome to grayscale", True
        
    End Select
    
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)

    Select Case Index
    
        'Atmosphere
        Case 0
            Process "Atmosphere"
            
        'Burn
        Case 1
            Process "Burn"
        
        'Fog
        Case 2
            Process "Fog"
        
        'Freeze
        Case 3
            Process "Freeze"
        
        'Lava
        Case 4
            Process "Lava"
                
        'Rainbow
        Case 5
            Process "Rainbow"
        
        'Steel
        Case 6
            Process "Steel"
        
        'Water
        Case 7
            Process "Water"
    
    End Select

End Sub

Private Sub MnuNegative_Click()
    Process "Film negative"
End Sub

Private Sub MnuInvert_Click()
    Process "Invert RGB"
End Sub

'All noise filters are handled here
Private Sub MnuNoise_Click(Index As Integer)

    Select Case Index
    
        'Film grain
        Case 0
            Process "Add film grain", True
        
        'RGB Noise
        Case 1
            Process "Add RGB noise", True
            
        'Separator
        Case 2
        
        'Median
        Case 3
            Process "Median", True
            
    End Select
        
End Sub

Private Sub MnuPlayMacroRecording_Click()
    Process "Play macro", True
End Sub

Private Sub MnuRadioactive_Click()
    Process "Radioactive"
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Load the MRU path that correlates to this index.  (If one is not found, a null string is returned)
    Dim tmpString As String
    tmpString = getSpecificMRU(Index)
    
    'Check - just in case - to make sure the path isn't empty
    If tmpString <> "" Then
        
        Message "Preparing to load recent file entry..."
        
        'Because PreLoadImage requires a string array, create an array to pass it
        Dim sFile(0) As String
        sFile(0) = tmpString
        
        PreLoadImage sFile
    End If
    
End Sub

'All rotation actions are initiated here
Private Sub MnuRotate_Click(Index As Integer)

    Select Case Index
    
        'Rotate 90
        Case 0
            Process "Rotate 90° clockwise"
        
        'Rotate 270
        Case 1
            Process "Rotate 90° counter-clockwise"
        
        'Rotate 180
        Case 2
            Process "Rotate 180°"
        
        'Rotate arbitrary
        Case 3
            Process "Arbitrary rotation", True
            
    End Select
            
End Sub

Private Sub MnuScanImage_Click()
    Process "Scan image", True
End Sub

Private Sub MnuScreenCapture_Click()
    Process "Screen capture", True
End Sub

'All select menu items are handled here
Private Sub MnuSelect_Click(Index As Integer)

    Select Case Index
    
        'Select all.  (Note that Square Selection is passed as the relevant tool for this action.)
        Case 0
            Process "Select all", , , 2, 0
        
        'Select none
        Case 1
            Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2
        
        'Invert
        Case 2
            Process "Invert selection", , , 2
        
        '<separator>
        Case 3
        
        'Grow selection
        Case 4
            Process "Grow selection", True, , 0
        
        'Shrink selection
        Case 5
            Process "Shrink selection", True, , 0
        
        'Border selection
        Case 6
            Process "Border selection", True, , 0
        
        'Feather selection
        Case 7
            Process "Feather selection", True, , 0
        
        'Sharpen selection
        Case 8
            Process "Sharpen selection", True, , 0
        
        '<separator>
        Case 9
        
        'Load selection
        Case 10
            Process "Load selection", True, , 0
        
        'Save current selection
        Case 11
            Process "Save selection", True, , 0
        
    End Select

End Sub

Private Sub MnuSelectScanner_Click()
    Process "Select scanner or camera", True
End Sub

'All sharpen filters are handled here
Private Sub MnuSharpen_Click(Index As Integer)

    Select Case Index
            
        'Sharpen
        Case 0
            Process "Sharpen", True
        
        'Unsharp mask
        Case 1
            Process "Unsharp mask", True
            
    End Select

End Sub

'These menu items correspond to specific zoom values
Private Sub MnuSpecificZoom_Click(Index As Integer)

    Select Case Index
    
        Case 0
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 2
        Case 1
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 4
        Case 2
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 8
        Case 3
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 10
        Case 4
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = ZOOM_100_PERCENT
        Case 5
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 14
        Case 6
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 16
        Case 7
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 19
        Case 8
            If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 21
    End Select

End Sub

Private Sub MnuStartMacroRecording_Click()
    Process "Start macro recording", , , False
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process "Stop macro recording", True
End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)

    Select Case Index
    
        'Antique
        Case 0
            Process "Antique"
    
        'Diffuse
        Case 1
            Process "Diffuse", True
        
        'Dilate (maximum rank)
        Case 2
            Process "Dilate (maximum rank)", True
        
        'Erode (minimum rank)
        Case 3
            Process "Erode (minimum rank)", True
        
        'Solarize
        Case 4
            Process "Solarize", True

        'Twins
        Case 5
            Process "Twins", True
            
        'Vignetting
        Case 6
            Process "Vignetting", True
    
    End Select

End Sub

Private Sub MnuSynthesize_Click()
    Process "Synthesize"
End Sub

Private Sub MnuTest_Click()
    MenuTest
End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)

    Select Case Index
    
        'Language editor
        Case 1
            If Not FormLanguageEditor.Visible Then FormLanguageEditor.Show vbModeless, FormMain
    
        'Options
        Case 5
            If Not FormPreferences.Visible Then FormPreferences.Show vbModal, FormMain
            
        'Plugin manager
        Case 6
            If Not FormPluginManager.Visible Then FormPluginManager.Show vbModal, FormMain
            
    End Select

End Sub

'Add / Remove / Modify an image's alpha channel with this menu
Private Sub MnuTransparency_Click(Index As Integer)

    Select Case Index
    
        'Add alpha channel
        Case 0
            
            'Ignore if the current image is already in 32bpp mode
            Process "Add alpha channel", True
            
        'Color to alpha
        Case 1
        
            'Can be used even if the image already has an alpha channel
            Process "Color to alpha", True
        
        '<separator>
        Case 2

        'Remove alpha channel
        Case 3

            'Ignore if the current image is already in 24bpp mode
            If pdImages(g_CurrentImage).mainLayer.getLayerColorDepth = 24 Then Exit Sub
            Process "Remove alpha channel", True
    
    End Select

End Sub

Private Sub MnuVibrate_Click()
    Process "Vibrate"
End Sub

'Because VB doesn't allow key tracking in MDIForms, we have to hook keypresses via this method.
' Many thanks to Steve McMahon for the usercontrol that helps implement this
Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)
    
    'Don't process accelerators when the main form is disabled (e.g. if a modal form is present, or if a previous
    ' action is in the middle of execution)
    If Not FormMain.Enabled Then Exit Sub
    
    'Don't process accelerators if the Language Editor is active
    If Not (FormLanguageEditor Is Nothing) Then
        If FormLanguageEditor.Visible Then Exit Sub
    End If

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than the system keyboard delay (250ms at minimum, 1s at maximum).
    Static lastAccelerator As Double
    If (Timer - lastAccelerator < getKeyboardDelay()) Then Exit Sub

    'Accelerators are divided into three groups, and they are processed in the following order:
    ' 1) Direct processor strings.  These are automatically submitted to the software processor.
    ' 2) Non-processor directives that can be fired if no images are present (e.g. Open, Paste)
    ' 3) Non-processor directives that require an image.

    '***********************************************************
    'Accelerators that are direct processor strings are handled automatically
    
    With ctlAccelerator
    
        If .isProcString(nIndex) Then
            
            'If the action requires an open image, check for that first
            If .imageRequired(nIndex) Then
                If g_OpenImageCount = 0 Then Exit Sub
                If Not (FormLanguageEditor Is Nothing) Then
                    If FormLanguageEditor.Visible Then Exit Sub
                End If
            End If
    
            Process .Key(nIndex), .displayDialog(nIndex), , .shouldCreateUndo(nIndex)
            Exit Sub
            
        End If
    
    End With

    '***********************************************************
    'Accelerators that DO NOT require at least one loaded image, and that require special handling:
    
    'Open program preferences
    If ctlAccelerator.Key(nIndex) = "Preferences" Then
        If Not FormPreferences.Visible Then
            FormPreferences.Show vbModal, FormMain
            Exit Sub
        End If
    End If
    
    If ctlAccelerator.Key(nIndex) = "Plugin manager" Then
        If Not FormPluginManager.Visible Then
            FormPluginManager.Show vbModal, FormMain
            Exit Sub
        End If
    End If
        
    'Escape - a separate function is used to cancel currently running filters.  This accelerator is only used
    ' to cancel batch conversions, but in the future it should be applied elsewhere.
    If ctlAccelerator.Key(nIndex) = "Escape" Then
        If MacroStatus = MacroBATCH Then MacroStatus = MacroCANCEL
    End If
    
    'MRU files
    Dim i As Integer
    For i = 0 To 9
        If ctlAccelerator.Key(nIndex) = ("MRU_" & i) Then
            If FormMain.mnuRecDocs.Count > i Then
                If FormMain.mnuRecDocs(i).Enabled Then
                    FormMain.mnuRecDocs_Click i
                End If
            End If
        End If
    Next i
    
    '***********************************************************
    'Accelerators that DO require at least one loaded image, and that require special handling:
    
    'If no images are loaded, or another form is active, exit.
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Fit on screen
    If ctlAccelerator.Key(nIndex) = "FitOnScreen" Then FitOnScreen
    
    'Zoom in
    If ctlAccelerator.Key(nIndex) = "Zoom_In" Then
        If toolbar_File.CmbZoom.Enabled And toolbar_File.CmbZoom.ListIndex > 0 Then toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex - 1
    End If
    
    'Zoom out
    If ctlAccelerator.Key(nIndex) = "Zoom_Out" Then
        If toolbar_File.CmbZoom.Enabled And toolbar_File.CmbZoom.ListIndex < (toolbar_File.CmbZoom.ListCount - 1) Then toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex + 1
    End If
    
    'Actual size
    If ctlAccelerator.Key(nIndex) = "Actual_Size" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = ZOOM_100_PERCENT
    End If
    
    'Various zoom values
    If ctlAccelerator.Key(nIndex) = "Zoom_161" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 2
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_81" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 4
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_41" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 8
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_21" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 10
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_12" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 14
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_14" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 16
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_18" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 19
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_116" Then
        If toolbar_File.CmbZoom.Enabled Then toolbar_File.CmbZoom.ListIndex = 21
    End If
    
    'Remove selection
    If ctlAccelerator.Key(nIndex) = "Remove selection" Then
        Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2
    End If
    
    'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
    If ctlAccelerator.Key(nIndex) = "Next_Image" Then moveToNextChildWindow True
    If ctlAccelerator.Key(nIndex) = "Prev_Image" Then moveToNextChildWindow False
    
    lastAccelerator = Timer
    
End Sub

'All "Window" menu items are handled here
Private Sub MnuWindow_Click(Index As Integer)

    Dim i As Long
    Dim MDIClient As Long

    Select Case Index
    
        'Show/hide file toolbox
        Case 0
            toggleToolbarVisibility FILE_TOOLBOX
        
        'Show/hide selection toolbox
        Case 1
            toggleToolbarVisibility SELECTION_TOOLBOX
        
        '<separator>
        Case 2
    
        'Floating toolbars
        Case 3
            toggleWindowFloating TOOLBAR_WINDOW, Not FormMain.MnuWindow(3).Checked
            
        'Floating image windows
        Case 4
            toggleWindowFloating IMAGE_WINDOW, Not FormMain.MnuWindow(4).Checked
            
        '<separator>
        Case 5
        
        'Next image
        Case 6
            moveToNextChildWindow True
            
        'Previous image
        Case 7
            moveToNextChildWindow False
    
        '<separator>
        Case 8
        
        'Cascade
        Case 9
            'Me.Arrange vbCascade
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Cascade"
                End If
            Next i
        
        'Tile horizontally
        Case 10
            'Me.Arrange vbTileHorizontal
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Tile vertically"
                End If
            Next i
    
        'Tile vertically
        Case 11
            'Me.Arrange vbTileVertical
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Tile vertically"
                End If
            Next i
    
        '<separator>
        Case 12
        
        'Minimize all windows
        Case 13
        
            'Run a loop through every child form and minimize it
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then
                        pdImages(i).containingForm.WindowState = vbMinimized
                    End If
                End If
            Next i
        
        'Restore all windows
        Case 14
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then
                        pdImages(i).containingForm.WindowState = vbNormal
                        PrepareViewport pdImages(i).containingForm, "Restore all windows"
                    End If
                End If
            Next i
    
    End Select
    

End Sub

'The "Next Image" and "Previous Image" options simply wrap this function.
Private Sub moveToNextChildWindow(ByVal moveForward As Boolean)

    'If one (or zero) images are loaded, ignore this option
    If g_OpenImageCount <= 1 Then Exit Sub
    
    Dim i As Long
    
    'Loop through all available images, and when we find one that is not this image, activate it and exit
    If moveForward Then
        i = g_CurrentImage + 1
    Else
        i = g_CurrentImage - 1
    End If
    
    Do While i <> g_CurrentImage
            
        'Loop back to the start of the window collection
        If moveForward Then
            If i > g_NumOfImagesLoaded Then i = 0
        Else
            If i < 0 Then i = g_NumOfImagesLoaded
        End If
                
        If Not pdImages(i) Is Nothing Then
            If pdImages(i).IsActive Then
                pdImages(i).containingForm.ActivateWorkaround
                Exit Do
            End If
        End If
                
        If moveForward Then
            i = i + 1
        Else
            i = i - 1
        End If
                
    Loop

End Sub

Private Sub MnuZoomIn_Click()
    If toolbar_File.CmbZoom.Enabled And toolbar_File.CmbZoom.ListIndex > 0 Then toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex - 1
End Sub

Private Sub MnuZoomOut_Click()
    If toolbar_File.CmbZoom.Enabled And toolbar_File.CmbZoom.ListIndex < (toolbar_File.CmbZoom.ListCount - 1) Then toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex + 1
End Sub

'When the form is resized, the progress bar at bottom needs to be manually redrawn.  Unfortunately, VB doesn't trigger
' the Resize() event properly for MDI parent forms, so we use the pic_ProgBar resize event instead.
Private Sub picProgBar_Resize()
    
    'When this main form is resized, reapply any custom visual styles
    If FormMain.Visible Then RedrawMainForm
    
End Sub

'Because we want tooltips preserved, outside functions should use THIS sub to request FormMain rethemes
Public Sub requestMakeFormPretty(Optional ByVal useDoEvents As Boolean = False)
    
    Dim eControl As Control
    
    'If we have not made a backup yet, do so now
    If tooltipBackup Is Nothing Then
        Set tooltipBackup = New Collection
        
        'Enumerate through every control on the form.  Store a copy of its tooltip inside our collection.
        For Each eControl In Me.Controls
            
            'If this is a control that has a tooltip, backup the tooltip now
            If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is PictureBox) Or (TypeOf eControl Is TextBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is colorSelector) Or (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then
                
                Dim tmpString As String
                tmpString = eControl.ToolTipText
                
                'If a translation is already active, we want to back up the ENGLISH tooltip - not a translated one.
                If g_Language.translationActive Then tmpString = g_Language.RestoreMessage(tmpString)
                
                If InControlArray(eControl) Then
                    tooltipBackup.Add tmpString, eControl.Name & "_" & eControl.Index
                Else
                    tooltipBackup.Add tmpString, eControl.Name
                End If
            End If
            
        Next
    
    'If we HAVE made a backup and this function is being called, restore all tooltips now.  That way, when they are passed to the
    ' makeFormPretty function, all tooltips will be available for translation.
    Else
    
        'Enumerate through every control on the form.  Restore its tooltip if found.
        For Each eControl In Me.Controls
            
            'If this is a control that has a tooltip, backup the tooltip now
            If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is PictureBox) Or (TypeOf eControl Is TextBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is colorSelector) Or (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then
                
                If InControlArray(eControl) Then
                    eControl.ToolTipText = tooltipBackup.Item(eControl.Name & "_" & eControl.Index)
                Else
                    eControl.ToolTipText = tooltipBackup.Item(eControl.Name)
                End If
                
            End If
            'Or (TypeOf eControl Is textUpDown)
            If (TypeOf eControl Is sliderTextCombo) Then
                eControl.refreshTooltipObject
            End If
            
        Next
        
        Set m_ToolTip = New clsToolTip
    
    End If
    
    makeFormPretty Me, m_ToolTip, , useDoEvents
End Sub
