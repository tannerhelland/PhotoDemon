VERSION 5.00
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   11355
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   18915
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
   ScaleHeight     =   757
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1261
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   1560
   End
   Begin PhotoDemon.pdCanvas mainCanvas 
      Height          =   3735
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
   End
   Begin PhotoDemon.vbalHookControl ctlAccelerator 
      Left            =   120
      Top             =   120
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin PhotoDemon.bluDownload updateChecker 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin PhotoDemon.ShellPipe shellPipeMain 
      Left            =   960
      Top             =   360
      _ExtentX        =   635
      _ExtentY        =   635
      ErrAsOut        =   0   'False
      PollInterval    =   5
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
         Begin VB.Menu MnuLoadAllMRU 
            Caption         =   "Load all recent images"
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
            Caption         =   "Screenshot..."
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Close"
         Index           =   4
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Close all"
         Index           =   5
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Save"
         Index           =   7
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save &as..."
         Index           =   8
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Revert"
         Index           =   9
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Batch process..."
         Index           =   11
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Print..."
         Index           =   13
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuFile 
         Caption         =   "E&xit"
         Index           =   15
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
         Caption         =   "&Copy "
         Index           =   4
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Copy Merged"
         Index           =   5
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Paste as new layer"
         Index           =   6
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Paste as new image"
         Index           =   7
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Empty clipboard"
         Index           =   9
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuFitOnScreen 
         Caption         =   "&Fit image on screen"
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
         Caption         =   "Resize..."
         Index           =   2
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Canvas size..."
         Index           =   4
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Fit canvas to active layer"
         Index           =   5
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Fit canvas around all layers"
         Index           =   6
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Crop to selection"
         Index           =   8
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Trim empty borders"
         Index           =   9
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Rotate"
         Index           =   11
         Begin VB.Menu MnuRotate 
            Caption         =   "Straighten"
            Index           =   0
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90° clockwise"
            Index           =   2
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90° counter-clockwise"
            Index           =   3
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "180°"
            Index           =   4
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Arbitrary..."
            Index           =   5
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip horizontal"
         Index           =   12
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip vertical"
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
   Begin VB.Menu MnuLayerTop 
      Caption         =   "&Layer"
      Begin VB.Menu MnuLayer 
         Caption         =   "Add"
         Index           =   0
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Blank layer"
            Index           =   0
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Duplicate of current layer"
            Index           =   1
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "From clipboard"
            Index           =   3
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "From file..."
            Index           =   4
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Delete"
         Index           =   1
         Begin VB.Menu MnuLayerDelete 
            Caption         =   "Current layer"
            Index           =   0
         End
         Begin VB.Menu MnuLayerDelete 
            Caption         =   "Hidden layers"
            Index           =   1
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge up"
         Index           =   3
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge down"
         Index           =   4
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Order"
         Index           =   5
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Raise layer"
            Index           =   0
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Lower layer"
            Index           =   1
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Layer to top"
            Index           =   3
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Layer to bottom"
            Index           =   4
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Orientation"
         Index           =   7
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Straighten"
            Index           =   0
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 90° clockwise"
            Index           =   2
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 90° counter-clockwise"
            Index           =   3
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 180°"
            Index           =   4
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate arbitrary..."
            Index           =   5
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Flip horizontal"
            Index           =   7
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Flip vertical"
            Index           =   8
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Size"
         Index           =   8
         Begin VB.Menu MnuLayerSize 
            Caption         =   "Reset to actual size"
            Index           =   0
         End
         Begin VB.Menu MnuLayerSize 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuLayerSize 
            Caption         =   "Resize..."
            Index           =   2
         End
         Begin VB.Menu MnuLayerSize 
            Caption         =   "Content-aware resize..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Transparency"
         Index           =   10
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Add basic transparency..."
            Index           =   0
         End
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Make color transparent..."
            Index           =   1
         End
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Remove transparency..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Flatten image"
         Index           =   12
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge visible layers"
         Index           =   13
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
      Begin VB.Menu MnuSelect 
         Caption         =   "Export"
         Index           =   12
         Begin VB.Menu MnuSelectExport 
            Caption         =   "Selected area as image..."
            Index           =   0
         End
         Begin VB.Menu MnuSelectExport 
            Caption         =   "Selection mask as image..."
            Index           =   1
         End
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
            Caption         =   "Vibrance..."
            Index           =   4
         End
         Begin VB.Menu MnuColor 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Black and white..."
            Index           =   6
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Colorize..."
            Index           =   7
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Replace color..."
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
            Caption         =   "Gamma..."
            Index           =   2
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Levels..."
            Index           =   3
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Shadows and highlights..."
            Index           =   4
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Temperature..."
            Index           =   5
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
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Photography"
         Index           =   14
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Exposure..."
            Index           =   0
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Photo filters..."
            Index           =   1
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Split toning..."
            Index           =   2
         End
      End
   End
   Begin VB.Menu MnuEffectsTop 
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
            Caption         =   "Glass tiles..."
            Index           =   3
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Kaleiodoscope..."
            Index           =   4
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Modern art..."
            Index           =   5
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Oil painting..."
            Index           =   6
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Pencil drawing"
            Index           =   7
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Posterize..."
            Index           =   8
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Relief"
            Index           =   9
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
            Caption         =   "Surface blur..."
            Index           =   2
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Motion blur..."
            Index           =   4
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Radial blur..."
            Index           =   5
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Zoom blur..."
            Index           =   6
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Chroma blur..."
            Index           =   8
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Fragment..."
            Index           =   9
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Grid blur"
            Index           =   10
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Pixelate..."
            Index           =   11
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Distort"
         Index           =   2
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Apply lens distortion..."
            Index           =   0
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Correct existing lens distortion..."
            Index           =   1
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Miscellaneous..."
            Index           =   2
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Pan and zoom..."
            Index           =   3
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Perspective..."
            Index           =   4
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Pinch and whirl..."
            Index           =   5
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Poke..."
            Index           =   6
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Polar conversion..."
            Index           =   7
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Ripple..."
            Index           =   8
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Rotate..."
            Index           =   9
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Shear..."
            Index           =   10
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Spherize..."
            Index           =   11
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Squish..."
            Index           =   12
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Swirl..."
            Index           =   13
         End
         Begin VB.Menu MnuDistortEffects 
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
            Caption         =   "Sunshine"
            Index           =   7
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Underwater"
            Index           =   8
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
         Caption         =   "File toolbox"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Layers toolbox"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tools toolbox"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Image tabstrip"
         Index           =   3
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Always show"
            Index           =   0
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Show when multiple images are loaded"
            Index           =   1
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Never show"
            Index           =   2
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Left"
            Index           =   4
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Top"
            Index           =   5
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Right"
            Index           =   6
         End
         Begin VB.Menu MnuWindowTabstrip 
            Caption         =   "Bottom"
            Index           =   7
         End
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Floating toolboxes"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Next image"
         Index           =   7
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Previous image"
         Index           =   8
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

'PhotoDemon is Copyright ©1999-2014 by Tanner Helland, tannerhelland.com

'Please visit photodemon.org for updates and additional downloads

'***************************************************************************
'Main Program Form
'Copyright ©2002-2014 by Tanner Helland
'Created: 15/September/02
'Last updated: 21/May/14
'Last update: update interactions with Autosave handler to reflect recent Autosave engine overhaul
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

'An outside class provides access to specialized mouse events (like mousewheel and forward/back keys)
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Horizontal mousewheel; note that the pdInput class automatically converts Shift+Wheel to horizontal wheel for us
Private Sub cMouseEvents_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    Dim newX As Long, newY As Long
    
    If g_OpenImageCount > 0 Then
    
        'Mouse is over the tabstrip
        If g_MouseOverImageTabstrip Then
            
            'Convert the x/y coordinates we received into the child window's coordinate space, then relay the mousewheel message
            Drawing.convertCoordsBetweenHwnds Me.hWnd, toolbar_ImageTabs.hWnd, x, y, newX, newY
            toolbar_ImageTabs.cMouseEvents_MouseWheelHorizontal Button, Shift, newX, newY, scrollAmount
        
        'Assume mouse is over the canvas
        Else
        
            'Convert the x/y coordinates we received into the child window's coordinate space, then relay the mousewheel message
            Drawing.convertCoordsBetweenHwnds Me.hWnd, FormMain.mainCanvas(0).hWnd, x, y, newX, newY
            FormMain.mainCanvas(0).cMouseEvents_MouseWheelHorizontal Button, Shift, newX, newY, scrollAmount
            
        End If
        
    End If

End Sub

'Vertical mousewheel; note that the pdInput class automatically converts Shift+Wheel and Ctrl+Wheel actions to dedicated events,
' so this function will only return plain MouseWheel events (or Alt+MouseWheel, I suppose)
Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    Dim newX As Long, newY As Long

    If g_OpenImageCount > 0 Then
        
        'Mouse is over the image tabstrip
        If g_MouseOverImageTabstrip Then
            
            'Convert the x/y coordinates we received into the child window's coordinate space, then relay the mousewheel message
            Drawing.convertCoordsBetweenHwnds Me.hWnd, toolbar_ImageTabs.hWnd, x, y, newX, newY
            toolbar_ImageTabs.cMouseEvents_MouseWheelVertical Button, Shift, newX, newY, scrollAmount
        
        'Assume mouse is over the main canvas
        Else
            
            'Convert the x/y coordinates we received into the child window's coordinate space, then relay the mousewheel message
            Drawing.convertCoordsBetweenHwnds Me.hWnd, FormMain.mainCanvas(0).hWnd, x, y, newX, newY
            FormMain.mainCanvas(0).cMouseEvents_MouseWheelVertical Button, Shift, newX, newY, scrollAmount
            
        End If
        
    End If

End Sub

'Ctrl+Wheel actions are detected by pdInput and sent to this dedicated class
Private Sub cMouseEvents_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)

    'The only child window that supports mousewheel zoom is the main canvas, so redirect any zoom events there.
    If g_OpenImageCount > 0 Then
    
        Dim newX As Long, newY As Long
    
        'Convert the x/y coordinates we received into the child window's coordinate space, then relay the mousewheel message
        Drawing.convertCoordsBetweenHwnds Me.hWnd, FormMain.mainCanvas(0).hWnd, x, y, newX, newY
        FormMain.mainCanvas(0).cMouseEvents_MouseWheelZoom Button, Shift, newX, newY, zoomAmount
    
    End If

End Sub

'When the main form is resized, we must re-align the main canvas
Private Sub Form_Resize()
    
    refreshAllCanvases

End Sub

'Resize all currently active canvases
Public Sub refreshAllCanvases()

    'If the main form has been minimized, don't refresh anything
    If FormMain.WindowState = vbMinimized Then Exit Sub

    'Start by reorienting the canvas to fill the full available client area
    Dim mainRect As winRect
    
    g_WindowManager.getActualMainFormClientRect mainRect, False, False
    
    'mainCanvas(0).Move mainRect.x1, mainRect.y1, mainRect.x2 - mainRect.x1, mainRect.y2 - mainRect.y1
    MoveWindow mainCanvas(0).hWnd, mainRect.x1, mainRect.y1, mainRect.x2 - mainRect.x1, mainRect.y2 - mainRect.y1, 1
    mainCanvas(0).fixChromeLayout
    
    'Refresh any current windows
    If g_OpenImageCount > 0 Then
        PrepareViewport pdImages(g_CurrentImage), mainCanvas(0), "Form_Resize(" & Me.ScaleWidth & "," & Me.ScaleHeight & ")"
    End If
    
End Sub

'Menu: Adjustments -> Photography
Private Sub MnuAdjustmentsPhoto_Click(Index As Integer)

    Select Case Index
    
        'Exposure
        Case 0
            Process "Exposure", True
    
        'Photo filters
        Case 1
            Process "Photo filter", True
            
        'Split-toning
        Case 2
            Process "Split toning", True
    
    End Select

End Sub

'Menu: top-level layer actions
Private Sub MnuLayer_Click(Index As Integer)

    Select Case Index
    
        'Add layer (top-level)
        Case 0
        
        'Delete layer (top-level)
        Case 1
        
        '<separator>
        Case 2
        
        'Merge up
        Case 3
            Process "Merge layer up", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        
        'Merge down
        Case 4
            Process "Merge layer down", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        
        'Order (top-level)
        Case 5
        
        '<separator>
        Case 6
        
        'Orientation (top-level)
        Case 7
        
        'Size (top-level)
        Case 8
        
        '<separator>
        Case 9
        
        'Transparency (top-level)
        Case 10
        
        '<separator>
        Case 11
        
        'Flatten layers
        Case 12
            Process "Flatten image", , , UNDO_IMAGE
        
        'Merge visible layers
        Case 13
            Process "Merge visible layers", , , UNDO_IMAGE
        
    End Select

End Sub

'Menu: remove layers from the image
Private Sub MnuLayerDelete_Click(Index As Integer)

    Select Case Index
    
        'Delete current layer
        Case 0
            Process "Delete layer", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        
        'Delete all hidden layers
        Case 1
            Process "Delete hidden layers", False, , UNDO_IMAGE
        
    End Select

End Sub

'Menu: add a layer to the image
Private Sub MnuLayerNew_Click(Index As Integer)

    Select Case Index
    
        'Blank layer
        Case 0
            Process "Add blank layer", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        
        'Duplicate of current layer
        Case 1
            Process "Duplicate Layer", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
        
        '<separator>
        Case 2
        
        'Import from clipboard
        Case 3
            Process "Paste as new layer", False, , UNDO_IMAGE, , False
        
        'Import from file
        Case 4
            Process "New layer from file", True
    
    End Select

End Sub

'Menu: change layer order
Private Sub MnuLayerOrder_Click(Index As Integer)

    Select Case Index
    
        'Raise layer
        Case 0
            Process "Raise layer", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGEHEADER
        
        'Lower layer
        Case 1
            Process "Lower layer", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGEHEADER
        
        '<separator>
        Case 2
        
        'Raise to top
        Case 3
            Process "Raise layer to top", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGEHEADER
        
        'Lower to bottom
        Case 4
            Process "Lower layer to bottom", False, Str(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGEHEADER
        
    End Select

End Sub

Private Sub MnuLayerOrientation_Click(Index As Integer)

    Select Case Index
    
        'Straighten
        Case 0
            Process "Straighten layer", True
        
        '<separator>
        Case 1
        
        'Rotate 90
        Case 2
            Process "Rotate layer 90° clockwise", , , UNDO_LAYER
        
        'Rotate 270
        Case 3
            Process "Rotate layer 90° counter-clockwise", , , UNDO_LAYER
        
        'Rotate 180
        Case 4
            Process "Rotate layer 180°", , , UNDO_LAYER
        
        'Rotate arbitrary
        Case 5
            Process "Arbitrary layer rotation", True
        
        '<separator>
        Case 6
        
        'Flip horizontal
        Case 7
            Process "Flip layer horizontally", , , UNDO_LAYER
        
        'Flip vertical
        Case 8
            Process "Flip layer vertically", , , UNDO_LAYER
    
    End Select

End Sub

Private Sub MnuLayerSize_Click(Index As Integer)

    Select Case Index
    
        'Reset to actual size
        Case 0
            Process "Reset layer size", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_LAYERHEADER
        
        '<separator>
        Case 1
            
        'Standard resize
        Case 2
            Process "Resize layer", True
        
        'Content-aware resize
        Case 3
            Process "Content-aware resize", True
    
    End Select

End Sub

Private Sub shellPipeMain_ErrDataArrival(ByVal CharsTotal As Long)

    Debug.Print "WARNING!  Asynchronous pipe source returned the following data on stderr: "
    Debug.Print shellPipeMain.ErrGetData()
    Debug.Print " -- End stderr output -- "

End Sub

'Append any new data to our master metadata string
Private Sub shellPipeMain_DataArrival(ByVal CharsTotal As Long)
    
    Dim receivedData As String
    receivedData = shellPipeMain.GetData()
    
    newMetadataReceived receivedData
    
    'DEBUG ONLY!
    'Debug.Print "Received " & LenB(receivedData) & " bytes of new data from ExifTool."
    'Debug.Print receivedData
    
End Sub

'Countdown timer for re-enabling disabled user input.  A delay is enforced to prevent double-clicks on child dialogs from
' "passing through" to the main form and causing goofy behavior.
Private Sub tmrCountdown_Timer()

    Static intervalCount As Long
    
    If intervalCount > 2 Then
        
        intervalCount = 0
        g_DisableUserInput = False
        tmrCountdown.Enabled = False
        
    End If
    
    intervalCount = intervalCount + 1

End Sub

'When download of the update information is complete, write out the current date to the preferences file
Private Sub updateChecker_Complete()
    Debug.Print "Update file download complete.  Update information has been saved at " & g_UserPreferences.getDataPath & "updates.xml"
    g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
End Sub

'Forward mousewheel events to the relevant window
Private Sub cMouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)



End Sub

Private Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)



End Sub

'THE BEGINNING OF EVERYTHING
' Actually, Sub "Main" in the module "modMain" is loaded first, but all it does is set up native theming.  Once it has done that, FormMain is loaded.
Private Sub Form_Load()

    '*************************************************************************************************************************************
    ' Before doing anything else, store the command line to memory
    '*************************************************************************************************************************************

    'Use a global variable to store any command line parameters we may have been passed
    g_CommandLine = Command$
    
    'Instantiate the themed tooltip class
    Set m_ToolTip = New clsToolTip
    
    'Create a blank pdImages() array, to avoid errors
    ReDim pdImages(0) As pdImage
    
    '*************************************************************************************************************************************
    ' Reroute control to "LoadTheProgram", which initializes all key PD systems
    '*************************************************************************************************************************************
    
    'The bulk of the loading code actually takes place inside the LoadTheprogram subroutine (which can be found in the "Loading" module)
    Loading.LoadTheProgram
    
    
    '*************************************************************************************************************************************
    ' Now that all engines are initialized, prep and display the main editing window
    '*************************************************************************************************************************************
    
    'We can now display the main form and any visible toolbars.  (There is currently a flicker if toolbars have been hidden by the user,
    ' and I'm working on a solution to that.)
    Me.Visible = True
    
    'Register all toolbox forms with the window manager
    g_WindowManager.registerChildForm toolbar_File, TOOLBAR_WINDOW, 1, FILE_TOOLBOX
    g_WindowManager.registerChildForm toolbar_Layers, TOOLBAR_WINDOW, 2, LAYER_TOOLBOX
    g_WindowManager.registerChildForm toolbar_Tools, TOOLBAR_WINDOW, 3, TOOLS_TOOLBOX
    
    g_WindowManager.registerChildForm toolbar_ImageTabs, IMAGE_TABSTRIP, , , , , 32
    
    'Display the various toolboxes per the user's display settings
    toolbar_File.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_File.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show File Toolbox", True)
    toolbar_Tools.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_Tools.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True)
    toolbar_Layers.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_Layers.hWnd, g_UserPreferences.GetPref_Boolean("Core", "Show Layers Toolbox", True)
    
    'We only display the image tab manager now if the user loaded two or more images from the command line
    toolbar_ImageTabs.Show vbModeless, Me
    g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, IIf(g_OpenImageCount > 1, True, False)
    
    'Synchronize the main canvas layout
    refreshAllCanvases
    
    'Enable mouse subclassing for events like mousewheel, forward/back keys, enter/leave
    Set cMouseEvents = New pdInput
    cMouseEvents.addInputTracker Me.hWnd
    
    
    '*************************************************************************************************************************************
    ' Next, make sure PD's previous session closed down successfully
    '*************************************************************************************************************************************
    
    Message "Checking for old autosave data..."
    
    If Not Image_Autosave_Handler.wasLastShutdownClean Then
    
        'Oh no!  Something went horribly wrong with the last PD session.  See if there's any AutoSave data worth recovering.
        If Image_Autosave_Handler.saveableImagesPresent > 0 Then
        
            'Autosave data was found!  Present it to the user.
            Dim userWantsAutosaves As VbMsgBoxResult
            Dim listOfFilesToSave() As AutosaveXML
            
            userWantsAutosaves = displayAutosaveWarning(listOfFilesToSave)
            
            'If the user wants to restore old Autosave data, do so now.
            If (userWantsAutosaves = vbYes) Then
            
                'listOfFilesToSave contains the list of Autosave files the user wants restored.
                ' Hand them off to the autosave handler, which will load and restore each file in turn.
                Image_Autosave_Handler.loadTheseAutosaveFiles listOfFilesToSave
                
                'Synchronize the interface to the restored files
                syncInterfaceToCurrentImage
                
                'With all data successfully loaded, purge the now-unnecessary Autosave entries.
                Image_Autosave_Handler.purgeOldAutosaveData
            
            Else
                
                'The user has no interest in recovering AutoSave data.  Purge all the entries we found, so they don't show
                ' up in future AutoSave searches.
                Image_Autosave_Handler.purgeOldAutosaveData
            
            End If
            
        
        'There's not any AutoSave data worth recovering.  Ask the user to submit a bug report??
        Else
        
            'TODO
        
        End If
    
    End If
    
    
    
    '*************************************************************************************************************************************
    ' Next, analyze the command line and load any image files (if present).
    '*************************************************************************************************************************************
    
    Message "Checking command line..."
    
    If Len(g_CommandLine) > 0 Then
        Message "Loading requested images..."
        Loading.LoadImagesFromCommandLine
    End If
    
    
    '*************************************************************************************************************************************
    ' Next, see if we need to display the language selection dialog (NOT IMPLEMENTED AT PRESENT)
    '*************************************************************************************************************************************
    
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
    
    
    '*************************************************************************************************************************************
    ' Next, see if we need to launch an asynchronous check for updates
    '*************************************************************************************************************************************
    
    'Start by seeing if we're allowed to check for software updates (the user can disable this check, and we want to honor their selection)
    Dim allowedToUpdate As Boolean
    allowedToUpdate = Software_Updater.isItTimeForAnUpdate()
    
    'If we're STILL allowed to update, do so now (unless this is the first time the user has run the program; in that case, suspend updates,
    ' as it is assumed the user already has an updated copy of the software - and we don't want to bother them already!)
    If allowedToUpdate Then
    
        Message "Initializing software updater (this feature can be disabled from the Tools -> Options menu)..."
        
        'Prior to January 2014, PhotoDemon updates were downloaded synchronously (meaning all other PD operations were
        ' suspended until the update completed).  In Jan 14, I moved to an asynchronous method, which means that updates
        ' are downloaded transparently in the background.
        
        'This means that all we need to do at this stage is initiate the update file download (from its standard location
        ' of photodemon.org/downloads/updates.xml).  If successful, the downloader will place the completed update file
        ' in the /Data subfolder.  On exit (or subsequent program runs), we can simply check for the presence of that file.
        FormMain.updateChecker.Download "http://photodemon.org/downloads/updates.xml", g_UserPreferences.getDataPath & "updates.xml", vbAsyncReadForceUpdate
                
    End If
    
    
    '*************************************************************************************************************************************
    ' Next, see if an update was previously loaded; if it was, display any relevant findings.
    '*************************************************************************************************************************************
    
    'It's possible that a past program instance downloaded update information for us; check for an update file now.
    ' (Note that this check can be skipped the first time the program is run, as we are guaranteed to not have update data yet!)
    If (Not g_IsFirstRun) Then
    
        Message "Checking for previously downloaded update data..."
    
        Dim updateNeeded As UpdateCheck
        updateNeeded = CheckForSoftwareUpdate
        
        'CheckForSoftwareUpdate can return one of four values:
        ' UPDATE_ERROR - something went wrong
        ' UPDATE_NOT_NEEDED - an update file was found, but the current software version is already up-to-date
        ' UPDATE_AVAILABLE - an update file was found, and an updated PD copy is available
        ' UPDATE_UNAVAILABLE - no update file was found (this is the most common occurrence, as updates are only checked every 10 days)
        
        Select Case updateNeeded
        
            Case UPDATE_ERROR
                Message "An error occurred while looking for an update file."
            
            Case UPDATE_NOT_NEEDED
                Message "Update data found, but this copy of PhotoDemon is already up-to-date."
                
                'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                
            Case UPDATE_AVAILABLE
                Message "New PhotoDemon update found!  Launching update notifier..."
                showPDDialog vbModal, FormSoftwareUpdate
                
            Case Else
                'No update data found - which is fine!  (This is actually the most common occurrence.)
            
        End Select
            
    End If
    
    
    '*************************************************************************************************************************************
    ' Next, check for missing core plugins
    '*************************************************************************************************************************************
    
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
            showPDDialog vbModal, FormPluginDownloader
            
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
    
    
    '*************************************************************************************************************************************
    ' Display any final messages and/or warnings
    '*************************************************************************************************************************************
        
    Message ""
    
    'TODO: As of 27 April '14, I have removed the warning below.
    'MsgBox "WARNING!  I am currently adding Layers support to PhotoDemon.  Because Layers are only partially complete, the program is extremely unstable, with many features completely broken." & vbCrLf & vbCrLf & "As long as this message remains, PhotoDemon may not function properly (or at all).  I've suspended nightly builds until things are stable.  If you've manually downloaded this build from GitHub, consider yourself warned." & vbCrLf & vbCrLf & "(Seriously: please do any editing with with the 6.2 stable release, available from photodemon.org)", vbExclamation + vbOKOnly + vbApplicationModal, "6.4 Development Warning"
    
    
    '*************************************************************************************************************************************
    ' For developers only, display an IDE avoidance warning (if it hasn't been dismissed before).
    '*************************************************************************************************************************************
    
    'Because people may be using this code in the IDE, warn them about the consequences of doing so
    If (Not g_IsProgramCompiled) And (g_UserPreferences.GetPref_Boolean("Core", "Display IDE Warning", True)) Then displayIDEWarning
    
    'Finally, return focus to the main form
    FormMain.SetFocus
     
End Sub

'Allow the user to drag-and-drop files and URLs onto the main form
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    Clipboard_Handler.loadImageFromDragDrop Data, Effect, False
    
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Or Data.GetFormat(vbCFText) Then
        'Inform the source that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files or text, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'If the user is attempting to close the program, run some checks.  Specifically, we want to make sure all child forms have been saved.
' Note: in VB6, the order of events for program closing is MDI Parent QueryUnload, MDI children QueryUnload, MDI children Unload, MDI Unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    'If the histogram form is open, close it
    'Unload FormHistogram
    
    'Store the main window's location to file now.  We will use this in the future to determine which monitor
    ' to display the splash screen on
    g_UserPreferences.SetPref_Long "Core", "Last Window State", Me.WindowState
    g_UserPreferences.SetPref_Long "Core", "Last Window Left", Me.Left / TwipsPerPixelXFix
    g_UserPreferences.SetPref_Long "Core", "Last Window Top", Me.Top / TwipsPerPixelYFix
    g_UserPreferences.SetPref_Long "Core", "Last Window Width", Me.Width / TwipsPerPixelXFix
    g_UserPreferences.SetPref_Long "Core", "Last Window Height", Me.Height / TwipsPerPixelYFix
    
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
                    QueryUnloadPDImage Cancel, UnloadMode, i
                    
                    If Not CBool(Cancel) Then UnloadPDImage Cancel, i
                    
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
    
    'Release GDIPlus (if applicable)
    If g_ImageFormats.GDIPlusEnabled Then releaseGDIPlus
    
    'Release ExifTool (if available)
    If g_ExifToolEnabled Then terminateExifTool
    
    'Releast FreeImage (if available)
    If Not (g_FreeImageHandle = 0) Then FreeLibrary g_FreeImageHandle
    
    'Stop tracking hotkeys
    ctlAccelerator.Enabled = False
    
    'Destroy all custom-created form icons
    destroyAllIcons
    
    'Release the hand cursor we use for all clickable objects
    unloadAllCursors

    'Save the MRU list to the preferences file.  (I've considered doing this as files are loaded, but the only time
    ' that would be an improvement is if the program crashes, and if it does crash, the user wouldn't want to re-load
    ' the problematic image anyway.)
    g_RecentFiles.MRU_SaveToFile
        
    'Restore the user's font smoothing setting as necessary
    handleClearType False
    
    'Release any Win7-specific features
    releaseWin7Features
    
    ReleaseFormTheming Me
    
    'Release this form from the window manager, and write out all window data to file
    g_WindowManager.unregisterForm Me
    g_WindowManager.saveAllWindowLocations
    
    'As a final failsafe, forcibly unload any remaining forms
    Dim tmpForm As Form
    For Each tmpForm In Forms
        
        'Note that there is no need to unload FormMain, as we're about to unload it anyway!
        If tmpForm.Name <> "FormMain" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
        
    Next tmpForm
    
    'The very last thing we do before terminating is notify the Autosave handler that everything shut down correctly
    Image_Autosave_Handler.purgeOldAutosaveData
    Image_Autosave_Handler.notifyCleanShutdown
    
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
            Process "Comic book", , , UNDO_LAYER
            
        'Figured glass
        Case 1
            Process "Figured glass", True
            
        'Film noir
        Case 2
            Process "Film noir", , , UNDO_LAYER
        
        'Glass tiles
        Case 3
            Process "Glass tiles", True
        
        'Kaleidoscope
        Case 4
            Process "Kaleidoscope", True
        
        'Modern art
        Case 5
            Process "Modern art", True
        
        'Oil painting
        Case 6
            Process "Oil painting", True
            
        'Pencil drawing
        Case 7
            Process "Pencil drawing", , , UNDO_LAYER
                
        'Posterize
        Case 8
            Process "Posterize", True
            
        'Relief
        Case 9
            Process "Relief", , , UNDO_LAYER
    
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
        
        'Surface Blur
        Case 2
            Process "Surface blur", True
        
        '<separator>
        Case 3
        
        'Motion blur
        Case 4
            Process "Motion blur", True
        
        'Radial blur
        Case 5
            Process "Radial blur", True
        
        'Zoom Blur
        Case 6
            Process "Zoom blur", True
            
        '<separator>
        Case 7
        
        'Chroma blur
        Case 8
            Process "Chroma blur", True
        
        'Fragment
        Case 9
            Process "Fragment", True
                
        'Grid blur
        Case 10
            Process "Grid blur", , , UNDO_LAYER
            
        'Pixelate (mosaic)
        Case 11
            Process "Pixelate", True
            
    End Select

End Sub

Private Sub MnuClearMRU_Click()
    g_RecentFiles.MRU_ClearList
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
        
        'Vibrance
        Case 4
            Process "Vibrance", True
        
        '<separator>
        Case 5
        
        'Grayscale (black and white)
        Case 6
            Process "Black and white", True
        
        'Colorize
        Case 7
            Process "Colorize", True
            
        'Replace color
        Case 8
            Process "Replace color", True
                
        'Sepia
        Case 9
            Process "Sepia", , , UNDO_LAYER

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
            Process "Maximum channel", , , UNDO_LAYER
        
        'Min channel
        Case 4
            Process "Minimum channel", , , UNDO_LAYER
            
        '<separator>
        Case 5
        
        'Shift colors left
        Case 6
            Process "Shift colors (left)", , , UNDO_LAYER
            
        'Shift colors right
        Case 7
            Process "Shift colors (right)", , , UNDO_LAYER
        
    End Select
    
End Sub

Private Sub MnuCompoundInvert_Click()
    Process "Compound invert", False, buildParams("128"), UNDO_LAYER
End Sub

Private Sub MnuCustomFilter_Click()
    Process "Custom filter", True
End Sub

'All distortion filters happen here
Private Sub MnuDistortEffects_Click(Index As Integer)

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
    Process "Dream", , , UNDO_LAYER
End Sub

Private Sub MnuEdge_Click(Index As Integer)

    Select Case Index
        
        'Emboss/engrave
        Case 0
            Process "Emboss or engrave", True
         
        'Enhance edges
        Case 1
            Process "Edge enhance", , , UNDO_LAYER
        
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
            Process "Undo", False
        
        'Redo
        Case 1
            Process "Redo", False
        
        'Repeat last
        Case 2
            'TODO: figure out Undo handling for "Repeat last action"
            Process "Repeat last action", False, , UNDO_IMAGE
        
        '<separator>
        Case 3
        
        'Copy default to clipboard
        Case 4
            Process "Copy", False, , UNDO_NOTHING, , False
        
        'Copy merged area to clipboard
        Case 5
            Process "Copy merged", False, , UNDO_NOTHING, , False
        
        'Paste as new layer
        Case 6
            Process "Paste as new layer", False, , UNDO_IMAGE, , False
        
        'Paste as new image
        Case 7
            Process "Paste as new image", False, , UNDO_NOTHING, , False
        
        '<separator>
        Case 8
        
        'Empty clipboard
        Case 9
            Process "Empty clipboard", False, , UNDO_NOTHING, , False
                
    
    End Select
    
End Sub

Private Sub MnuFadeLastEffect_Click()
    Process "Fade last effect", , , UNDO_LAYER
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
        
        'Close
        Case 4
            Process "Close", True
        
        'Close all
        Case 5
            Process "Close all", True
            
        '<separator>
        Case 6
        
        'Save
        Case 7
            Process "Save", True
            
        'Save as
        Case 8
            Process "Save as", True
        
        'Revert
        Case 9
            'TODO: figure out correct Undo behavior for REVERT action
            Process "Revert", False, , UNDO_NOTHING
        
        '<separator>
        Case 10
        
        'Batch wizard
        Case 11
            Process "Batch wizard", True
        
        '<separator>
        Case 12
        
        'Print
        Case 13
            Process "Print", True
            
        '<separator>
        Case 14
        
        'Exit
        Case 15
            Process "Exit program", True
        
    
    End Select
    
End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuHeatmap_Click()
    Process "Thermograph (heat map)", , , UNDO_LAYER
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
            updateNeeded = CheckForSoftwareUpdate(True)
    
            'CheckForSoftwareUpdate can return one of three values:
            ' 0 - something went wrong (no Internet connection, etc)
            ' 1 - the check was successful, but this version is up-to-date
            ' 2 - the check was successful, and an update is available
            Select Case updateNeeded
        
                Case 0
                    pdMsgBox "An error occurred while checking for updates.  Please try again later.", vbOKOnly + vbInformation + vbApplicationModal, "PhotoDemon Updates"
                    Message "Software update check postponed."
                    
                Case 1
                    pdMsgBox "This copy of PhotoDemon is the newest version available." & vbCrLf & vbCrLf & "(Current version: %1.%2.%3)", vbOKOnly + vbInformation + vbApplicationModal, "PhotoDemon Updates", App.Major, App.Minor, App.Revision
                    Message "This copy of PhotoDemon is up to date."
                        
                    'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                    g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                        
                Case 2
                    Message "Software update found!  Launching update notifier..."
                    showPDDialog vbModal, FormSoftwareUpdate
                    
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
            showPDDialog vbModal, FormAbout
        
    End Select

End Sub

Private Sub MnuHistogram_Click()
    'Process "Display histogram", True
    showPDDialog vbModal, FormHistogram
End Sub

Private Sub MnuHistogramEqualize_Click()
    Process "Equalize", True
End Sub

Private Sub MnuHistogramStretch_Click()
    Process "Stretch histogram", , , UNDO_LAYER
End Sub

'All top-level Image menu actions are handled here
Private Sub MnuImage_Click(Index As Integer)

    Select Case Index
    
        'Duplicate
        Case 0
            Process "Duplicate image", , , UNDO_NOTHING
        
        '<separator>
        Case 1
        
        'Resize
        Case 2
            Process "Resize image", True
        
        '<separator>
        Case 3
        
        'Canvas resize
        Case 4
            Process "Canvas size", True
            
        'Fit canvas to active layer
        Case 5
            Process "Fit canvas to layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGEHEADER
        
        'Fit canvas around all layers
        Case 6
            Process "Fit canvas to all layers", False, , UNDO_IMAGEHEADER
            
        '<separator>
        Case 7
            
        'Crop to selection
        Case 8
            Process "Crop", , , UNDO_IMAGE
        
        'Trim empty borders
        Case 9
            Process "Trim empty borders", , , UNDO_IMAGEHEADER
        
        '<separator>
        Case 10
        
        'Top-level Rotate
        Case 11
        
        'Flip horizontal (mirror)
        Case 12
            Process "Flip image horizontally", , , UNDO_IMAGE
        
        'Flip vertical
        Case 13
            Process "Flip image vertically", , , UNDO_IMAGE
        
        'NOTE: isometric view was removed in 6.4.  I may include it at a later date if there is demand.
        'Isometric view
        'Case 12
        '    Process "Isometric conversion"
            
        '<separator>
        Case 14
        
        'Indexed color
        Case 15
            Process "Reduce colors", True
        
        'Tile
        Case 16
            Process "Tile", True
            
        '<separator>
        Case 17
        
        'Metadata top-level
        Case 18
    
    End Select

End Sub

'This is the exact same thing as "Paste as New Image".  It is provided in two locations for convenience.
Private Sub MnuImportClipboard_Click()
    Process "Paste as new image", False, , UNDO_NOTHING, , False
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    Process "Internet import", True
End Sub

Private Sub MnuAlien_Click()
    Process "Alien", , , UNDO_LAYER
End Sub

Private Sub MnuInvertHue_Click()
    Process "Invert hue", , , UNDO_LAYER
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    
    'Because loading a language can take some time, display a wait screen to discourage attempted interaction
    displayWaitScreen g_Language.TranslateMessage("Please wait while the new language is applied..."), Me
    
    'Remove the existing translation from any visible windows
    Message "Removing existing translation..."
    g_Language.undoTranslations FormMain, True
    g_Language.undoTranslations toolbar_File, True
    g_Language.undoTranslations toolbar_ImageTabs, True
    g_Language.undoTranslations toolbar_Tools, True
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.activateNewLanguage Index, True
    
    Message "Language changed successfully."
    
    hideWaitScreen
    
    Screen.MousePointer = vbDefault
    
    'Added 09 January 2014.  Let the user know that some translations will not take affect until the program is restarted.
    pdMsgBox "Language changed successfully!" & vbCrLf & vbCrLf & "Note: some minor program text (such as hover tooltips) cannot be live-updated.  Such text will be properly translated the next time you start the application.", vbApplicationModal + vbOKOnly + vbInformation, "Language changed successfully"
    
End Sub

Private Sub MnuLighting_Click(Index As Integer)

    Select Case Index
            
        'Brightness/Contrast
        Case 0
            Process "Brightness and contrast", True
        
        'Curves
        Case 1
            Process "Curves", True
            
        'Gamma correction
        Case 2
            Process "Gamma", True
            
        'Levels
        Case 3
            Process "Levels", True

        'Shadows/Midtones/Highlights
        Case 4
            Process "Shadows and highlights", True
            
        'Temperature
        Case 5
            Process "Temperature", True
    
    End Select

End Sub

'Load all images in the current "Recent Files" menu
Private Sub MnuLoadAllMRU_Click()
    
    'Fill a string array with all current MRU entries
    Dim sFile() As String
    ReDim sFile(0 To g_RecentFiles.MRU_ReturnCount() - 1) As String
    
    Dim i As Long
    For i = 0 To UBound(sFile)
        sFile(i) = g_RecentFiles.getSpecificMRU(i)
    Next i
    
    'Load all images in the list
    LoadFileAsNewImage sFile
    
    'If the image loaded successfully, activate it and bring it to the foreground
    If g_OpenImageCount > 0 Then activatePDImage g_CurrentImage, "finished loading all recent images"
    
End Sub

'All metadata sub-menu options are handled here
Private Sub MnuMetadata_Click(Index As Integer)

    Select Case Index
    
        'Browse metadata
        Case 0
        
            'Before doing anything else, see if we've already loaded metadata.  If we haven't, do so now.
            If Not pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata Then
                'pdImages(g_CurrentImage).imgMetadata.loadAllMetadata pdImages(g_CurrentImage).locationOnDisk, pdImages(g_CurrentImage).originalFileFormat
                
                'Update the interface to reflect any changes to the metadata menu (for example, if we found GPS data
                ' during the metadata load process)
                syncInterfaceToCurrentImage
            End If
            
            'If the image STILL doesn't have metadata, warn the user and exit.
            If Not pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata Then
                Message "No metadata available."
                pdMsgBox "This image does not contain any metadata.", vbInformation + vbOKOnly + vbApplicationModal, "No metadata available"
                Exit Sub
            End If
            
            showPDDialog vbModal, FormMetadata
        
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
                'pdImages(g_CurrentImage).imgMetadata.loadAllMetadata pdImages(g_CurrentImage).locationOnDisk, pdImages(g_CurrentImage).originalFileFormat
                
                'Determine whether metadata is present, and dis/enable metadata menu items accordingly
                syncInterfaceToCurrentImage
            
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
            Process "Atmosphere", , , UNDO_LAYER
            
        'Burn
        Case 1
            Process "Burn", , , UNDO_LAYER
        
        'Fog
        Case 2
            Process "Fog", , , UNDO_LAYER
        
        'Freeze
        Case 3
            Process "Freeze", , , UNDO_LAYER
        
        'Lava
        Case 4
            Process "Lava", , , UNDO_LAYER
                
        'Rainbow
        Case 5
            Process "Rainbow", , , UNDO_LAYER
        
        'Steel
        Case 6
            Process "Steel", , , UNDO_LAYER
        
        'Sunshine
        Case 7
            Process "Sunshine", True
        
        'Water
        Case 8
            Process "Water", , , UNDO_LAYER
    
    End Select

End Sub

Private Sub MnuNegative_Click()
    Process "Film negative", , , UNDO_LAYER
End Sub

Private Sub MnuInvert_Click()
    Process "Invert RGB", , , UNDO_LAYER
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
    Process "Radioactive", , , UNDO_LAYER
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Load the MRU path that correlates to this index.  (If one is not found, a null string is returned)
    Dim tmpString As String
    tmpString = g_RecentFiles.getSpecificMRU(Index)
    
    'Check - just in case - to make sure the path isn't empty
    If tmpString <> "" Then
        
        'Message "Preparing to load recent file entry..."
        
        'Because LoadFileAsNewImage requires a string array, create an array to pass it
        Dim sFile(0) As String
        sFile(0) = tmpString
        
        LoadFileAsNewImage sFile
    End If
    
    'If the image loaded successfully, activate it and bring it to the foreground
    If g_OpenImageCount > 0 Then activatePDImage g_CurrentImage, "MRU entry finished loading"
    
End Sub

'All rotation actions are initiated here
Private Sub MnuRotate_Click(Index As Integer)

    Select Case Index
    
        'Straighten
        Case 0
            Process "Straighten image", True
        
        '<separator>
        Case 1
        
        'Rotate 90
        Case 2
            Process "Rotate image 90° clockwise", , , UNDO_IMAGE
        
        'Rotate 270
        Case 3
            Process "Rotate image 90° counter-clockwise", , , UNDO_IMAGE
        
        'Rotate 180
        Case 4
            Process "Rotate image 180°", , , UNDO_IMAGE
        
        'Rotate arbitrary
        Case 5
            Process "Arbitrary image rotation", True
            
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
            Process "Select all", , , UNDO_SELECTION, 0
        
        'Select none
        Case 1
            Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION
        
        'Invert
        Case 2
            Process "Invert selection", , , UNDO_SELECTION
        
        '<separator>
        Case 3
        
        'Grow selection
        Case 4
            Process "Grow selection", True
        
        'Shrink selection
        Case 5
            Process "Shrink selection", True
        
        'Border selection
        Case 6
            Process "Border selection", True
        
        'Feather selection
        Case 7
            Process "Feather selection", True
        
        'Sharpen selection
        Case 8
            Process "Sharpen selection", True
        
        '<separator>
        Case 9
        
        'Load selection
        Case 10
            Process "Load selection", True
        
        'Save current selection
        Case 11
            Process "Save selection", True
            
        '<Export top-level>
        Case 12
            
    End Select

End Sub

'All Select -> Export menu items are handled here
Private Sub MnuSelectExport_Click(Index As Integer)

    Select Case Index
    
        'Export selected area as image
        Case 0
            Process "Export selected area as image", True
        
        'Export selection mask itself as image
        Case 1
            Process "Export selection mask as image", True
    
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

    'Only attempt to change zoom if the primary zoom box is not currently disabled
    If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then

        Select Case Index
        
            Case 0
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 2
            Case 1
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 4
            Case 2
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 8
            Case 3
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 10
            Case 4
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getZoom100Index
            Case 5
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 14
            Case 6
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 16
            Case 7
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 19
            Case 8
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 21
                
        End Select

    End If

End Sub

Private Sub MnuStartMacroRecording_Click()
    Process "Start macro recording", , , UNDO_NOTHING
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process "Stop macro recording", True
End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)

    Select Case Index
    
        'Antique
        Case 0
            Process "Antique", , , UNDO_LAYER
    
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
    Process "Synthesize", , , UNDO_LAYER
End Sub

Private Sub MnuTest_Click()
    
    showPDDialog vbModal, FormGlassTiles
    
    MenuTest
End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)

    Select Case Index
    
        'Language editor
        Case 1
            If Not FormLanguageEditor.Visible Then showPDDialog vbModal, FormLanguageEditor
    
        'Options
        Case 5
            If Not FormPreferences.Visible Then showPDDialog vbModal, FormPreferences
            
        'Plugin manager
        Case 6
            If Not FormPluginManager.Visible Then showPDDialog vbModal, FormPluginManager
            
    End Select

End Sub

'Add / Remove / Modify a layer's alpha channel with this menu
Private Sub MnuLayerTransparency_Click(Index As Integer)

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

            'TODO: reevaluate the wisdom of having this option in the Image menu, vs a dedicated Layers menu
            'Ignore if the current image is already in 24bpp mode
            'If pdImages(g_CurrentImage).mainDIB.getDIBColorDepth = 24 Then Exit Sub
            Process "Remove alpha channel", True
    
    End Select

End Sub

Private Sub MnuVibrate_Click()
    Process "Vibrate", , , UNDO_LAYER
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
            showPDDialog vbModal, FormPreferences
            Exit Sub
        End If
    End If
    
    If ctlAccelerator.Key(nIndex) = "Plugin manager" Then
        If Not FormPluginManager.Visible Then
            showPDDialog vbModal, FormPluginManager
            Exit Sub
        End If
    End If
        
    'Escape - a separate function is used to cancel currently running filters.  This accelerator is only used
    ' to cancel batch conversions, but in the future it should be applied elsewhere.
    'If ctlAccelerator.Key(nIndex) = "Escape" Then
    '    If MacroStatus = MacroBATCH Then MacroStatus = MacroCANCEL
    'End If
    
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
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex > 0 Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex - 1
    End If
    
    'Zoom out
    If ctlAccelerator.Key(nIndex) = "Zoom_Out" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex < (FormMain.mainCanvas(0).getZoomDropDownReference().ListCount - 1) Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex + 1
    End If
    
    'Actual size
    If ctlAccelerator.Key(nIndex) = "Actual_Size" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getZoom100Index
    End If
    
    'Various zoom values
    If ctlAccelerator.Key(nIndex) = "Zoom_161" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 2
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_81" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 4
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_41" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 8
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_21" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 10
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_12" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 14
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_14" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 16
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_18" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 19
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_116" Then
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = 21
    End If
    
    'Remove selection
    If ctlAccelerator.Key(nIndex) = "Remove selection" Then
        Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION
    End If
    
    'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
    If ctlAccelerator.Key(nIndex) = "Next_Image" Then moveToNextChildWindow True
    If ctlAccelerator.Key(nIndex) = "Prev_Image" Then moveToNextChildWindow False
    
    lastAccelerator = Timer
    
End Sub

'All "Window" menu items are handled here
Private Sub MnuWindow_Click(Index As Integer)

    Dim i As Long
    
    Dim prevActiveWindow As Long
    prevActiveWindow = g_CurrentImage

    Select Case Index
    
        'Show/hide file toolbox
        Case 0
            toggleToolbarVisibility FILE_TOOLBOX
        
        'Show/hide layer toolbox
        Case 1
            toggleToolbarVisibility LAYER_TOOLBOX
            
        'Show/hide selection toolbox
        Case 2
            toggleToolbarVisibility TOOLS_TOOLBOX
        
        '<top-level Image tabstrip>
        Case 3
        
        '<separator>
        Case 4
    
        'Floating toolbars
        Case 5
            toggleWindowFloating TOOLBAR_WINDOW, Not FormMain.MnuWindow(5).Checked
        
        '<separator>
        Case 6
        
        'Next image
        Case 7
            moveToNextChildWindow True
            
        'Previous image
        Case 8
            moveToNextChildWindow False

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
                activatePDImage i, "user requested next/previous image"
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

'Unlike other toolbars, the image tabstrip has a more complicated window menu, because it is viewable under a variety
' of conditions, and we allow the user to specify any alignment.
Private Sub MnuWindowTabstrip_Click(Index As Integer)

    Select Case Index
    
        'Always display image tabstrip
        Case 0
            toggleImageTabstripVisibility Index
        
        'Display tabstrip for 2+ images (default)
        Case 1
            toggleImageTabstripVisibility Index
        
        'Never display image tabstrip
        Case 2
            toggleImageTabstripVisibility Index
        
        '<separator>
        Case 3
        
        'Align left
        Case 4
            toggleImageTabstripAlignment vbAlignLeft
        
        'Align top
        Case 5
            toggleImageTabstripAlignment vbAlignTop
        
        'Align right
        Case 6
            toggleImageTabstripAlignment vbAlignRight
        
        'Align bottom
        Case 7
            toggleImageTabstripAlignment vbAlignBottom
    
    End Select

End Sub

Private Sub MnuZoomIn_Click()
    If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex > 0 Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex - 1
End Sub

Private Sub MnuZoomOut_Click()
    If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex < (FormMain.mainCanvas(0).getZoomDropDownReference().ListCount - 1) Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex + 1
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

