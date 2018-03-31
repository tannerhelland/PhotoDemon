VERSION 5.00
Begin VB.Form FormMain 
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.photodemon.org"
   ClientHeight    =   11130
   ClientLeft      =   1290
   ClientTop       =   1065
   ClientWidth     =   15510
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1034
   Begin PhotoDemon.pdAccelerator pdHotkeys 
      Left            =   120
      Top             =   2280
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdCanvas MainCanvas 
      Height          =   5055
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
   End
   Begin PhotoDemon.pdDownload asyncDownloader 
      Left            =   120
      Top             =   1680
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Menu MnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu MnuFile 
         Caption         =   "&New..."
         Index           =   0
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Open &recent"
         Index           =   2
         Begin VB.Menu MnuRecDocs 
            Caption         =   "Empty"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MnuRecentSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuLoadAllMRU 
            Caption         =   "Open all recent images"
         End
         Begin VB.Menu MnuClearMRU 
            Caption         =   "Clear recent image list"
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Import"
         Index           =   3
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
         Index           =   4
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Close"
         Index           =   5
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Close all"
         Index           =   6
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Save"
         Index           =   8
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save copy (&lossless)"
         Index           =   9
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save &as..."
         Index           =   10
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Revert"
         Index           =   11
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Export"
         Index           =   12
         Begin VB.Menu MnuFileExport 
            Caption         =   "Palette..."
            Index           =   0
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Batch operations"
         Index           =   14
         Begin VB.Menu MnuBatch 
            Caption         =   "Process..."
            Index           =   0
         End
         Begin VB.Menu MnuBatch 
            Caption         =   "Repair..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Print..."
         Index           =   16
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuFile 
         Caption         =   "E&xit"
         Index           =   18
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
         Caption         =   "Undo history..."
         Index           =   2
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Repeat"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Fade..."
         Index           =   5
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Cu&t"
         Index           =   7
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Cut from layer"
         Index           =   8
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Copy"
         Index           =   9
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Copy from layer"
         Index           =   10
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Paste as new image"
         Index           =   11
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Paste as new layer"
         Index           =   12
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Empty clipboard"
         Index           =   14
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
         Caption         =   "Content-aware resize..."
         Index           =   3
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Canvas size..."
         Index           =   5
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Fit canvas to active layer"
         Index           =   6
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Fit canvas around all layers"
         Index           =   7
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Crop to selection"
         Index           =   9
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Trim empty borders"
         Index           =   10
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Rotate"
         Index           =   12
         Begin VB.Menu MnuRotate 
            Caption         =   "Straighten..."
            Index           =   0
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90 clockwise"
            Index           =   2
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90 counter-clockwise"
            Index           =   3
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "180"
            Index           =   4
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Arbitrary..."
            Index           =   5
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip horizontal"
         Index           =   13
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip vertical"
         Index           =   14
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Metadata"
         Index           =   16
         Begin VB.Menu MnuMetadata 
            Caption         =   "Edit metadata..."
            Index           =   0
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Remove all metadata"
            Index           =   1
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Count unique colors"
            Index           =   3
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Map photo location..."
            Index           =   4
         End
      End
   End
   Begin VB.Menu MnuLayerTop 
      Caption         =   "&Layer"
      Begin VB.Menu MnuLayer 
         Caption         =   "Add"
         Index           =   0
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Basic layer..."
            Index           =   0
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Blank layer"
            Index           =   1
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Duplicate of current layer"
            Index           =   2
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "From clipboard"
            Index           =   4
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "From file..."
            Index           =   5
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "From visible layers"
            Index           =   6
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
            Caption         =   "Straighten..."
            Index           =   0
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 90 clockwise"
            Index           =   2
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 90 counter-clockwise"
            Index           =   3
         End
         Begin VB.Menu MnuLayerOrientation 
            Caption         =   "Rotate 180"
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
         Caption         =   "Crop to selection"
         Index           =   9
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Transparency"
         Index           =   11
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Make color transparent..."
            Index           =   0
         End
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Remove transparency..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Rasterize"
         Index           =   13
         Begin VB.Menu MnuLayerRasterize 
            Caption         =   "Current layer"
            Index           =   0
         End
         Begin VB.Menu MnuLayerRasterize 
            Caption         =   "All layers"
            Index           =   1
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge visible layers"
         Index           =   15
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Flatten image..."
         Index           =   16
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
         Caption         =   "Erase selected area"
         Index           =   10
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Load selection..."
         Index           =   12
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Save current selection..."
         Index           =   13
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Export"
         Index           =   14
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
         Caption         =   "Auto correct"
         Index           =   0
         Begin VB.Menu MnuAutoCorrect 
            Caption         =   "Color"
            Index           =   0
         End
         Begin VB.Menu MnuAutoCorrect 
            Caption         =   "Contrast"
            Index           =   1
         End
         Begin VB.Menu MnuAutoCorrect 
            Caption         =   "Lighting"
            Index           =   2
         End
         Begin VB.Menu MnuAutoCorrect 
            Caption         =   "Shadows and highlights"
            Index           =   3
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Auto enhance"
         Index           =   1
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Color"
            Index           =   0
         End
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Contrast"
            Index           =   1
         End
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Lighting"
            Index           =   2
         End
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Shadows and highlights"
            Index           =   3
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Black and white..."
         Index           =   3
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Brightness and contrast..."
         Index           =   4
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Color balance..."
         Index           =   5
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Curves..."
         Index           =   6
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Levels..."
         Index           =   7
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Shadows and highlights..."
         Index           =   8
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Vibrance..."
         Index           =   9
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "White balance..."
         Index           =   10
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Channels"
         Index           =   12
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
            Caption         =   "Maximum"
            Index           =   3
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Minimum"
            Index           =   4
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift left"
            Index           =   6
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift right"
            Index           =   7
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Color"
         Index           =   13
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
            Caption         =   "Temperature..."
            Index           =   4
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Tint..."
            Index           =   5
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Vibrance..."
            Index           =   6
         End
         Begin VB.Menu MnuColor 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Black and white..."
            Index           =   8
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Colorize..."
            Index           =   9
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Replace color..."
            Index           =   10
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Sepia"
            Index           =   11
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Histogram"
         Index           =   14
         Begin VB.Menu MnuHistogram 
            Caption         =   "Display..."
            Index           =   0
         End
         Begin VB.Menu MnuHistogram 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuHistogram 
            Caption         =   "Equalize..."
            Index           =   2
         End
         Begin VB.Menu MnuHistogram 
            Caption         =   "Stretch"
            Index           =   3
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Invert"
         Index           =   15
         Begin VB.Menu MnuInvert 
            Caption         =   "CMYK (film negative)"
            Index           =   0
         End
         Begin VB.Menu MnuInvert 
            Caption         =   "Hue"
            Index           =   1
         End
         Begin VB.Menu MnuInvert 
            Caption         =   "RGB"
            Index           =   2
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Lighting"
         Index           =   16
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
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Monochrome"
         Index           =   17
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Color to monochrome..."
            Index           =   0
         End
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Monochrome to gray..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Photography"
         Index           =   18
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Exposure..."
            Index           =   0
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "HDR..."
            Index           =   1
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Photo filters..."
            Index           =   2
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Red-eye removal..."
            Index           =   3
         End
         Begin VB.Menu MnuAdjustmentsPhoto 
            Caption         =   "Split toning..."
            Index           =   4
         End
      End
   End
   Begin VB.Menu MnuEffectsTop 
      Caption         =   "Effe&cts"
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Artistic"
         Index           =   0
         Begin VB.Menu MnuArtistic 
            Caption         =   "Colored pencil..."
            Index           =   0
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Comic book..."
            Index           =   1
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Figured glass (dents)..."
            Index           =   2
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Film noir..."
            Index           =   3
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Glass tiles..."
            Index           =   4
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Kaleiodoscope..."
            Index           =   5
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Modern art..."
            Index           =   6
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Oil painting..."
            Index           =   7
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Plastic wrap..."
            Index           =   8
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Posterize..."
            Index           =   9
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Relief..."
            Index           =   10
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Stained glass..."
            Index           =   11
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
            Caption         =   "Kuwahara filter..."
            Index           =   8
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Symmetric nearest-neighbor..."
            Index           =   9
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Distort"
         Index           =   2
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Correct existing distortion..."
            Index           =   0
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Donut..."
            Index           =   2
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Lens..."
            Index           =   3
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Pinch and whirl..."
            Index           =   4
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Poke..."
            Index           =   5
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Ripple..."
            Index           =   6
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Squish..."
            Index           =   7
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Swirl..."
            Index           =   8
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Waves..."
            Index           =   9
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu MnuDistortEffects 
            Caption         =   "Miscellaneous..."
            Index           =   11
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Edge"
         Index           =   3
         Begin VB.Menu MnuEdge 
            Caption         =   "Emboss..."
            Index           =   0
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Enhance edges..."
            Index           =   1
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Find edges..."
            Index           =   2
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Range filter..."
            Index           =   3
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Trace contour..."
            Index           =   4
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Light and shadow"
         Index           =   4
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Black light..."
            Index           =   0
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Cross-screen..."
            Index           =   1
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Rainbow..."
            Index           =   2
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Sunshine..."
            Index           =   3
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Dilate..."
            Index           =   5
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Erode..."
            Index           =   6
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Natural"
         Index           =   5
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Atmosphere..."
            Index           =   0
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Fog..."
            Index           =   1
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Ignite..."
            Index           =   2
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Lava..."
            Index           =   3
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Metal..."
            Index           =   4
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Snow..."
            Index           =   5
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Underwater..."
            Index           =   6
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
            Caption         =   "Anisotropic diffusion..."
            Index           =   3
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Bilateral filter..."
            Index           =   4
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Harmonic mean..."
            Index           =   5
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Mean shift..."
            Index           =   6
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Median..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Pixelate"
         Index           =   7
         Begin VB.Menu MnuPixelate 
            Caption         =   "Color halftone..."
            Index           =   0
         End
         Begin VB.Menu MnuPixelate 
            Caption         =   "Crystallize..."
            Index           =   1
         End
         Begin VB.Menu MnuPixelate 
            Caption         =   "Fragment..."
            Index           =   2
         End
         Begin VB.Menu MnuPixelate 
            Caption         =   "Mezzotint..."
            Index           =   3
         End
         Begin VB.Menu MnuPixelate 
            Caption         =   "Mosaic..."
            Index           =   4
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
            Caption         =   "Antique..."
            Index           =   0
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Diffuse..."
            Index           =   1
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Outline..."
            Index           =   2
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Palettize..."
            Index           =   3
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Portrait glow..."
            Index           =   4
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Solarize..."
            Index           =   5
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Twins..."
            Index           =   6
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Vignetting..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Transform"
         Index           =   10
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Pan and zoom..."
            Index           =   0
         End
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Perspective..."
            Index           =   1
         End
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Polar conversion..."
            Index           =   2
         End
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Rotate..."
            Index           =   3
         End
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Shear..."
            Index           =   4
         End
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Spherize..."
            Index           =   5
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuCustomFilter 
         Caption         =   "Custom filter..."
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MnuTool 
         Caption         =   "Language"
         Index           =   0
         Begin VB.Menu MnuLanguages 
            Caption         =   "English (US)"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Language editor..."
         Index           =   1
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Theme..."
         Index           =   3
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Record macro"
         Index           =   5
         Begin VB.Menu MnuRecordMacro 
            Caption         =   "Start recording"
            Index           =   0
         End
         Begin VB.Menu MnuRecordMacro 
            Caption         =   "Stop recording..."
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Play macro..."
         Index           =   6
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Recent macros"
         Index           =   7
         Begin VB.Menu MnuRecentMacros 
            Caption         =   "Empty"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MnuRecentMacroSepBar 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MnuClearRecentMacros 
            Caption         =   "Clear recent macro list"
         End
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Options..."
         Index           =   9
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Plugin manager..."
         Index           =   10
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Developers"
         Index           =   12
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Theme editor..."
            Index           =   0
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Build theme package..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuWindowTop 
      Caption         =   "&Window"
      Begin VB.Menu MnuWindow 
         Caption         =   "Toolbox"
         Index           =   0
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Display toolbox"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Display tool category titles"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Small buttons"
            Index           =   4
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Normal buttons"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Large buttons"
            Index           =   6
         End
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tool options"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Layers"
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
         Caption         =   "Reset all toolboxes"
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
         Caption         =   "Support us with a small donation (thank you!)"
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
         Caption         =   "&Visit PhotoDemon website"
         Index           =   6
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Download PhotoDemon source code"
         Index           =   7
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Read license and terms of use"
         Index           =   8
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&About"
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

'PhotoDemon is Copyright 1999-2018 by Tanner Helland, tannerhelland.com

'Please visit photodemon.org for updates and additional downloads

'***************************************************************************
'Primary PhotoDemon Window
'Copyright 2002-2018 by Tanner Helland
'Created: 15/September/02
'Last updated: 27/March/18
'Last update: new export menu items added
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its primary purpose is sending
' parameters to other, more interesting sections of the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This main dialog houses a few timer objects; these can be started and/or stopped by external functions.  See the timer
' start/stop functions for additional details.
Private WithEvents m_InterfaceTimer As pdTimer
Attribute m_InterfaceTimer.VB_VarHelpID = -1
Private WithEvents m_MetadataTimer As pdTimer
Attribute m_MetadataTimer.VB_VarHelpID = -1

Private m_AllowedToReflowInterface As Boolean

Private Sub MnuTest_Click()
    
    Filters_Sci.InternalFFTTest
    
    'Want to test a new dialog?  Call it here, using a line like the following:
    'showPDDialog vbModal, FormToTest
    
End Sub

'Whenever the asynchronous downloader completes its work, we forcibly release all resources associated with downloads we've finished processing.
Private Sub asyncDownloader_FinishedAllItems(ByVal allDownloadsSuccessful As Boolean)
    
    'Core program updates are handled specially, so their resources can be freed without question.
    asyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK"
    asyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK_USER"
    
    FormMain.MainCanvas(0).SetNetworkState False
    Debug.Print "All downloads complete."
    
End Sub

'When an asynchronous download completes, deal with it here
Private Sub asyncDownloader_FinishedOneItem(ByVal downloadSuccessful As Boolean, ByVal entryKey As String, ByVal OptionalType As Long, downloadedData() As Byte, ByVal savedToThisFile As String)
    
    'On a typical PD install, updates are checked every session, but users can specify a larger interval in the preferences dialog.
    ' As part of honoring that preference, whenever an update check successfully completes, we write the current date out to the
    ' preferences file, so subsequent runs can limit their check frequency accordingly.
    If Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK", True) Or Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK_USER", True) Then
        
        If downloadSuccessful Then
        
            'The update file downloaded correctly.  Write today's date to the master preferences file, so we can correctly calculate
            ' weekly/monthly update checks for users that require it.
            Debug.Print "Update file download complete.  Update information has been saved at " & savedToThisFile
            UserPrefs.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
            
            'Retrieve the file contents into a string
            Dim updateXML As String
            updateXML = StrConv(downloadedData, vbUnicode)
            
            'Offload the rest of the check to a separate update function.  It will initiate subsequent downloads as necessary.
            Dim updateAvailable As Boolean
            updateAvailable = Updates.ProcessProgramUpdateFile(updateXML)
            
            'If the user initiated the download, display a modal notification now
            If (StrComp(entryKey, "PROGRAM_UPDATE_CHECK_USER") = 0) Then
                
                If updateAvailable Then
                    Message "A new version of PhotoDemon is available.  The update is automatically processing in the background..."
                Else
                    Message "This copy of PhotoDemon is up to date."
                End If
                
                'Perform a low-risk yield to events, so the status bar message has time to repaint itself before the message box appears
                DoEvents
                
                If updateAvailable Then
                    PDMsgBox "A new version of PhotoDemon is available!" & vbCrLf & vbCrLf & "The update is automatically processing in the background.  You will receive a new notification when it completes.", vbOKOnly Or vbInformation, "PhotoDemon Updates", App.Major, App.Minor, App.Revision
                Else
                    PDMsgBox "This copy of PhotoDemon is the newest version available." & vbCrLf & vbCrLf & "(Current version: %1.%2.%3)", vbOKOnly Or vbInformation, "PhotoDemon Updates", App.Major, App.Minor, App.Revision
                End If
                
                'If the update managed to download while the reader was staring at the message box, display the restart notification immediately
                If g_ShowUpdateNotification Then Updates.DisplayUpdateNotification
                
            End If
            
        Else
            Debug.Print "Update file was not downloaded.  asyncDownloader returned this error message: " & asyncDownloader.GetLastErrorNumber & " - " & asyncDownloader.GetLastErrorDescription
        End If
    
    'If PROGRAM_UPDATE_CHECK (above) finds updated program or plugin files, it will trigger their download.  When the download arrives,
    ' we can start patching immediately.
    ElseIf (OptionalType = PD_PATCH_IDENTIFIER) Then
        
        If downloadSuccessful Then
            
            'Notify the software updater that an update package was downloaded successfully.  It will make a note of this, so it can
            ' complete the actual patching when PD closes.
            Updates.NotifyUpdatePackageAvailable savedToThisFile
            
            'Display a notification to the user
            Updates.DisplayUpdateNotification
                        
        Else
            Debug.Print "WARNING!  A program update was found, but the download was interrupted.  PD is postponing further patches until a later session."
        End If
        
    End If

End Sub

'External functions can request asynchronous downloads via this function.
Public Function RequestAsynchronousDownload(ByRef downloadKey As String, ByRef urlString As String, Optional ByVal OptionalDownloadType As Long = 0, Optional ByVal asyncFlags As AsyncReadConstants = vbAsyncReadResynchronize, Optional ByVal saveToThisFileWhenComplete As String = vbNullString, Optional ByVal checksumToVerify As Long = 0) As Boolean
    FormMain.MainCanvas(0).SetNetworkState True
    RequestAsynchronousDownload = Me.asyncDownloader.AddToQueue(downloadKey, urlString, OptionalDownloadType, asyncFlags, True, saveToThisFileWhenComplete, checksumToVerify)
End Function

'When the main form is resized, we must re-align the main canvas to match
Private Sub Form_Resize()
    If (Not g_WindowManager Is Nothing) Then
        If g_WindowManager.GetAutoRefreshMode Then UpdateMainLayout
    Else
        UpdateMainLayout
    End If
End Sub

Public Function ToolbarsAllowedToReflow() As Boolean
    ToolbarsAllowedToReflow = m_AllowedToReflowInterface
End Function

'Resize all currently active canvases.  This was an important function back when PD used an MDI engine, but now that we
' use our own tabbed interface, it's due for a major revisit.  If we could kill this function entirely, I'd be very happy.
Public Sub UpdateMainLayout(Optional ByVal resizeToolboxesToo As Boolean = True)

    'If the main form has been minimized, don't refresh anything
    If (FormMain.WindowState = vbMinimized) Then Exit Sub
    If (Not m_AllowedToReflowInterface) Then Exit Sub
    
    'As of 7.0, a new, lightweight toolbox manager can calculate idealized window positions for us.
    Dim mainRect As winRect, canvasRect As winRect
    g_WindowManager.GetClientWinRect FormMain.hWnd, mainRect
    Toolboxes.CalculateNewToolboxRects mainRect, canvasRect
    
    'With toolbox positions successfully calculated, we can now synchronize each toolbox to its calculated rect.
    If resizeToolboxesToo Then
        Toolboxes.PositionToolbox PDT_LeftToolbox, toolbar_Toolbox.hWnd, FormMain.hWnd
        Toolboxes.PositionToolbox PDT_RightToolbox, toolbar_Layers.hWnd, FormMain.hWnd
        Toolboxes.PositionToolbox PDT_BottomToolbox, toolbar_Options.hWnd, FormMain.hWnd
    End If
    
    'Similarly, we can drop the canvas into place using the helpful rect provided by the toolbox module.
    ' Note that resizing the canvas rect will automatically trigger a redraw of the viewport, as necessary.
    With canvasRect
        FormMain.MainCanvas(0).SetPositionAndSize .x1, .y1, .x2 - .x1, .y2 - .y1
    End With
    
    'If all three toolboxes are hidden, Windows may try to hide the main window as well.  Reset focus manually.
    If Toolboxes.AreAllToolboxesHidden Then g_WindowManager.SetFocusAPI FormMain.hWnd
    
End Sub

'Some functions need to artificially delay handling user input to prevent "click-through".  Use this function to do so.
Public Sub StartInterfaceTimer()

    If (m_InterfaceTimer Is Nothing) Then
        Set m_InterfaceTimer = New pdTimer
        m_InterfaceTimer.Interval = 50
    End If
    
    m_InterfaceTimer.StartTimer
    
End Sub

'Countdown timer for re-enabling disabled user input.  A delay is enforced to prevent double-clicks on child dialogs from
' "passing through" to the main form and causing goofy behavior.
Private Sub m_InterfaceTimer_Timer()

    Static intervalCount As Long
    
    If (intervalCount >= 1) Then
        intervalCount = 0
        g_DisableUserInput = False
        m_InterfaceTimer.StopTimer
    End If
    
    intervalCount = intervalCount + 1

End Sub

Public Sub StartMetadataTimer()
    
    If (m_MetadataTimer Is Nothing) Then
        Set m_MetadataTimer = New pdTimer
        m_MetadataTimer.Interval = 250
    End If
    
    If (Not m_MetadataTimer.IsActive) Then m_MetadataTimer.StartTimer
    
End Sub

'This metadata timer is a final failsafe for images with huge metadata collections that take a long time to parse.  If an image has successfully
' loaded but its metadata parsing is still in-progress, PD's load function will activate this timer.  The timer will wait (asynchronously) for
' metadata parsing to finish, and when it does, it will copy the metadata into the active pdImage object, then disable itself.
Private Sub m_MetadataTimer_Timer()

    'I don't like resorting to hackneyed error-handling, but ExifTool can be unpredictable, especially if the user loads a bajillion
    ' images simultaneously.  Rather than bring down the whole program, I'd prefer to simply ignore metadata for the problematic image.
    On Error Resume Next

    If ExifTool.IsMetadataFinished Then
    
        'Start by disabling this timer (as it's no longer needed)
        m_MetadataTimer.StopTimer
        
        'Cache the current UI message (if any)
        Dim prevMessage As String
        prevMessage = Interface.GetLastFullMessage()
                
        Message "Asynchronous metadata check complete!  Updating metadata collection..."
        
        'Retrieve the completed metadata string
        Dim mdString As String, tmpString As String
        mdString = ExifTool.RetrieveMetadataString()
        
        Dim curImageID As Long
        
        'Now comes some messy string parsing.  If the user has loaded multiple images at once, the metadata string returned by ExifTool will contain
        ' ALL METADATA for ALL IMAGES in one giant string.  We need to parse out each image's metadata, supply it to the correct image, then repeat
        ' until all images have received their relevant metadata.
        
        'Start by finding the first occurrence of ExifTool's unique "{ready}" message, which signifies its success in completing a single coherent
        ' -execute request.
        Dim startPosition As Long, terminalPosition As Long
        startPosition = 1
        terminalPosition = InStr(1, mdString, "{ready", vbBinaryCompare)
        
        Do While (terminalPosition <> 0)
        
            'terminalPosition now contains the position of ExifTool's "{ready123}" tag, where 123 is the ID of the image whose metadata
            ' is contained prior to that point.  Start by figuring out what that ID number actually is.
            Dim lenFailsafe As Long
            
            If (terminalPosition + 6 < Len(mdString)) Then
                lenFailsafe = InStr(terminalPosition + 6, mdString, "}", vbBinaryCompare) - (terminalPosition + 6)
            Else
                lenFailsafe = 0
            End If
            
            If (lenFailsafe <> 0) Then
                
                'Attempt to retrieve the relevant image ID for this section of metadata
                If (terminalPosition + 6 + lenFailsafe) < Len(mdString) Then
                
                    tmpString = Mid$(mdString, terminalPosition + 6, lenFailsafe)
                    
                    If IsNumeric(tmpString) Then
                        curImageID = CLng(tmpString)
                    'Else
                        'Debug.Print "Metadata ID calculation invalid - was ExifTool updated? - " & tmpString
                    End If
                    
                    'Now we know where the metadata for this image *ends*, but we still need to find where it *starts*.  All metadata XML entries start with
                    ' a standard XML header.  Search backwards from the {ready123} message until such a header is found.
                    startPosition = InStrRev(mdString, "<?xml", terminalPosition, vbBinaryCompare)
                    
                    'Using the start and final markers, extract the relevant metadata and forward it to the relevant pdImage object
                    If (startPosition > 0) And ((terminalPosition - startPosition) > 0) Then
                        
                        'Make sure we calculated our curImageID value correctly
                        If (curImageID >= 0) And (curImageID <= UBound(pdImages)) Then
                            If (Not pdImages(curImageID) Is Nothing) Then
                            
                                'Create the imgMetadata object as necessary, and load the selected metadata into it!
                                If (pdImages(curImageID).ImgMetadata Is Nothing) Then Set pdImages(curImageID).ImgMetadata = New pdMetadata
                                pdImages(curImageID).ImgMetadata.LoadAllMetadata Mid$(mdString, startPosition, terminalPosition - startPosition), curImageID
                                
                                'Now comes kind of a weird requirement.  Because metadata is loaded asynchronously, it may
                                ' arrive after the image import engine has already written our first Undo entry out to file
                                ' (this happens at image load-time, so we have a backup if the original file disappears).
                                '
                                'If this occurs, request a rewrite from the Undo engine, so we can make sure metadata gets
                                ' added to the Undo/Redo stack.
                                If pdImages(curImageID).UndoManager.HasFirstUndoWriteOccurred Then
                                    PDDebug.LogAction "Adding late-arrival metadata to original undo entry..."
                                    pdImages(curImageID).UndoManager.ForceLastUndoDataToIncludeEverything
                                End If
                                
                            End If
                        End If
                        
                        'Find the next chunk of image metadata, if any
                        terminalPosition = InStr(terminalPosition + 6, mdString, "{ready", vbBinaryCompare)
                        
                    Else
                        Debug.Print "(startPosition > 0) And ((terminalPosition - startPosition) > 0) failed"
                        terminalPosition = 0
                    End If
                                        
                Else
                    Debug.Print "(terminalPosition + 6 + lenFailsafe) was greater than Len(mdString)"
                    terminalPosition = 0
                End If
                
            Else
                Debug.Print "lenFailsafe = 0"
                terminalPosition = 0
            End If
        
        Loop
        
        'Update the interface to match the active image.  (This must be done if things like GPS tags were found in the metadata,
        ' because their presence affects the enabling of certain metadata-related menu entries.)
        Interface.SyncInterfaceToCurrentImage
        
        'Restore the original on-screen message and exit
        Interface.Message prevMessage
        
    End If

End Sub

'Menu: Adjustments -> Photography
Private Sub MnuAdjustmentsPhoto_Click(Index As Integer)

    Select Case Index
    
        'Exposure
        Case 0
            Process "Exposure", True
        
        'HDR
        Case 1
            Process "HDR", True
        
        'Photo filters
        Case 2
            Process "Photo filter", True
        
        'Red-eye removal
        Case 3
            Process "Red-eye removal", True
        
        'Split-toning
        Case 4
            Process "Split toning", True
    
    End Select

End Sub

Private Sub MnuAutoCorrect_Click(Index As Integer)

    Select Case Index
    
        'Color
        Case 0
            Process "Auto correct color", , , UNDO_Layer
        
        'Contrast
        Case 1
            Process "Auto correct contrast", , , UNDO_Layer
        
        'Lighting
        Case 2
            Process "Auto correct lighting", , , UNDO_Layer
            
        'Shadows and highlights
        Case 3
            Process "Auto correct shadows and highlights", , , UNDO_Layer
        
    End Select

End Sub

Private Sub MnuAutoEnhance_Click(Index As Integer)

    Select Case Index
    
        'Color
        Case 0
            Process "Auto enhance color", , , UNDO_Layer
        
        'Contrast
        Case 1
            Process "Auto enhance contrast", , , UNDO_Layer
        
        'Lighting
        Case 2
            Process "Auto enhance lighting", , , UNDO_Layer
            
        'Shadows and highlights
        Case 3
            Process "Auto enhance shadows and highlights", , , UNDO_Layer
        
    End Select
    
End Sub

Private Sub MnuBatch_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            Process "Batch wizard", True
        
        Case 1
            ShowPDDialog vbModal, FormBatchRepair
        
    End Select
    
End Sub

Private Sub mnuClearRecentMacros_Click()
    g_RecentMacros.MRU_ClearList
End Sub

'The Developer Tools menu is automatically hidden in production builds, so (obviously) do not put anything here that end-users might want access to.
Private Sub mnuDevelopers_Click(Index As Integer)
    
    Select Case Index
    
        'Theme Editor
        Case 0
            ShowPDDialog vbModal, FormThemeEditor
            
        'Build theme package
        Case 1
            g_Themer.BuildThemePackage
            
    End Select

End Sub

'Menu: effect > transform actions
Private Sub MnuEffectTransform_Click(Index As Integer)

    Select Case Index
    
        'Pan and zoom
        Case 0
            Process "Pan and zoom", True
            
        'Perspective (free)
        Case 1
            Process "Perspective", True
        
        'Polar conversion
        Case 2
            Process "Polar conversion", True
            
        'Rotate
        Case 3
            Process "Rotate", True
        
        'Shear
        Case 4
            Process "Shear", True
            
        'Spherize
        Case 5
            Process "Spherize", True
        
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
            Process "Merge layer up", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Image
        
        'Merge down
        Case 4
            Process "Merge layer down", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Image
        
        'Order (top-level)
        Case 5
        
        '<separator>
        Case 6
        
        'Orientation (top-level)
        Case 7
        
        'Size (top-level)
        Case 8
        
        'Crop to selection
        Case 9
            Process "Crop layer to selection", , , UNDO_Layer
        
        '<separator>
        Case 10
        
        'Transparency (top-level)
        Case 11
        
        '<separator>
        Case 12
        
        'Rasterize (top-level)
        Case 13
        
        '<separator>
        Case 14
        
        'Merge visible layers
        Case 15
            Process "Merge visible layers", , , UNDO_Image
        
        'Flatten layers
        Case 16
            Process "Flatten image", True
        
    End Select

End Sub

'Menu: remove layers from the image
Private Sub MnuLayerDelete_Click(Index As Integer)

    Select Case Index
    
        'Delete current layer
        Case 0
            Process "Delete layer", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Image_VectorSafe
        
        'Delete all hidden layers
        Case 1
            Process "Delete hidden layers", False, , UNDO_Image_VectorSafe
        
    End Select

End Sub

'Menu: add a layer to the image
Private Sub MnuLayerNew_Click(Index As Integer)

    Select Case Index
        
        'Basic layer
        Case 0
            Process "Add new layer", True
        
        'Blank layer
        Case 1
            Process "Add blank layer", False, BuildParamList("targetlayer", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Image_VectorSafe
        
        'Duplicate of current layer
        Case 2
            Process "Duplicate Layer", False, BuildParamList("targetlayer", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Image_VectorSafe
        
        '<separator>
        Case 3
        
        'Import from clipboard
        Case 4
            Process "Paste as new layer", False, , UNDO_Image_VectorSafe, , False
        
        'Import from file
        Case 5
            Process "New layer from file", True
            
        'Import from visible layers in current image
        Case 6
            Process "New layer from visible layers", False, , UNDO_Image_VectorSafe
    
    End Select

End Sub

'Menu: change layer order
Private Sub MnuLayerOrder_Click(Index As Integer)

    Select Case Index
    
        'Raise layer
        Case 0
            Process "Raise layer", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_ImageHeader
        
        'Lower layer
        Case 1
            Process "Lower layer", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_ImageHeader
        
        '<separator>
        Case 2
        
        'Raise to top
        Case 3
            Process "Raise layer to top", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_ImageHeader
        
        'Lower to bottom
        Case 4
            Process "Lower layer to bottom", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_ImageHeader
        
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
            Process "Rotate layer 90 clockwise", , , UNDO_Layer
        
        'Rotate 270
        Case 3
            Process "Rotate layer 90 counter-clockwise", , , UNDO_Layer
        
        'Rotate 180
        Case 4
            Process "Rotate layer 180", , , UNDO_Layer
        
        'Rotate arbitrary
        Case 5
            Process "Arbitrary layer rotation", True
        
        '<separator>
        Case 6
        
        'Flip horizontal
        Case 7
            Process "Flip layer horizontally", , , UNDO_Layer
        
        'Flip vertical
        Case 8
            Process "Flip layer vertically", , , UNDO_Layer
    
    End Select

End Sub

Private Sub MnuLayerRasterize_Click(Index As Integer)
    
    Select Case Index
    
        'Current layer
        Case 0
            Process "Rasterize layer", , , UNDO_Layer
            
        'All layers
        Case 1
            Process "Rasterize all layers", , , UNDO_Image
            
    End Select
    
End Sub

Private Sub MnuLayerSize_Click(Index As Integer)

    Select Case Index
    
        'Reset to actual size
        Case 0
            Process "Reset layer size", False, BuildParamList("layerindex", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_LayerHeader
        
        '<separator>
        Case 1
            
        'Standard resize
        Case 2
            Process "Resize layer", True
        
        'Content-aware resize
        Case 3
            Process "Content-aware layer resize", True
    
    End Select

End Sub

'Light and shadow effect menu
Private Sub MnuLightShadow_Click(Index As Integer)

    Select Case Index
    
        'Black light
        Case 0
            Process "Black light", True
        
        'Cross-screen (stars)
        Case 1
            Process "Cross-screen", True
        
        'Rainbow
        Case 2
            Process "Rainbow", True
            
        'Sunshine
        Case 3
            Process "Sunshine", True
        
        '<separator>
        Case 4
        
        'Dilate (maximum rank)
        Case 5
            Process "Dilate (maximum rank)", True
        
        'Erode (minimum rank)
        Case 6
            Process "Erode (minimum rank)", True
        
    
    End Select

End Sub

Private Sub MnuPixelate_Click(Index As Integer)
    
    Select Case Index
    
        'Color halftone
        Case 0
            Process "Color halftone", True
            
        'Crystallize
        Case 1
            Process "Crystallize", True
        
        'Fragment
        Case 2
            Process "Fragment", True
            
        'Mezzotint
        Case 3
            Process "Mezzotint", True
            
        'Mosaic (pixelate)
        Case 4
            Process "Mosaic", True
        
    End Select
    
End Sub

Private Sub mnuRecentMacros_Click(Index As Integer)
    
    'Load the MRU Macro path that correlates to this index.  (If one is not found, a null string is returned)
    Dim tmpString As String
    tmpString = g_RecentMacros.GetSpecificMRU(Index)
    
    'Check - just in case - to make sure the path isn't empty
    If (LenB(tmpString) <> 0) Then Macros.PlayMacroFromFile tmpString
    
End Sub

Private Sub MnuRecordMacro_Click(Index As Integer)
    
    Select Case Index
    
        'Start recording
        Case 0
            Process "Start macro recording", , , UNDO_Nothing
        
        'Stop recording
        Case 1
            Process "Stop macro recording", True
        
    End Select
    
End Sub

Private Sub MnuWindowToolbox_Click(Index As Integer)
    
    Select Case Index
    
        'Toggle toolbox visibility
        Case 0
            ToggleToolboxVisibility PDT_LeftToolbox
        
        '<separator>
        Case 1
        
        'Toggle category labels
        Case 2
            toolbar_Toolbox.ToggleToolCategoryLabels
        
        '<separator>
        Case 3
        
        'Changes to button size (small, normal, large)
        Case 4, 5, 6
            toolbar_Toolbox.UpdateButtonSize Index - 4
            
    End Select
    
End Sub

Private Sub pdHotkeys_Accelerator(ByVal acceleratorIndex As Long)
    
    'Accelerators are divided into three groups, and they are processed in the following order:
    ' 1) Direct processor strings.  These are automatically submitted to the software processor.
    ' 2) Non-processor directives that can be fired if no images are present (e.g. Open, Paste)
    ' 3) Non-processor directives that require an image.

    '***********************************************************
    'Accelerators that are direct processor strings are handled automatically
    
    With pdHotkeys
    
        If .IsProcessorString(acceleratorIndex) Then
            
            'If the action requires an open image, check for that first
            If .IsImageRequired(acceleratorIndex) Then
                If (g_OpenImageCount = 0) Then Exit Sub
                If Not (FormLanguageEditor Is Nothing) Then
                    If FormLanguageEditor.Visible Then Exit Sub
                End If
            End If
            
            'If this action is associated with a menu, make sure that corresponding menu is enabled
            If (.HasMenu(acceleratorIndex)) Then
                If (Not Menus.IsMenuEnabled(.GetMenuName(acceleratorIndex))) Then Exit Sub
            End If
            
            Process .HotKeyName(acceleratorIndex), .IsDialogDisplayed(acceleratorIndex), , .ProcUndoValue(acceleratorIndex)
            Exit Sub
            
        End If
    
        '***********************************************************
        'This block of code holds:
        ' - Accelerators that DO NOT require at least one loaded image
        Dim keyName As String
        keyName = .HotKeyName(acceleratorIndex)
        
        'Tool selection
        If Strings.StringsEqual(keyName, "tool_activate_hand", True) Then
            toolbar_Toolbox.SelectNewTool NAV_DRAG
        ElseIf Strings.StringsEqual(keyName, "tool_activate_move", True) Then
            toolbar_Toolbox.SelectNewTool NAV_MOVE
        ElseIf Strings.StringsEqual(keyName, "tool_activate_colorpicker", True) Then
            toolbar_Toolbox.SelectNewTool COLOR_PICKER
        ElseIf Strings.StringsEqual(keyName, "tool_activate_selectrect", True) Then
            If (g_CurrentTool = SELECT_RECT) Then toolbar_Toolbox.SelectNewTool SELECT_CIRC Else toolbar_Toolbox.SelectNewTool SELECT_RECT
        ElseIf Strings.StringsEqual(keyName, "tool_activate_selectlasso", True) Then
            If (g_CurrentTool = SELECT_LASSO) Then toolbar_Toolbox.SelectNewTool SELECT_POLYGON Else toolbar_Toolbox.SelectNewTool SELECT_LASSO
        ElseIf Strings.StringsEqual(keyName, "tool_activate_selectwand", True) Then
            toolbar_Toolbox.SelectNewTool SELECT_WAND
        ElseIf Strings.StringsEqual(keyName, "tool_activate_text", True) Then
            If (g_CurrentTool = VECTOR_TEXT) Then toolbar_Toolbox.SelectNewTool VECTOR_FANCYTEXT Else toolbar_Toolbox.SelectNewTool VECTOR_TEXT
        ElseIf Strings.StringsEqual(keyName, "tool_activate_pencil", True) Then
            toolbar_Toolbox.SelectNewTool PAINT_BASICBRUSH
        ElseIf Strings.StringsEqual(keyName, "tool_activate_brush", True) Then
            toolbar_Toolbox.SelectNewTool PAINT_SOFTBRUSH
        ElseIf Strings.StringsEqual(keyName, "tool_activate_eraser", True) Then
            toolbar_Toolbox.SelectNewTool PAINT_ERASER
        ElseIf Strings.StringsEqual(keyName, "tool_activate_fill", True) Then
            toolbar_Toolbox.SelectNewTool PAINT_FILL
            
        'Menus
        ElseIf Strings.StringsEqual(keyName, "Preferences", True) Then
            If Not FormOptions.Visible Then
                ShowPDDialog vbModal, FormOptions
                Exit Sub
            End If
        
        ElseIf Strings.StringsEqual(keyName, "Plugin manager", True) Then
            If Not FormPluginManager.Visible Then
                ShowPDDialog vbModal, FormPluginManager
                Exit Sub
            End If
        End If
        
        'MRU files
        Dim i As Integer
        For i = 0 To 9
            If .HotKeyName(acceleratorIndex) = ("MRU_" & i) Then
                If FormMain.MnuRecDocs.Count > i Then
                    If FormMain.MnuRecDocs(i).Enabled Then
                        Call FormMain.mnuRecDocs_Click(i)
                        Exit Sub
                    End If
                End If
            End If
        Next i
        
        '***********************************************************
        'This block of code holds:
        ' - Accelerators that DO require at least one loaded image
        
        'If no images are loaded, exit immediately
        If (g_OpenImageCount = 0) Then Exit Sub
        
        'Fit on screen
        If .HotKeyName(acceleratorIndex) = "FitOnScreen" Then FitOnScreen
        
        'Zoom in
        If .HotKeyName(acceleratorIndex) = "Zoom_In" Then
            Call MnuZoomIn_Click
        End If
        
        'Zoom out
        If .HotKeyName(acceleratorIndex) = "Zoom_Out" Then
            Call MnuZoomOut_Click
        End If
        
        'Actual size
        If .HotKeyName(acceleratorIndex) = "Actual_Size" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetZoom100Index
        End If
        
        'Various zoom values
        If .HotKeyName(acceleratorIndex) = "Zoom_161" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 2
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_81" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 4
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_41" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 8
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_21" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 10
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_12" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 14
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_14" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 16
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_18" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 19
        End If
        
        If .HotKeyName(acceleratorIndex) = "Zoom_116" Then
            If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 21
        End If
        
        'Remove selection
        If .HotKeyName(acceleratorIndex) = "Remove selection" Then
            Process "Remove selection", , , UNDO_Selection
        End If
        
        'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
        If .HotKeyName(acceleratorIndex) = "Next_Image" Then MoveToNextChildWindow True
        If .HotKeyName(acceleratorIndex) = "Prev_Image" Then MoveToNextChildWindow False
    
    End With
        
End Sub

'Note that FormMain is only loaded after pdMain.Main() has triggered.  Look there for the *true* start of the program.
Private Sub Form_Load()
    
    On Error GoTo FormMainLoadError
    
    '*************************************************************************************************************************************
    ' Start by rerouting control to "LoadTheProgram", which initializes all key PD systems
    '*************************************************************************************************************************************
    
    'The bulk of the loading code actually takes place inside the main module's ContinueLoadingProgram() function
    If pdMain.ContinueLoadingProgram() Then
    
        '*************************************************************************************************************************************
        ' Now that all program engines are initialized, we can finally display the primary window
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Registering toolbars with the window manager..."
        m_AllowedToReflowInterface = True
        
        'Now that the main form has been correctly positioned on-screen, position all toolbars and the primary canvas
        ' to match, then display the window.
        g_WindowManager.SetAutoRefreshMode True
        FormMain.UpdateMainLayout
        g_WindowManager.SetAutoRefreshMode False
        
        'DWM may cause issues inside the IDE, so forcibly refresh the main form after displaying it.
        ' (The DoEvents fixes an unpleasant flickering issue on Windows Vista/7 when the DWM isn't running full Aero.)
        FormMain.Show vbModeless
        FormMain.Refresh
        DoEvents
        
        'Visibility for the options toolbox is automatically set according to the current tool; this is different from other dialogs.
        ' (Note that the .ResetToolButtonStates function checks the relevant preference prior to changing the window state, so all
        '  cases are covered nicely.)
        toolbar_Toolbox.ResetToolButtonStates
        
        'With all toolboxes loaded, we can safely reactivate automatic syncing of toolboxes and the main window
        g_WindowManager.SetAutoRefreshMode True
        
        
        '*************************************************************************************************************************************
        ' Next, make sure PD's previous session closed down successfully
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Checking for old autosave data..."
        Autosaves.InitializeAutosave
        
        
        '*************************************************************************************************************************************
        ' Next, analyze the command line and load any image files (if present).
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Checking command line..."
        
        'Retrieve a Unicode-friendly copy of any command line parameters
        Dim cmdLineParams As pdStringStack
        If OS.CommandW(cmdLineParams, True) Then
            PDDebug.LogAction "Command line might contain images.  Attempting to load..."
            Loading.LoadMultipleImageFiles cmdLineParams, True
        End If
        
        
        '*************************************************************************************************************************************
        ' Next, see if we need to launch an asynchronous check for updates
        '*************************************************************************************************************************************
        
        'Update checks only work in portable mode (because we require write access to our own folder to do an
        ' in-place update).
        If (Not UserPrefs.IsNonPortableModeActive()) Then Updates.StandardUpdateChecks
        
        
        '*************************************************************************************************************************************
        ' Display any final messages and/or warnings
        '*************************************************************************************************************************************
        
        Message vbNullString
        FormMain.Refresh
        DoEvents
        
        'I occasionally add dire messages to nightly builds.  The line below is the best place to enable that, as necessary.
        'PDMsgBox "WARNING!  I am currently overhauling PhotoDemon's image export capabilities.  Because this work impacts the reliability of the File > Save and File > Save As commands, I DO NOT RECOMMEND using this build for serious work." & vbCrLf & vbCrLf & "(Seriously: please do any serious editing with with the last stable release, available from photodemon.org)", vbExclamation + vbOKOnly + vbApplicationModal, "7.0 Development Warning"
        
        '*************************************************************************************************************************************
        ' Next, see if we need to display the language/theme selection dialog
        '*************************************************************************************************************************************
        
        'In v7.0, a new "choose your language and UI theme" dialog was added to the project.  This is helpful for first-time
        ' users to help them get everything set up just the way they want it.
        
        'See if we've shown this dialog before; if we have, suspend its load.
        If (Not UserPrefs.GetPref_Boolean("Themes", "HasSeenThemeDialog", False)) Then DialogManager.PromptUITheme
        
        '*************************************************************************************************************************************
        ' For developers only, calculate some debug counts and show an IDE avoidance warning (if it hasn't been dismissed before).
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Current PD custom control count: " & UserControls.GetPDControlCount
        
        'Because people may be using this code in the IDE, warn them about the consequences of doing so
        If (Not OS.IsProgramCompiled) Then
            If (UserPrefs.GetPref_Boolean("Core", "Display IDE Warning", True)) Then DisplayIDEWarning
        End If
        
        'Because various user preferences may have been modified during the load process (to account for
        ' failure states, system configurations, etc), write a copy of our potentially-modified
        ' preference list out to file.
        UserPrefs.ForceWriteToFile False
        
        'In debug mode, note that we are about to turn control over to the user
        PDDebug.LogAction "Program initialization complete.  Second baseline memory measurement:"
        PDDebug.LogAction vbNullString, PDM_Mem_Report
        
        'Finally, return focus to the main form
        g_WindowManager.SetFocusAPI FormMain.hWnd
        
        Exit Sub
        
FormMainLoadError:
        PDDebug.LogAction "WARNING!  FormMain_Load experienced an error: #" & Err.Number & ", " & Err.Description
        
    'Something went catastrophically wrong during the load process.  Do not continue with the loading process.
    Else
        MsgBox "PhotoDemon has experienced a critical startup error." & vbCrLf & vbCrLf & "This can occur when the application is placed in a restricted system folder, like C:\Program Files\ or C:\Windows\.  Because PhotoDemon is a portable application, security precautions require it to operate from a non-system folder, like Desktop, Documents, or Downloads.  Please relocate the program to one of these folders, then try again." & vbCrLf & vbCrLf & "(The application will now close.)", vbOKOnly + vbCritical, "Startup failure"
        Unload Me
    End If
    
End Sub

'Allow the user to drag-and-drop files and URLs onto the main form
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If (Not g_AllowDragAndDrop) Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    Dim dropAsNewLayer As VbMsgBoxResult
    dropAsNewLayer = DialogManager.PromptForDropAsNewLayer()
    If (dropAsNewLayer <> vbCancel) Then g_Clipboard.LoadImageFromDragDrop Data, Effect, (dropAsNewLayer = vbNo)
    
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    
    'PD supports a lot of potential drop sources these days.  These values are defined and addressed by the main
    ' clipboard handler, as Drag/Drop and clipboard actions share a ton of similar code.
    If g_Clipboard.IsObjectDragDroppable(Data) And g_AllowDragAndDrop Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If

End Sub

'If the user is attempting to close the program, run some checks.  Specifically, we want to make sure all child forms have been saved.
' Note: in VB6, the order of events for program closing is MDI Parent QueryUnload, MDI children QueryUnload, MDI children Unload, MDI Unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Set a public variable to let other functions know that the user has initiated a program-wide shutdown
    g_ProgramShuttingDown = True
    
    'An external function handles unloading.  If it fails, we will also cancel our unload.
    Cancel = (Not CanvasManager.CloseAllImages())
    If Cancel Then
        g_ProgramShuttingDown = False
        If (g_OpenImageCount <> 0) Then Message vbNullString
    End If
    
End Sub

'UNLOAD EVERYTHING
Private Sub Form_Unload(Cancel As Integer)
    
    'FYI, this function includes a fair amount of debug code!
    PDDebug.LogAction "Shutdown initiated"
    
    'Store the main window's location to file now.  We will use this in the future to determine which monitor
    ' to display the splash screen on
    UserPrefs.SetPref_Long "Core", "Last Window State", Me.WindowState
    UserPrefs.SetPref_Long "Core", "Last Window Left", Me.Left / TwipsPerPixelXFix
    UserPrefs.SetPref_Long "Core", "Last Window Top", Me.Top / TwipsPerPixelYFix
    UserPrefs.SetPref_Long "Core", "Last Window Width", Me.Width / TwipsPerPixelXFix
    UserPrefs.SetPref_Long "Core", "Last Window Height", Me.Height / TwipsPerPixelYFix
    
    'Hide the main window to make it appear as if we shut down quickly
    Me.Visible = False
    Interface.ReleaseResources
    
    'Cancel any pending downloads
    PDDebug.LogAction "Checking for (and terminating) any in-progress downloads..."
    
    Me.asyncDownloader.Reset
    
    'Allow any objects on this form to save preferences and other user data
    PDDebug.LogAction "Asking all FormMain components to write out final user preference values..."
    
    FormMain.MainCanvas(0).WriteUserPreferences
    Toolboxes.SaveToolboxData
    
    'Release the clipboard manager.  If we are responsible for the current clipboard data, we must manually upload a
    ' copy of all supported formats - for this reason, this step may be a little slow.
    PDDebug.LogAction "Shutting down clipboard manager..."
    
    If (Not g_Clipboard Is Nothing) Then
    
        If (g_Clipboard.IsPDDataOnClipboard And OS.IsProgramCompiled) Then
            PDDebug.LogAction "PD's data remains on the clipboard.  Rendering any additional formats now..."
            g_Clipboard.RenderAllClipboardFormatsManually
        End If
    
        Set g_Clipboard = Nothing
        
    End If
    
    'Most core plugins are released as a final step, but ExifTool only matters when images are loaded, and we know
    ' no images are loaded by this point.  Because it takes some time to shut down, trigger it prematurely.
    If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) Then
        ExifTool.TerminateExifTool
        PDDebug.LogAction "ExifTool terminated"
    End If
    
    'Perform any printer-related cleanup
    PDDebug.LogAction "Removing printer temp files..."
    Printing.PerformPrinterCleanup
    
    'Stop tracking hotkeys
    PDDebug.LogAction "Turning off hotkey manager..."
    If (Not pdHotkeys Is Nothing) Then
        pdHotkeys.DeactivateHook True
        pdHotkeys.ReleaseResources
    End If
    
    'Release the tooltip tracker
    PDDebug.LogAction "Releasing tooltip manager..."
    UserControls.FinalTooltipUnload
    
    'Destroy all custom-created icons and cursors
    PDDebug.LogAction "Destroying custom icons and cursors..."
    IconsAndCursors.DestroyAllIcons
    IconsAndCursors.UnloadAllCursors
    
    'Destroy all paint-related resources
    PDDebug.LogAction "Destroying paint tool resources..."
    Paintbrush.FreeBrushResources
    FillTool.FreeFillResources
        
    'Save all MRU lists to the preferences file.  (I've considered doing this as files are loaded, but the only time
    ' that would be an improvement is if the program crashes, and if it does crash, the user wouldn't want to re-load
    ' the problematic image anyway.)
    PDDebug.LogAction "Saving recent file list..."
    If (Not g_RecentFiles Is Nothing) Then
        g_RecentFiles.WriteListToFile
        g_RecentMacros.MRU_SaveToFile
    End If
    
    'Release any Win7-specific features
    PDDebug.LogAction "Releasing custom Windows 7+ features..."
    OS.StopWin7PlusFeatures
    
    'Tool panels are forms that we manually embed inside other forms.  Manually unload them now.
    PDDebug.LogAction vbNullString, PDM_Mem_Report
    PDDebug.LogAction "Unloading tool panels..."
    
    'Now that toolpanels are loaded/unloaded on-demand, we don't need to manually unload them at shutdown.
    ' Instead, just unload the *active* one (which we can infer from the active tool).
    If (g_CurrentTool = NAV_MOVE) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_MoveSize.hWnd
        Unload toolpanel_MoveSize
        Set toolpanel_MoveSize = Nothing
    ElseIf (g_CurrentTool = COLOR_PICKER) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_ColorPicker.hWnd
        Unload toolpanel_ColorPicker
        Set toolpanel_ColorPicker = Nothing
    ElseIf (g_CurrentTool = SELECT_RECT) Or (g_CurrentTool = SELECT_CIRC) Or (g_CurrentTool = SELECT_LINE) Or (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_WAND) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Selections.hWnd
        Unload toolpanel_Selections
        Set toolpanel_Selections = Nothing
    ElseIf (g_CurrentTool = VECTOR_TEXT) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Text.hWnd
        Unload toolpanel_Text
        Set toolpanel_Text = Nothing
    ElseIf (g_CurrentTool = VECTOR_FANCYTEXT) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_FancyText.hWnd
        Unload toolpanel_FancyText
        Set toolpanel_FancyText = Nothing
    ElseIf (g_CurrentTool = PAINT_BASICBRUSH) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Pencil.hWnd
        Unload toolpanel_Pencil
        Set toolpanel_Pencil = Nothing
    ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Paintbrush.hWnd
        Unload toolpanel_Paintbrush
        Set toolpanel_Paintbrush = Nothing
    ElseIf (g_CurrentTool = PAINT_ERASER) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Eraser.hWnd
        Unload toolpanel_Eraser
        Set toolpanel_Eraser = Nothing
    ElseIf (g_CurrentTool = PAINT_FILL) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Fill.hWnd
        Unload toolpanel_Fill
        Set toolpanel_Fill = Nothing
    End If
    
    'With all tool panels unloaded, unload all toolboxes as well
    PDDebug.LogAction "Unloading toolboxes..."
    
    Unload toolbar_Layers
    Set toolbar_Layers = Nothing
    
    Unload toolbar_Options
    Set toolbar_Options = Nothing
    
    Unload toolbar_Toolbox
    Set toolbar_Toolbox = Nothing
    
    'Release this form from the window manager, and write out all window data to file
    PDDebug.LogAction "Shutting down window manager..."
    Interface.ReleaseFormTheming Me
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.UnregisterMainForm Me
    
    'As a final failsafe, forcibly unload any remaining forms
    PDDebug.LogAction "Forcibly unloading any remaining forms..."
    
    Dim tmpForm As Form
    For Each tmpForm In Forms

        'Note that there is no need to unload FormMain, as we're about to unload it anyway!
        If Strings.StringsNotEqual(tmpForm.Name, "FormMain", True) Then
            Debug.Print "Forcibly unloading " & tmpForm.Name
            Unload tmpForm
            Set tmpForm = Nothing
        End If

    Next tmpForm
    
    'If an update package was downloaded, this is a good time to apply it
    If Updates.IsUpdatePackageAvailable And pdMain.WasStartupSuccessful() Then
        
        If Updates.PatchProgramFiles() Then
            PDDebug.LogAction "Updates.PatchProgramFiles returned TRUE.  Program update will proceed after PD finishes unloading."
            
            'If the user wants a restart, create a restart batch file now
            'If g_UserWantsRestart Then Updates.CreateRestartBatchFile
            
        Else
            PDDebug.LogAction "WARNING!  One or more errors were encountered while applying an update.  PD has attempted to roll everything back to its original state."
        End If
        
    End If
        
    'Because PD can now auto-update between runs, it's helpful to log the current program version to the preferences file.  The next time PD runs,
    ' it can compare its version against this value, to infer if an update occurred.
    PDDebug.LogAction "Writing session data to file..."
    UserPrefs.SetPref_String "Core", "LastRunVersion", App.Major & "." & App.Minor & "." & App.Revision
    
    'All core PD functions appear to have terminated correctly, so notify the Autosave handler that this session was clean.
    PDDebug.LogAction "Final step: writing out new autosave checksum..."
    Autosaves.PurgeOldAutosaveData
    Autosaves.NotifyCleanShutdown
    
    PDDebug.LogAction "Shutdown appears to be clean.  Turning final control over to pdMain.FinalShutdown()..."
    pdMain.FinalShutdown
    
    'If a restart is allowed, the last thing we do before exiting is shell a new PhotoDemon instance
    'If g_UserWantsRestart Then Updates.InitiateRestart
    
End Sub

'The top-level adjustments menu provides some shortcuts to most-used items.
Private Sub MnuAdjustments_Click(Index As Integer)

    Select Case Index
    
        'Auto correct (top-level)
        Case 0
            
        'Auto enhance (top-level)
        Case 1
        
        '<separator>
        Case 2
            
        'Black and white
        Case 3
            Process "Black and white", True
        
        'Brightness and contrast
        Case 4
            Process "Brightness and contrast", True
        
        'Color balance
        Case 5
            Process "Color balance", True
        
        'Curves
        Case 6
            Process "Curves", True
        
        'Levels
        Case 7
            Process "Levels", True
        
        'Shadows and highlights
        Case 8
            Process "Shadow and highlight", True
        
        'Vibrance
        Case 9
            Process "Vibrance", True
        
        'White balance
        Case 10
            Process "White balance", True
    
    End Select

End Sub

Private Sub MnuArtistic_Click(Index As Integer)

    Select Case Index
        
        Case 0
            Process "Colored pencil", True
        
        Case 1
            Process "Comic book", True
            
        Case 2
            Process "Figured glass", True
            
        Case 3
            Process "Film noir", True
        
        Case 4
            Process "Glass tiles", True
        
        Case 5
            Process "Kaleidoscope", True
        
        Case 6
            Process "Modern art", True
        
        Case 7
            Process "Oil painting", True
            
        Case 8
            Process "Plastic wrap", True
            
        Case 9
            Process "Posterize", True
            
        Case 10
            Process "Relief", True
            
        Case 11
            Process "Stained glass", True
    
    End Select

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
        
        'Kuwahara
        Case 8
            Process "Kuwahara filter", True
            
        'Symmetric nearest-neighbor
        Case 9
            Process "Symmetric nearest-neighbor", True
            
        'Currently unused:
        
        'Grid blur
        'Case X
        '    Process "Grid blur", , , UNDO_LAYER
            
    End Select

End Sub

Private Sub MnuClearMRU_Click()
    g_RecentFiles.ClearList
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
            
        'Temperature
        Case 4
            Process "Temperature", True
            
        'Tint
        Case 5
            Process "Tint", True
        
        'Vibrance
        Case 6
            Process "Vibrance", True
        
        '<separator>
        Case 7
        
        'Grayscale (black and white)
        Case 8
            Process "Black and white", True
        
        'Colorize
        Case 9
            Process "Colorize", True
            
        'Replace color
        Case 10
            Process "Replace color", True
                
        'Sepia
        Case 11
            Process "Sepia", , , UNDO_Layer

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
            Process "Maximum channel", , , UNDO_Layer
        
        'Min channel
        Case 4
            Process "Minimum channel", , , UNDO_Layer
            
        '<separator>
        Case 5
        
        'Shift colors left
        Case 6
            Process "Shift colors (left)", , , UNDO_Layer
            
        'Shift colors right
        Case 7
            Process "Shift colors (right)", , , UNDO_Layer
        
    End Select
    
End Sub

Private Sub MnuCustomFilter_Click()
    Process "Custom filter", True
End Sub

'All distortion filters happen here
Private Sub MnuDistortEffects_Click(Index As Integer)

    Select Case Index
        
        'Correct existing distortion(s)
        Case 0
            Process "Correct lens distortion", True
        
        '<separator>
        Case 1
        
        'Donut
        Case 2
            Process "Donut", True
        
        'Lens
        Case 3
            Process "Apply lens distortion", True
        
        'Pinch and whirl
        Case 4
            Process "Pinch and whirl", True
        
        'Poke
        Case 5
            Process "Poke", True
        
        'Ripple
        Case 6
            Process "Ripple", True
        
        'Squish (formerly Fixed Perspective)
        Case 7
            Process "Squish", True
        
        'Swirl
        Case 8
            Process "Swirl", True
        
        'Waves
        Case 9
            Process "Waves", True
            
        '<separator>
        Case 10
        
        'Miscellaneous
        Case 11
            Process "Miscellaneous distort", True
        
    End Select

End Sub

Private Sub MnuEdge_Click(Index As Integer)

    Select Case Index
        
        'Emboss/engrave
        Case 0
            Process "Emboss", True
         
        'Enhance edges
        Case 1
            Process "Enhance edges", True
        
        'Find edges
        Case 2
            Process "Find edges", True
        
        'Range filter
        Case 3
            Process "Range filter", True
        
        'Trace contour
        Case 4
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
        
        'Undo history
        Case 2
            Process "Undo history", True
            
        '<separator>
        Case 3
        
        'Repeat last
        Case 4
            'TODO: figure out Undo handling for "Repeat last action"
            Process "Repeat last action", False, , UNDO_Image
            
        'Fade...
        Case 5
            Process "Fade", True
        
        '<separator>
        Case 6
        
        'Cut from image
        Case 7
            Process "Cut", False, , UNDO_Image, , True
            
        'Cut from layer
        Case 8
        
            'If a selection is active, the Undo/Redo engine can simply back up the current layer contents.  If, however, no selection is active,
            ' we must delete the entire layer.  That requires a backup of the full image contents.
            If pdImages(g_CurrentImage).IsSelectionActive Then
                Process "Cut from layer", False, , UNDO_Layer, , True
            Else
                Process "Cut from layer", False, , UNDO_Image, , True
            End If
            
        'Copy from image
        Case 9
            Process "Copy", False, , UNDO_Nothing, , False
        
        'Copy from layer
        Case 10
            Process "Copy from layer", False, , UNDO_Nothing, , False
        
        'Paste as new image
        Case 11
            Process "Paste as new image", False, , UNDO_Nothing, , False
        
        'Paste as new layer
        Case 12
            Process "Paste as new layer", False, , UNDO_Image_VectorSafe, , False
        
        '<separator>
        Case 13
        
        'Empty clipboard
        Case 14
            Process "Empty clipboard", False, , UNDO_Nothing, , False
                
    
    End Select
    
End Sub

'All file menu actions are launched from here
Private Sub MnuFile_Click(Index As Integer)

    Select Case Index
    
        'New
        Case 0
            Process "New image", True
        
        'Open
        Case 1
            Process "Open", True
        
        '<Open Recent top-level>
        Case 2
        
        '<Import top-level>
        Case 3
        
        '<separator>
        Case 4
        
        'Close
        Case 5
            Process "Close", True
        
        'Close all
        Case 6
            Process "Close all", True
            
        '<separator>
        Case 7
        
        'Save
        Case 8
            Process "Save", True
            
        'Save copy (lossless)
        Case 9
            Process "Save copy", , , UNDO_Nothing
            
        'Save as
        Case 10
            Process "Save as", True
        
        'Revert
        Case 11
            'TODO: figure out correct Undo behavior for REVERT action
            Process "Revert", False, , UNDO_Nothing
        
        'Export
        Case 12
        
        '<separator>
        Case 13
        
        'Batch top-level menu
        Case 14
        
        '<separator>
        Case 15
        
        'Print
        Case 16
            Process "Print", True
            
        '<separator>
        Case 17
        
        'Exit
        Case 18
            Process "Exit program", True
        
    
    End Select
    
End Sub

Private Sub MnuFileExport_Click(Index As Integer)

    Select Case Index
    
        'Export palette
        Case 0
            Process "Export palette", True
    
    End Select

End Sub

Private Sub MnuFitOnScreen_Click()
    CanvasManager.FitOnScreen
End Sub

'All help menu entries are launched from here
Private Sub MnuHelp_Click(Index As Integer)

    Select Case Index
        
        'Donations are so very, very welcome!
        Case 0
            Web.OpenURL "http://photodemon.org/donate"
            
        'Check for updates
        Case 2
            Message "Checking for software updates..."
            
            'Initiate an asynchronous download of the standard PD update file (currently hosted @ GitHub).
            ' When the asynchronous download completes, the downloader will place the completed update file in the /Data/Updates subfolder.
            ' On exit (or subsequent program runs), PD will check for the presence of that file, then proceed accordingly.
            FormMain.RequestAsynchronousDownload "PROGRAM_UPDATE_CHECK_USER", "https://raw.githubusercontent.com/tannerhelland/PhotoDemon-Updates/master/summary/pdupdate.xml", , vbAsyncReadForceUpdate, UserPrefs.GetUpdatePath & "updates.xml"
            
        'Submit feedback
        Case 3
            Web.OpenURL "http://photodemon.org/about/contact/"
        
        'Submit bug report
        Case 4
            'GitHub requires a login for submitting Issues; check for that first
            Dim msgReturn As VbMsgBoxResult
            
            'If the user has previously been prompted about having a GitHub account, use their previous answer
            If UserPrefs.DoesValueExist("Core ", "Has GitHub Account") Then
            
                Dim hasGitHub As Boolean
                hasGitHub = UserPrefs.GetPref_Boolean("Core", "Has GitHub Account", False)
                
                If hasGitHub Then msgReturn = vbYes Else msgReturn = vbNo
            
            'If this is the first time they are submitting feedback, ask them if they have a GitHub account
            Else
            
                msgReturn = PDMsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbInformation Or vbYesNoCancel, "Thanks for fixing PhotoDemon")
                
                'If their answer was anything but "Cancel", store that answer to file
                If msgReturn = vbYes Then UserPrefs.SetPref_Boolean "Core", "Has GitHub Account", True
                If msgReturn = vbNo Then UserPrefs.SetPref_Boolean "Core", "Has GitHub Account", False
                
            End If
            
            'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the photodemon.org contact form
            If (msgReturn = vbYes) Then
                'Shell a browser window with the GitHub issue report form
                Web.OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            ElseIf msgReturn = vbNo Then
                'Shell a browser window with the photodemon.org contact form
                Web.OpenURL "http://photodemon.org/about/contact/"
            End If
            
        'PhotoDemon's homepage
        Case 6
            Web.OpenURL "http://www.photodemon.org"
            
        'Download source code
        Case 7
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon"
        
        'Read terms and license agreement
        Case 8
            Web.OpenURL "http://photodemon.org/about/license/"
            
        'Display About page
        Case 10
            ShowPDDialog vbModal, FormAbout
        
    End Select

End Sub

Private Sub MnuHistogram_Click(Index As Integer)
    
    Select Case Index
    
        'Display histogram (TODO: convert to processor?)
        Case 0
            ShowPDDialog vbModal, FormHistogram
            
        '<separator>
        Case 1
        
        'Equalize
        Case 2
            Process "Equalize", True
        
        'Stretch
        Case 3
            Process "Stretch histogram", , , UNDO_Layer
        
    End Select
    
End Sub

'All top-level Image menu actions are handled here
Private Sub MnuImage_Click(Index As Integer)

    Select Case Index
    
        'Duplicate
        Case 0
            Process "Duplicate image", , , UNDO_Nothing
        
        '<separator>
        Case 1
        
        'Resize
        Case 2
            Process "Resize image", True
            
        'Content-aware resize
        Case 3
            Process "Content-aware image resize", True
        
        '<separator>
        Case 4
        
        'Canvas resize
        Case 5
            Process "Canvas size", True
            
        'Fit canvas to active layer
        Case 6
            Process "Fit canvas to layer", False, BuildParamList("targetlayer", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_ImageHeader
        
        'Fit canvas around all layers
        Case 7
            Process "Fit canvas to all layers", False, , UNDO_ImageHeader
            
        '<separator>
        Case 8
            
        'Crop to selection
        Case 9
            Process "Crop", True
            
        'Trim empty borders
        Case 10
            Process "Trim empty borders", , , UNDO_ImageHeader
        
        '<separator>
        Case 11
        
        'Top-level Rotate
        Case 12
        
        'Flip horizontal (mirror)
        Case 13
            Process "Flip image horizontally", , , UNDO_Image
        
        'Flip vertical
        Case 14
            Process "Flip image vertically", , , UNDO_Image
        
        'NOTE: isometric view was removed in 6.4.  I may include it at a later date if there is demand.
        'Isometric view
        'Case 12
        '    Process "Isometric conversion"
            
        '<separator>
        Case 15
        
        'Metadata top-level
        Case 16
    
    End Select

End Sub

'This is the exact same thing as "Paste as New Image".  It is provided in two locations for convenience.
Private Sub MnuImportClipboard_Click()
    Process "Paste as new image", False, , UNDO_Nothing, , False
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    Process "Internet import", True
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    
    'Because loading a language can take some time, display a wait screen to discourage attempted interaction
    DisplayWaitScreen g_Language.TranslateMessage("Please wait while the new language is applied..."), Me
    
    'Remove the existing translation from any visible windows
    Message "Removing existing translation..."
    g_Language.UndoTranslations FormMain
    g_Language.UndoTranslations toolbar_Toolbox
    g_Language.UndoTranslations toolbar_Options
    g_Language.UndoTranslations toolbar_Layers
    DoEvents
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.ActivateNewLanguage Index
    g_Language.ApplyLanguage True, True
    
    Message "Language changed successfully."
    
    HideWaitScreen
    
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
            
        'Gamma correction
        Case 2
            Process "Gamma", True
            
        'Levels
        Case 3
            Process "Levels", True

        'Shadows/Midtones/Highlights
        Case 4
            Process "Shadow and highlight", True
            
    End Select

End Sub

'Load all images in the current "Recent Files" menu
Private Sub MnuLoadAllMRU_Click()
    
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    Dim i As Long
    For i = 0 To g_RecentFiles.GetNumOfItems() - 1
        listOfFiles.AddString g_RecentFiles.GetFullPath(i)
    Next i
    
    Loading.LoadMultipleImageFiles listOfFiles, True
    
End Sub

'All metadata sub-menu options are handled here
Private Sub MnuMetadata_Click(Index As Integer)

    Select Case Index
    
        'Browse metadata
        Case 0
            Process "Edit metadata", True
        
        'Remove all metadata
        Case 1
            Process "Remove all metadata", False, , UNDO_ImageHeader
        
        'Separator
        Case 2
        
        'Count colors
        Case 3
            Process "Count image colors", True
        
        'Map photo location
        Case 4
            
            If (Not pdImages(g_CurrentImage).ImgMetadata.HasGPSMetadata) Then
                PDMsgBox "This image does not contain any GPS metadata.", vbOKOnly Or vbInformation, "No GPS data found"
                Exit Sub
            End If
            
            Dim gMapsURL As String, latString As String, lonString As String
            If pdImages(g_CurrentImage).ImgMetadata.FillLatitudeLongitude(latString, lonString) Then
                
                'Build a valid Google maps URL (you can use Google to see what the various parameters mean)
                                
                'Note: I find a zoom of 18 ideal, as that is a common level for switching to an "aerial"
                ' view instead of a satellite view.  Much higher than that and you run the risk of not
                ' having data available at that high of zoom.
                gMapsURL = "https://maps.google.com/maps?f=q&z=18&t=h&q=" & latString & "%2c+" & lonString
                
                'As a convenience, request Google Maps in the current language
                If g_Language.TranslationActive Then
                    gMapsURL = gMapsURL & "&hl=" & g_Language.GetCurrentLanguage()
                Else
                    gMapsURL = gMapsURL & "&hl=en"
                End If
                
                'Launch Google maps in the user's browser
                Web.OpenURL gMapsURL
                
            End If
            
    End Select
    
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    
    Select Case Index
        
        'Convert color to monochrome
        Case 0
            Process "Color to monochrome", True
        
        'Convert Monochrome to gray
        Case 1
            Process "Monochrome to gray", True
        
    End Select
    
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)

    Select Case Index
    
        'Atmosphere
        Case 0
            Process "Atmosphere", True
        
        'Fog
        Case 1
            Process "Fog", True
        
        'Ignite
        Case 2
            Process "Ignite", True
        
        'Lava
        Case 3
            Process "Lava", True
        
        'Metal (formerly "steel")
        Case 4
            Process "Metal", True
        
        'Snow
        Case 5
            Process "Snow", True
            
        'Water
        Case 6
            Process "Water", True
    
    End Select

End Sub

Private Sub MnuInvert_Click(Index As Integer)
    
    Select Case Index
        
        'CMYK (film negative)
        Case 0
            Process "Film negative", , , UNDO_Layer
        
        'Hue
        Case 1
            Process "Invert hue", , , UNDO_Layer
        
        'RGB (standard)
        Case 2
            Process "Invert RGB", , , UNDO_Layer
    
    End Select
    
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
        
        'Anisotropic diffusion
        Case 3
            Process "Anisotropic diffusion", True
        
        'Bilateral smoothing
        Case 4
            Process "Bilateral smoothing", True
        
        'Harmonic mean
        Case 5
            Process "Harmonic mean", True
            
        'Mean shift
        Case 6
            Process "Mean shift", True
        
        'Median
        Case 7
            Process "Median", True
            
    End Select
        
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Load the MRU path that correlates to this index.  (If one is not found, a null string is returned)
    If (Len(g_RecentFiles.GetFullPath(Index)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(Index)
    
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
            Process "Rotate image 90 clockwise", , , UNDO_Image
        
        'Rotate 270
        Case 3
            Process "Rotate image 90 counter-clockwise", , , UNDO_Image
        
        'Rotate 180
        Case 4
            Process "Rotate image 180", , , UNDO_Image
        
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
            Process "Select all", , , UNDO_Selection
        
        'Select none
        Case 1
            Process "Remove selection", , , UNDO_Selection
        
        'Invert
        Case 2
            Process "Invert selection", , , UNDO_Selection
        
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
        
        'Erase selected area
        Case 10
            Process "Erase selected area", False, BuildParamList("targetlayer", pdImages(g_CurrentImage).GetActiveLayerIndex), UNDO_Layer
        
        '<separator>
        Case 11
        
        'Load selection
        Case 12
            Process "Load selection", True
        
        'Save current selection
        Case 13
            Process "Save selection", True
            
        '<Export top-level>
        Case 14
            
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
    If FormMain.MainCanvas(0).IsZoomEnabled Then

        Select Case Index
        
            Case 0
                FormMain.MainCanvas(0).SetZoomDropDownIndex 2
            Case 1
                FormMain.MainCanvas(0).SetZoomDropDownIndex 4
            Case 2
                FormMain.MainCanvas(0).SetZoomDropDownIndex 8
            Case 3
                FormMain.MainCanvas(0).SetZoomDropDownIndex 10
            Case 4
                FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetZoom100Index
            Case 5
                FormMain.MainCanvas(0).SetZoomDropDownIndex 14
            Case 6
                FormMain.MainCanvas(0).SetZoomDropDownIndex 16
            Case 7
                FormMain.MainCanvas(0).SetZoomDropDownIndex 19
            Case 8
                FormMain.MainCanvas(0).SetZoomDropDownIndex 21
                
        End Select

    End If

End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)

    Select Case Index
    
        'Antique
        Case 0
            Process "Antique", True
        
        'Diffuse
        Case 1
            Process "Diffuse", True
            
        'Outline
        Case 2
            Process "Outline", True
        
        'Palettize
        Case 3
            Process "Palettize", True
            
        'Portrait glow
        Case 4
            Process "Portrait glow", True
        
        'Solarize
        Case 5
            Process "Solarize", True

        'Twins
        Case 6
            Process "Twins", True
            
        'Vignetting
        Case 7
            Process "Vignetting", True
    
    End Select

End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)

    Select Case Index
        
        'Languages (top-level)
        Case 0
        
        'Language editor
        Case 1
            If (Not FormLanguageEditor.Visible) Then
                pdHotkeys.Enabled = False
                ShowPDDialog vbModal, FormLanguageEditor
                pdHotkeys.Enabled = True
            End If
            
        '(separator)
        Case 2
        
        'Theme
        Case 3
            DialogManager.PromptUITheme
        
        '(separator)
        Case 4
        
        'Record macro (top-level)
        Case 5
        
        'Play saved macro
        Case 6
            Process "Play macro", True
        
        'Recent macros (top-level)
        Case 7
        
        '(separator)
        Case 8
    
        'Options
        Case 9
            ShowPDDialog vbModal, FormOptions
            
        'Plugin manager
        Case 10
            ShowPDDialog vbModal, FormPluginManager
            
        '(separator)
        Case 11
        
        'Developer tools (top-level)
        Case 12
            
    End Select

End Sub

'Add / Remove / Modify a layer's alpha channel with this menu
Private Sub MnuLayerTransparency_Click(Index As Integer)

    Select Case Index
            
        'Color to alpha
        Case 0
            Process "Color to alpha", True
        
        'Remove alpha channel
        Case 1
            Process "Remove alpha channel", True
    
    End Select

End Sub

'All "Window" menu items are handled here
Private Sub MnuWindow_Click(Index As Integer)
    
    Select Case Index
    
        '<top-level Primary Toolbox options>
        Case 0
            
        'Show/hide tool options
        Case 1
            Toolboxes.ToggleToolboxVisibility PDT_BottomToolbox
        
        'Show/hide layer toolbox
        Case 2
            Toolboxes.ToggleToolboxVisibility PDT_RightToolbox
        
        '<top-level Image tabstrip>
        Case 3
        
        '<separator>
        Case 4
        
        'Reset all toolbox settings
        Case 5
            Toolboxes.ResetAllToolboxSettings
        
        '<separator>
        Case 6
        
        'Next image
        Case 7
            MoveToNextChildWindow True
            
        'Previous image
        Case 8
            MoveToNextChildWindow False

    End Select

End Sub

'The "Next Image" and "Previous Image" options simply wrap this function.
Private Sub MoveToNextChildWindow(ByVal moveForward As Boolean)

    'If one (or zero) images are loaded, ignore this option
    If g_OpenImageCount <= 1 Then Exit Sub
    
    Dim i As Long
    
    'Loop through all available images, and when we find one that is not this image, activate it and exit
    If moveForward Then
        i = g_CurrentImage + 1
    Else
        i = g_CurrentImage - 1
    End If
    
    Do While (i <> g_CurrentImage)
            
        'Loop back to the start of the window collection
        If moveForward Then
            If (i > g_NumOfImagesLoaded) Then i = 0
            If (i > UBound(pdImages)) Then i = 0
        Else
            If (i < 0) Then i = g_NumOfImagesLoaded
            If (i > UBound(pdImages)) Then i = UBound(pdImages)
        End If
                
        If Not (pdImages(i) Is Nothing) Then
            If pdImages(i).IsActive Then
                ActivatePDImage i, "user requested next/previous image"
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
            Interface.ToggleImageTabstripVisibility Index
        
        'Display tabstrip for 2+ images (default)
        Case 1
            Interface.ToggleImageTabstripVisibility Index
        
        'Never display image tabstrip
        Case 2
            Interface.ToggleImageTabstripVisibility Index
        
        '<separator>
        Case 3
        
        'Align left
        Case 4
            Interface.ToggleImageTabstripAlignment vbAlignLeft
        
        'Align top
        Case 5
            Interface.ToggleImageTabstripAlignment vbAlignTop
        
        'Align right
        Case 6
            Interface.ToggleImageTabstripAlignment vbAlignRight
        
        'Align bottom
        Case 7
            Interface.ToggleImageTabstripAlignment vbAlignBottom
    
    End Select

End Sub

'Zoom in/out rely on the g_Zoom object to calculate a new value
Private Sub MnuZoomIn_Click()
    If FormMain.MainCanvas(0).IsZoomEnabled Then
        If (FormMain.MainCanvas(0).GetZoomDropDownIndex > 0) Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetNearestZoomInIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
    End If
End Sub

Private Sub MnuZoomOut_Click()
    If FormMain.MainCanvas(0).IsZoomEnabled Then
        If (FormMain.MainCanvas(0).GetZoomDropDownIndex <> g_Zoom.GetZoomCount) Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetNearestZoomOutIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
    End If
End Sub

'Update the main form against the current theme.  At present, this is just a thin wrapper against the public ApplyThemeAndTranslations() function,
' but once the form's menu is owner-drawn, we will likely need some custom code to handle menu redraws and translations.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal useDoEvents As Boolean = False)
    
    'Start by notifying all controls that new translations and theme settings are required
    Interface.ApplyThemeAndTranslations Me, useDoEvents
    
    'Next, menus must be handled separately, because they are implemented using API menus
    Menus.UpdateAgainstCurrentTheme
    IconsAndCursors.ApplyAllMenuIcons
    IconsAndCursors.ResetMenuIcons
    
End Sub
