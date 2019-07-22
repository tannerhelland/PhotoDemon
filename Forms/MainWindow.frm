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
   Begin PhotoDemon.pdAccelerator HotkeyManager 
      Left            =   120
      Top             =   720
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdCanvas MainCanvas 
      Height          =   5055
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
   End
   Begin PhotoDemon.pdDownload AsyncDownloader 
      Left            =   120
      Top             =   120
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
         Begin VB.Menu MnuFileImport 
            Caption         =   "From clipboard"
            Index           =   0
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "From scanner or camera..."
            Index           =   2
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "Select which scanner or camera to use..."
            Index           =   3
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "Online image..."
            Index           =   5
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu MnuFileImport 
            Caption         =   "Screenshot..."
            Index           =   7
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
            Caption         =   "Color profile..."
            Index           =   0
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Palette..."
            Index           =   1
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
         Caption         =   "Special"
         Index           =   13
         Begin VB.Menu MnuEditSpecial 
            Caption         =   "Cut special..."
            Index           =   0
         End
         Begin VB.Menu MnuEditSpecial 
            Caption         =   "Copy special..."
            Index           =   1
         End
         Begin VB.Menu MnuEditSpecial 
            Caption         =   "Paste special..."
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Empty clipboard"
         Index           =   15
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
            Caption         =   "Go to top layer"
            Index           =   0
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Go to layer above"
            Index           =   1
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Go to layer below"
            Index           =   2
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Go to bottom layer"
            Index           =   3
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Move layer to top"
            Index           =   5
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Move layer up"
            Index           =   6
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Move layer down"
            Index           =   7
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Move layer to bottom"
            Index           =   8
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu MnuLayerOrder 
            Caption         =   "Reverse"
            Index           =   10
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
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Auto enhance"
         Index           =   1
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
            Caption         =   "Kaleidoscope..."
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
         Caption         =   "Create macro"
         Index           =   5
         Begin VB.Menu MnuMacroCreate 
            Caption         =   "From history..."
            Index           =   0
         End
         Begin VB.Menu MnuMacroCreate 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuMacroCreate 
            Caption         =   "Start recording"
            Index           =   2
         End
         Begin VB.Menu MnuMacroCreate 
            Caption         =   "Stop recording..."
            Enabled         =   0   'False
            Index           =   3
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
         Begin VB.Menu MnuDevelopers 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Build standalone package..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu MnuView 
         Caption         =   "&Fit image on screen"
         Index           =   0
      End
      Begin VB.Menu MnuView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom &in"
         Index           =   2
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom &out"
         Index           =   3
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom to value"
         Index           =   4
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
      Begin VB.Menu MnuView 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuView 
         Caption         =   "Show rulers"
         Index           =   6
      End
      Begin VB.Menu MnuView 
         Caption         =   "Show status bar"
         Index           =   7
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
         Caption         =   "Support us on Patreon..."
         Index           =   0
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Support us with a one-time donation..."
         Index           =   1
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Check for &updates..."
         Index           =   3
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit bug report or feedback..."
         Index           =   4
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&Visit PhotoDemon website..."
         Index           =   6
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Download PhotoDemon source code..."
         Index           =   7
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Read license and terms of use..."
         Index           =   8
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&About..."
         Index           =   10
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please see the included README file for additional information on licensing and redistribution.

'PhotoDemon is Copyright 1999-2019 by Tanner Helland, tannerhelland.com

'Please visit https://photodemon.org for updates and additional downloads

'***************************************************************************
'Primary PhotoDemon Window
'Copyright 2002-2019 by Tanner Helland
'Created: 15/September/02
'Last updated: 27/March/18
'Last update: new export menu items added
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its primary purpose is sending
' parameters to other, more interesting sections of the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This main dialog houses a few timer objects; these can be started and/or stopped by external functions.  See the timer
' start/stop functions for additional details.
Private WithEvents m_MetadataTimer As pdTimer
Attribute m_MetadataTimer.VB_VarHelpID = -1

'Focus detection is used to correct some hotkey behavior.  (Specifically, when this form loses focus,
' PD resets its hotkey tracker; this solves problems created by Alt+Tabbing away from the program,
' and PD thinking the Alt-key is still down when the user returns.)
Private WithEvents m_FocusDetector As pdFocusDetector
Attribute m_FocusDetector.VB_VarHelpID = -1

Private m_AllowedToReflowInterface As Boolean

Private Sub m_FocusDetector_GotFocusReliable()
    HotkeyManager.RecaptureKeyStates
End Sub

Private Sub m_FocusDetector_LostFocusReliable()
    HotkeyManager.ResetKeyStates
End Sub

Private Sub MnuFileImport_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "file_import_paste"
        Case 1
            '(separator)
        Case 2
            Menus.ProcessDefaultAction_ByName "file_import_scanner"
        Case 3
            Menus.ProcessDefaultAction_ByName "file_import_selectscanner"
        Case 4
            '(separator)
        Case 5
            Menus.ProcessDefaultAction_ByName "file_import_web"
        Case 6
            '(separator)
        Case 7
            Menus.ProcessDefaultAction_ByName "file_import_screenshot"
    End Select
End Sub

Private Sub MnuMacroCreate_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "tools_macrofromhistory"
        Case 1
            '(separator)
        Case 2
            Menus.ProcessDefaultAction_ByName "tools_recordmacro"
        Case 3
            Menus.ProcessDefaultAction_ByName "tools_stopmacro"
    End Select
End Sub

Private Sub MnuTest_Click()
    
    On Error GoTo StopTestImmediately
    
    'Filters_Scientific.InternalFFTTest
    
    'Want to test a new dialog?  Call it here, using a line like the following:
    'ShowPDDialog vbModal, FormToTest
    
    Exit Sub
    
StopTestImmediately:
    Debug.Print "Error in test sub: " & Err.Number & ", " & Err.Description

End Sub

'Whenever the asynchronous downloader completes its work, we forcibly release all resources associated with downloads we've finished processing.
Private Sub AsyncDownloader_FinishedAllItems(ByVal allDownloadsSuccessful As Boolean)
    
    'Core program updates are handled specially, so their resources can be freed without question.
    AsyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK"
    AsyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK_USER"
    
    FormMain.MainCanvas(0).SetNetworkState False
    Debug.Print "All downloads complete."
    
End Sub

'When an asynchronous download completes, deal with it here
Private Sub AsyncDownloader_FinishedOneItem(ByVal downloadSuccessful As Boolean, ByVal entryKey As String, ByVal OptionalType As Long, downloadedData() As Byte, ByVal savedToThisFile As String)
    
    'On a typical PD install, updates are checked every session, but users can specify a larger interval in the preferences dialog.
    ' As part of honoring that preference, whenever an update check successfully completes, we write the current date out to the
    ' preferences file, so subsequent runs can limit their check frequency accordingly.
    If Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK", True) Or Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK_USER", True) Then
        
        If downloadSuccessful Then
        
            'The update file downloaded correctly.  Write today's date to the master preferences file, so we can correctly calculate
            ' weekly/monthly update checks for users that require it.
            PDDebug.LogAction "Update file download complete.  Update information has been saved at " & savedToThisFile
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
            PDDebug.LogAction "Update file was not downloaded.  asyncDownloader returned this error message: " & AsyncDownloader.GetLastErrorNumber & " - " & AsyncDownloader.GetLastErrorDescription
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
            PDDebug.LogAction "WARNING!  A program update was found, but the download was interrupted.  PD is postponing further patches until a later session."
        End If
        
    End If

End Sub

'External functions can request asynchronous downloads via this function.
Public Function RequestAsynchronousDownload(ByRef downloadKey As String, ByRef urlString As String, Optional ByVal OptionalDownloadType As Long = 0, Optional ByVal asyncFlags As AsyncReadConstants = vbAsyncReadResynchronize, Optional ByVal saveToThisFileWhenComplete As String = vbNullString, Optional ByVal checksumToVerify As Long = 0) As Boolean
    FormMain.MainCanvas(0).SetNetworkState True
    RequestAsynchronousDownload = Me.AsyncDownloader.AddToQueue(downloadKey, urlString, OptionalDownloadType, asyncFlags, True, saveToThisFileWhenComplete, checksumToVerify)
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
                        If PDImages.IsImageActive(curImageID) Then
                            
                            'Create the imgMetadata object as necessary, and load the selected metadata into it!
                            If (PDImages.GetImageByID(curImageID).ImgMetadata Is Nothing) Then Set PDImages.GetImageByID(curImageID).ImgMetadata = New pdMetadata
                            PDImages.GetImageByID(curImageID).ImgMetadata.LoadAllMetadata Mid$(mdString, startPosition, terminalPosition - startPosition), curImageID
                            
                            'Now comes kind of a weird requirement.  Because metadata is loaded asynchronously, it may
                            ' arrive after the image import engine has already written our first Undo entry out to file
                            ' (this happens at image load-time, so we have a backup if the original file disappears).
                            '
                            'If this occurs, request a rewrite from the Undo engine, so we can make sure metadata gets
                            ' added to the Undo/Redo stack.
                            If PDImages.GetImageByID(curImageID).UndoManager.HasFirstUndoWriteOccurred Then
                                PDDebug.LogAction "Adding late-arrival metadata to original undo entry..."
                                PDImages.GetImageByID(curImageID).UndoManager.ForceLastUndoDataToIncludeEverything
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
        Case 0
            Menus.ProcessDefaultAction_ByName "adj_exposure"
        Case 1
            Menus.ProcessDefaultAction_ByName "adj_hdr"
        Case 2
            Menus.ProcessDefaultAction_ByName "adj_photofilters"
        Case 3
            Menus.ProcessDefaultAction_ByName "adj_redeyeremoval"
        Case 4
            Menus.ProcessDefaultAction_ByName "adj_splittone"
    End Select
End Sub

Private Sub MnuBatch_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "file_batch_process"
        Case 1
            Menus.ProcessDefaultAction_ByName "file_batch_repair"
    End Select
End Sub

Private Sub MnuClearRecentMacros_Click()
    g_RecentMacros.MRU_ClearList
End Sub

'The Developer Tools menu is automatically hidden in production builds, so (obviously) do not put anything here that end-users might want access to.
Private Sub MnuDevelopers_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuDevelopers(Index).Caption
End Sub

'Menu: effect > transform actions
Private Sub MnuEffectTransform_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "effects_panandzoom"
        Case 1
            Menus.ProcessDefaultAction_ByName "effects_perspective"
        Case 2
            Menus.ProcessDefaultAction_ByName "effects_polarconversion"
        Case 3
            Menus.ProcessDefaultAction_ByName "effects_rotate"
        Case 4
            Menus.ProcessDefaultAction_ByName "effects_shear"
        Case 5
            Menus.ProcessDefaultAction_ByName "effects_spherize"
    End Select
End Sub

'Menu: top-level layer actions
Private Sub MnuLayer_Click(Index As Integer)
    Select Case Index
        Case 0
            'Add submenu
        Case 1
            'Delete submenu
        Case 2
            '(separator)
        Case 3
            Menus.ProcessDefaultAction_ByName "layer_mergeup"
        Case 4
            Menus.ProcessDefaultAction_ByName "layer_mergedown"
        Case 5
            'Order submenu
        Case 6
            '(separator)
        Case 7
            'Orientation submenu
        Case 8
            'Size submenu
        Case 9
            Menus.ProcessDefaultAction_ByName "layer_crop"
        Case 10
            '(separator)
        Case 11
            'Transparency submenu
        Case 12
            '(separator)
        Case 13
            'Rasterize submenu
        Case 14
            '(separator)
        Case 15
            Menus.ProcessDefaultAction_ByName "layer_mergevisible"
        Case 16
            Menus.ProcessDefaultAction_ByName "layer_flatten"
    End Select
End Sub

'Menu: remove layers from the image
Private Sub MnuLayerDelete_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_deletecurrent"
        Case 1
            Menus.ProcessDefaultAction_ByName "layer_deletehidden"
    End Select
End Sub

'Menu: add a layer to the image
Private Sub MnuLayerNew_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_addbasic"
        Case 1
            Menus.ProcessDefaultAction_ByName "layer_addblank"
        Case 2
            Menus.ProcessDefaultAction_ByName "layer_duplicate"
        Case 3
            '(separator)
        Case 4
            Menus.ProcessDefaultAction_ByName "layer_addfromclipboard"
        Case 5
            Menus.ProcessDefaultAction_ByName "layer_addfromfile"
        Case 6
            Menus.ProcessDefaultAction_ByName "layer_addfromvisiblelayers"
    End Select
End Sub

'Menu: change layer order
Private Sub MnuLayerOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_gotop"
        Case 1
            Menus.ProcessDefaultAction_ByName "layer_goup"
        Case 2
            Menus.ProcessDefaultAction_ByName "layer_godown"
        Case 3
            Menus.ProcessDefaultAction_ByName "layer_gobottom"
        Case 4
            'Separator
        Case 5
            Menus.ProcessDefaultAction_ByName "layer_movetop"
        Case 6
            Menus.ProcessDefaultAction_ByName "layer_moveup"
        Case 7
            Menus.ProcessDefaultAction_ByName "layer_movedown"
        Case 8
            Menus.ProcessDefaultAction_ByName "layer_movebottom"
        Case 9
            'Separator
        Case 10
            Menus.ProcessDefaultAction_ByName "layer_reverse"
    End Select
    
End Sub

Private Sub MnuLayerOrientation_Click(Index As Integer)
    
    'Normally, we process menu commands by caption, but this menu is unique because it shares
    ' names with the Image > Orientation menu; as such, we must explicitly request actions
    Select Case Index
    
        'Straighten
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_straighten"
            
        '<separator>
        Case 1
        
        'Rotate 90
        Case 2
            Menus.ProcessDefaultAction_ByName "layer_rotate90"
        
        'Rotate 270
        Case 3
            Menus.ProcessDefaultAction_ByName "layer_rotate270"
        
        'Rotate 180
        Case 4
            Menus.ProcessDefaultAction_ByName "layer_rotate180"
        
        'Rotate arbitrary
        Case 5
            Menus.ProcessDefaultAction_ByName "layer_rotatearbitrary"
        
        '<separator>
        Case 6
        
        'Flip horizontal
        Case 7
            Menus.ProcessDefaultAction_ByName "layer_fliphorizontal"
        
        'Flip vertical
        Case 8
            Menus.ProcessDefaultAction_ByName "layer_flipvertical"
    
    End Select

End Sub

Private Sub MnuLayerRasterize_Click(Index As Integer)
    
    'Normally, we process menu commands by caption, but this menu is unique because it shares
    ' names with the Layer > Delete menu (e.g. "Current layer"); as such, we must explicitly
    ' request actions.
     Select Case Index
    
        'Current layer
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_rasterizecurrent"
            
        'All layers
        Case 1
            Menus.ProcessDefaultAction_ByName "layer_rasterizeall"
            
    End Select
    
End Sub

Private Sub MnuLayerSize_Click(Index As Integer)
    
    'Normally, we process menu commands by caption, but this menu is unique because it shares
    ' names with the Image > Orientation menu; as such, we must explicitly request actions
    Select Case Index
    
        'Reset to actual size
        Case 0
            Menus.ProcessDefaultAction_ByName "layer_resetsize"
        
        '<separator>
        Case 1
            
        'Standard resize
        Case 2
            Menus.ProcessDefaultAction_ByName "layer_resize"
        
        'Content-aware resize
        Case 3
            Menus.ProcessDefaultAction_ByName "layer_contentawareresize"
    
    End Select
    
End Sub

'Light and shadow effect menu
Private Sub MnuLightShadow_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuLightShadow(Index).Caption
End Sub

Private Sub MnuPixelate_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuPixelate(Index).Caption
End Sub

Private Sub mnuRecentMacros_Click(Index As Integer)
    
    'Load the MRU Macro path that correlates to this index.  (If one is not found, a null string is returned)
    Dim tmpString As String
    tmpString = g_RecentMacros.GetSpecificMRU(Index)
    
    'Check - just in case - to make sure the path isn't empty
    If (LenB(tmpString) <> 0) Then Macros.PlayMacroFromFile tmpString
    
End Sub

Private Sub MnuView_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuView(Index).Caption
End Sub

Private Sub MnuWindowToolbox_Click(Index As Integer)
    
    'Because this is a checkbox-based menu, we handle its commands specially
    Select Case Index
    
        'Toggle toolbox visibility
        Case 0
            Menus.ProcessDefaultAction_ByName "window_displaytoolbox"
        
        '<separator>
        Case 1
        
        'Toggle category labels
        Case 2
            Menus.ProcessDefaultAction_ByName "window_displaytoolcategories"
        
        '<separator>
        Case 3
        
        'Changes to button size (small, normal, large)
        Case 4
            Menus.ProcessDefaultAction_ByName "window_smalltoolbuttons"
        Case 5
            Menus.ProcessDefaultAction_ByName "window_normaltoolbuttons"
        Case 6
            Menus.ProcessDefaultAction_ByName "window_largetoolbuttons"
             
            
    End Select
    
End Sub

Private Sub HotkeyManager_Accelerator(ByVal acceleratorIndex As Long)
    
    'Accelerators are divided into three groups, and they are processed in the following order:
    ' 1) Direct processor strings.  These are automatically submitted to the software processor.
    ' 2) Non-processor directives that can be fired if no images are present (e.g. Open, Paste)
    ' 3) Non-processor directives that require an image.

    '***********************************************************
    'Accelerators that are direct processor strings are handled automatically
    
    With HotkeyManager
    
        If .IsProcessorString(acceleratorIndex) Then
            
            'If the action requires an open image, check for that first
            If .IsImageRequired(acceleratorIndex) Then
                If (Not PDImages.IsImageActive()) Then Exit Sub
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
            If (g_CurrentTool = COLOR_PICKER) Then toolbar_Toolbox.SelectNewTool ND_MEASURE Else toolbar_Toolbox.SelectNewTool COLOR_PICKER
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
        ElseIf Strings.StringsEqual(keyName, "tool_activate_gradient", True) Then
            toolbar_Toolbox.SelectNewTool PAINT_GRADIENT
        
        'Search
        ElseIf Strings.StringsEqual(keyName, "tool_search", True) Then
            toolbar_Layers.SetFocusToSearchBox
            
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
        If (Not PDImages.IsImageActive()) Then Exit Sub
        
        'Fit on screen
        If .HotKeyName(acceleratorIndex) = "FitOnScreen" Then FitOnScreen
        
        'Zoom in
        If .HotKeyName(acceleratorIndex) = "Zoom_In" Then Call MnuView_Click(3)
        
        'Zoom out
        If .HotKeyName(acceleratorIndex) = "Zoom_Out" Then Call MnuView_Click(4)
        
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
        If .HotKeyName(acceleratorIndex) = "Next_Image" Then PDImages.MoveToNextImage True
        If .HotKeyName(acceleratorIndex) = "Prev_Image" Then PDImages.MoveToNextImage False
    
    End With
        
End Sub

'Note that FormMain is only loaded after pdMain.Main() has triggered.  Look there for the *true* start of the program.
Private Sub Form_Load()
    
    On Error GoTo FormMainLoadError
    
    '*************************************************************************************************************************************
    ' Start by rerouting control to "LoadTheProgram", which initializes all key PD systems
    '*************************************************************************************************************************************
    
    'The bulk of the loading code actually takes place inside the main module's ContinueLoadingProgram() function
    Dim suspendAdditionalMessages As Boolean
    If PDMain.ContinueLoadingProgram(suspendAdditionalMessages) Then
    
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
        
        'Visibility for the options toolbox is automatically set according to the current tool; this is different from
        ' other dialogs. (Note that the .ResetToolButtonStates function checks the relevant preference prior to changing
        ' the window state, so all cases are covered nicely.)
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
        If (Not UserPrefs.GetPref_Boolean("Themes", "HasSeenThemeDialog", False)) Then Dialogs.PromptUITheme
        
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
        
        'Before setting focus to the main form, active a focus tracker; this class catches some cases
        ' that VB's built-in focus events do not.
        Set m_FocusDetector = New pdFocusDetector
        m_FocusDetector.StartFocusTracking FormMain.hWnd
        
        'Finally, return focus to the main form
        g_WindowManager.SetFocusAPI FormMain.hWnd
        
        Exit Sub
        
FormMainLoadError:
        PDDebug.LogAction "WARNING!  FormMain_Load experienced an error: #" & Err.Number & ", " & Err.Description
        
    'Something went catastrophically wrong during the load process.  Do not continue with the loading process.
    Else
    
        'Because we can't guarantee that the translation subsystem is loaded, default to
        ' a plain English error message, then terminate the program.
        If (Not suspendAdditionalMessages) Then
            Dim tmpMsg As String, tmpTitle As String
            tmpMsg = "PhotoDemon has experienced a critical startup error." & vbCrLf & vbCrLf & "This can occur when the application is placed in a restricted system folder, like C:\Program Files\ or C:\Windows\.  Because PhotoDemon is a portable application, security precautions require it to operate from a non-system folder, like Desktop, Documents, or Downloads.  Please relocate the program to one of these folders, then try again." & vbCrLf & vbCrLf & "(The application will now close.)"
            tmpTitle = "Startup failure"
            MsgBox tmpMsg, vbOKOnly + vbCritical, tmpTitle
        End If
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
    dropAsNewLayer = Dialogs.PromptForDropAsNewLayer()
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
        If (PDImages.GetNumOpenImages() > 0) Then Message vbNullString
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
    
    Me.AsyncDownloader.Reset
    
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
    If (Not HotkeyManager Is Nothing) Then
        HotkeyManager.DeactivateHook True
        HotkeyManager.ReleaseResources
    End If
    
    'Release the tooltip tracker
    PDDebug.LogAction "Releasing tooltip manager..."
    UserControls.FinalTooltipUnload
    
    'Destroy all custom-created icons and cursors
    PDDebug.LogAction "Destroying custom icons and cursors..."
    IconsAndCursors.UnloadAllCursors
    
    'Destroy all paint-related resources
    PDDebug.LogAction "Destroying paint tool resources..."
    Tools_Paint.FreeBrushResources
    Tools_Fill.FreeFillResources
        
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
    ElseIf (g_CurrentTool = ND_MEASURE) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Measure.hWnd
        Unload toolpanel_Measure
        Set toolpanel_Measure = Nothing
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
    ElseIf (g_CurrentTool = PAINT_GRADIENT) Then
        g_WindowManager.DeactivateToolPanel True, toolpanel_Gradient.hWnd
        Unload toolpanel_Gradient
        Set toolpanel_Gradient = Nothing
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
            PDDebug.LogAction "Forcibly unloading " & tmpForm.Name
            Unload tmpForm
            Set tmpForm = Nothing
        End If

    Next tmpForm
    
    'If an update package was downloaded, this is a good time to apply it
    If Updates.IsUpdatePackageAvailable And PDMain.WasStartupSuccessful() Then
        
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
    PDMain.FinalShutdown
    
    'If a restart is allowed, the last thing we do before exiting is shell a new PhotoDemon instance
    'If g_UserWantsRestart Then Updates.InitiateRestart
    
End Sub

'The top-level adjustments menu provides some shortcuts to most-used items.
Private Sub MnuAdjustments_Click(Index As Integer)
    
    'Check the index; if it's past 10, then it's just a top-level menu for one of the
    ' adjustment categories; we don't want to auto-trigger these, as one menu is named
    ' "Invert" which is also a command in the "Selection" menu!
    If (Index <= 10) Then Menus.ProcessDefaultAction_ByCaption MnuAdjustments(Index).Caption
    
End Sub

Private Sub MnuArtistic_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuArtistic(Index).Caption
End Sub

'All blur filters are handled here
Private Sub MnuBlurFilter_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuBlurFilter(Index).Caption
End Sub

Private Sub MnuClearMRU_Click()
    g_RecentFiles.ClearList
End Sub

'All Color sub-menu entries are handled here.
Private Sub MnuColor_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuColor(Index).Caption
End Sub

'All entries in the Color -> Components sub-menu are handled here
Private Sub MnuColorComponents_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuColorComponents(Index).Caption
End Sub

Private Sub MnuCustomFilter_Click()
    Menus.ProcessDefaultAction_ByName "effects_customfilter"
End Sub

'All distortion filters happen here
Private Sub MnuDistortEffects_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuDistortEffects(Index).Caption
End Sub

Private Sub MnuEdge_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuEdge(Index).Caption
End Sub

Private Sub MnuEdit_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            Menus.ProcessDefaultAction_ByName "edit_undo"
        
        Case 1
            Menus.ProcessDefaultAction_ByName "edit_redo"
        
        Case 2
            Menus.ProcessDefaultAction_ByName "edit_history"
        
        Case 3
            'separator
        
        Case 4
            Menus.ProcessDefaultAction_ByName "edit_repeat"
        
        Case 5
            Menus.ProcessDefaultAction_ByName "edit_fade"
        
        Case 6
            'separator
        
        Case 7
            Menus.ProcessDefaultAction_ByName "edit_cut"
        
        Case 8
            Menus.ProcessDefaultAction_ByName "edit_cutlayer"
        
        Case 9
            Menus.ProcessDefaultAction_ByName "edit_copy"
            
        Case 10
            Menus.ProcessDefaultAction_ByName "edit_copylayer"
            
        Case 11
            Menus.ProcessDefaultAction_ByName "edit_pasteasimage"
            
        Case 12
            Menus.ProcessDefaultAction_ByName "edit_pasteaslayer"
            
        Case 13
            'Top-level "cut/copy/paste special"
            
        Case 14
            'separator
            
        Case 15
            Menus.ProcessDefaultAction_ByName "edit_emptyclipboard"
    
    End Select
    
End Sub

Private Sub MnuEditSpecial_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Menus.ProcessDefaultAction_ByName "edit_specialcut"
        
        Case 1
            Menus.ProcessDefaultAction_ByName "edit_specialcopy"
        
        Case 2
            'Menus.ProcessDefaultAction_ByName "edit_specialpaste"
    
    End Select

End Sub

Private Sub MnuFile_Click(Index As Integer)
    
    Select Case Index
    
        'New
        Case 0
            Menus.ProcessDefaultAction_ByName "file_new"
        
        'Open
        Case 1
            Menus.ProcessDefaultAction_ByName "file_open"
        
        '<Open Recent top-level>
        Case 2
        
        '<Import top-level>
        Case 3
        
        '<separator>
        Case 4
        
        'Close
        Case 5
            Menus.ProcessDefaultAction_ByName "file_close"
        
        'Close all
        Case 6
            Menus.ProcessDefaultAction_ByName "file_closeall"
            
        '<separator>
        Case 7
        
        'Save
        Case 8
            Menus.ProcessDefaultAction_ByName "file_save"
            
        'Save copy (lossless)
        Case 9
            Menus.ProcessDefaultAction_ByName "file_savecopy"
            
        'Save as
        Case 10
            Menus.ProcessDefaultAction_ByName "file_saveas"
        
        'Revert
        Case 11
            Menus.ProcessDefaultAction_ByName "file_revert"
        
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
            Menus.ProcessDefaultAction_ByName "file_print"
            
        '<separator>
        Case 17
        
        'Exit
        Case 18
            Menus.ProcessDefaultAction_ByName "file_quit"
    
    End Select
    
End Sub

Private Sub MnuFileExport_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuFileExport(Index).Caption
End Sub

'All help menu entries are launched from here
Private Sub MnuHelp_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuHelp(Index).Caption
End Sub

Private Sub MnuHistogram_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuHistogram(Index).Caption
End Sub

'All top-level Image menu actions are handled here
Private Sub MnuImage_Click(Index As Integer)

    'Normally we can auto-generate commands from menu names, but the Image menu shares some menu
    ' commands with the Layer menu (e.g. "Resize", "Crop to Selection"), so we manually request
    ' actions for most items in this menu.
    Select Case Index
    
        'Duplicate
        Case 0
            Menus.ProcessDefaultAction_ByName "image_duplicate"
        
        '<separator>
        Case 1
        
        'Resize
        Case 2
            Menus.ProcessDefaultAction_ByName "image_resize"
            
        'Content-aware resize
        Case 3
            Menus.ProcessDefaultAction_ByName "image_contentawareresize"
            
        '<separator>
        Case 4
        
        'Canvas resize
        Case 5
            Menus.ProcessDefaultAction_ByName "image_canvassize"
            
        'Fit canvas to active layer
        Case 6
            Menus.ProcessDefaultAction_ByName "image_fittolayer"
            
        'Fit canvas around all layers
        Case 7
            Menus.ProcessDefaultAction_ByName "image_fitalllayers"
            
        '<separator>
        Case 8
            
        'Crop to selection
        Case 9
            Menus.ProcessDefaultAction_ByName "image_crop"
            
        'Trim empty borders
        Case 10
            Menus.ProcessDefaultAction_ByName "image_trim"
            
        '<separator>
        Case 11
        
        'Top-level Rotate
        Case 12
        
        'Flip horizontal (mirror)
        Case 13
            Menus.ProcessDefaultAction_ByName "image_fliphorizontal"
            
        'Flip vertical
        Case 14
            Menus.ProcessDefaultAction_ByName "image_flipvertical"
            
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
    Menus.ProcessDefaultAction_ByCaption MnuLighting(Index).Caption
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
    Menus.ProcessDefaultAction_ByCaption MnuMetadata(Index).Caption
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuMonochrome(Index).Caption
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuNatureFilter(Index).Caption
End Sub

Private Sub MnuInvert_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuInvert(Index).Caption
End Sub

'All noise filters are handled here
Private Sub MnuNoise_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuNoise(Index).Caption
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    If (LenB(g_RecentFiles.GetFullPath(Index)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(Index)
End Sub

'All rotation actions are initiated here
Private Sub MnuRotate_Click(Index As Integer)

    'Normally, we process menu commands by caption, but this menu is unique because it shares
    ' names with the Layer > Orientation menu; as such, we must explicitly request actions.
    Select Case Index
    
        'Straighten
        Case 0
            Menus.ProcessDefaultAction_ByName "image_straighten"
            
        '<separator>
        Case 1
        
        'Rotate 90
        Case 2
            Menus.ProcessDefaultAction_ByName "image_rotate90"
            
        'Rotate 270
        Case 3
            Menus.ProcessDefaultAction_ByName "image_rotate270"
            
        'Rotate 180
        Case 4
            Menus.ProcessDefaultAction_ByName "image_rotate180"
            
        'Rotate arbitrary
        Case 5
            Menus.ProcessDefaultAction_ByName "image_rotatearbitrary"
            
    End Select
            
End Sub

'All select menu items are handled here
Private Sub MnuSelect_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuSelect(Index).Caption
End Sub

'All Select -> Export menu items are handled here
Private Sub MnuSelectExport_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuSelectExport(Index).Caption
End Sub

'All sharpen filters are handled here
Private Sub MnuSharpen_Click(Index As Integer)
    
    'Because the top-level Sharpen menu shares a name with an item in the child Sharpen menu,
    ' we manually request actions (instead of relying on caption-matching).
    Select Case Index
            
        'Sharpen
        Case 0
            Menus.ProcessDefaultAction_ByName "effects_sharpen"
        
        'Unsharp mask
        Case 1
            Menus.ProcessDefaultAction_ByName "effects_unsharp"
            
    End Select

End Sub

'These menu items correspond to specific zoom values
Private Sub MnuSpecificZoom_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuSpecificZoom(Index).Caption
End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuStylize(Index).Caption
End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuTool(Index).Caption
End Sub

'Add / Remove / Modify a layer's alpha channel with this menu
Private Sub MnuLayerTransparency_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuLayerTransparency(Index).Caption
End Sub

'All "Window" menu items are handled here
Private Sub MnuWindow_Click(Index As Integer)
    Menus.ProcessDefaultAction_ByCaption MnuWindow(Index).Caption
End Sub

'Unlike other toolbars, the image tabstrip has a more complicated window menu, because it is viewable
' under a variety of conditions, and we allow the user to specify any alignment.
Private Sub MnuWindowTabstrip_Click(Index As Integer)
    Select Case Index
        Case 0
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_alwaysshow"
        Case 1
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_shownormal"
        Case 2
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_nevershow"
        Case 3
            '<separator>
        Case 4
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_alignleft"
        Case 5
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_aligntop"
        Case 6
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_alignright"
        Case 7
            Menus.ProcessDefaultAction_ByName "window_imagetabstrip_alignbottom"
    End Select
End Sub

'Update the main form against the current theme.  At present, this is just a thin wrapper against the public ApplyThemeAndTranslations() function,
' but once the form's menu is owner-drawn, we will likely need some custom code to handle menu redraws and translations.
Public Sub UpdateAgainstCurrentTheme()
    
    'Start by notifying all controls that new translations and theme settings are required
    Interface.ApplyThemeAndTranslations Me
    
    'Next, menus must be handled separately, because they are implemented using API menus
    Menus.UpdateAgainstCurrentTheme
    IconsAndCursors.ApplyAllMenuIcons
    IconsAndCursors.ResetMenuIcons
    
End Sub
