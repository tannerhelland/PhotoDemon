VERSION 5.00
Begin VB.Form FormMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "PhotoDemon by Tanner Helland - www.photodemon.org"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   765
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
   Icon            =   "MainWindow.frx":0000
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
      Caption         =   "File"
      Begin VB.Menu MnuFile 
         Caption         =   "New..."
         Index           =   0
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Open..."
         Index           =   1
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Open recent"
         Index           =   2
         Begin VB.Menu MnuRecentFileList 
            Caption         =   "empty"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MnuRecentFiles 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MnuRecentFiles 
            Caption         =   "Open all recent images"
            Index           =   1
         End
         Begin VB.Menu MnuRecentFiles 
            Caption         =   "Clear recent image list"
            Index           =   2
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Import"
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
         Caption         =   "Close"
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
         Caption         =   "Save"
         Index           =   8
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save copy (lossless)"
         Index           =   9
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Save as..."
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
            Caption         =   "Image to file..."
            Index           =   0
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Layers to files..."
            Index           =   1
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Animation..."
            Index           =   3
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Color lookup..."
            Index           =   5
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Color profile..."
            Index           =   6
         End
         Begin VB.Menu MnuFileExport 
            Caption         =   "Palette..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Batch operations"
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
         Caption         =   "Print..."
         Index           =   16
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Exit"
         Index           =   18
      End
   End
   Begin VB.Menu MnuEditTop 
      Caption         =   "Edit"
      Begin VB.Menu MnuEdit 
         Caption         =   "Undo"
         Index           =   0
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Redo"
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
         Caption         =   "Cut"
         Index           =   7
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Cut merged"
         Index           =   8
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Copy"
         Index           =   9
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Copy merged"
         Index           =   10
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Paste"
         Index           =   11
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Paste to new image"
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
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu MnuEditSpecial 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuEditSpecial 
            Caption         =   "Empty clipboard"
            Index           =   4
         End
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Clear"
         Index           =   15
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Content-aware fill..."
         Index           =   16
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Fill..."
         Index           =   17
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Stroke..."
         Index           =   18
      End
   End
   Begin VB.Menu MnuImageTop 
      Caption         =   "Image"
      Begin VB.Menu MnuImage 
         Caption         =   "Duplicate"
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
            Caption         =   "Rotate 90 clockwise"
            Index           =   2
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Rotate 90 counter-clockwise"
            Index           =   3
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Rotate 180"
            Index           =   4
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Rotate arbitrary..."
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
         Caption         =   "Merge visible layers"
         Index           =   16
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flatten image..."
         Index           =   17
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Animation..."
         Index           =   19
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Compare"
         Index           =   20
         Begin VB.Menu MnuImageCompare 
            Caption         =   "Create color lookup..."
            Index           =   0
         End
         Begin VB.Menu MnuImageCompare 
            Caption         =   "Similarity..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Metadata"
         Index           =   21
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
      Begin VB.Menu MnuImage 
         Caption         =   "Show in file manager..."
         Index           =   22
      End
   End
   Begin VB.Menu MnuLayerTop 
      Caption         =   "Layer"
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
         Begin VB.Menu MnuLayerNew 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Layer via copy"
            Index           =   8
         End
         Begin VB.Menu MnuLayerNew 
            Caption         =   "Layer via cut"
            Index           =   9
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
         Caption         =   "Replace"
         Index           =   2
         Begin VB.Menu MnuLayerReplace 
            Caption         =   "From clipboard"
            Index           =   0
         End
         Begin VB.Menu MnuLayerReplace 
            Caption         =   "From file..."
            Index           =   1
         End
         Begin VB.Menu MnuLayerReplace 
            Caption         =   "From visible layers"
            Index           =   2
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge up"
         Index           =   4
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Merge down"
         Index           =   5
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Order"
         Index           =   6
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
         Caption         =   "Visibility"
         Index           =   7
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "Show this layer"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "Show only this layer"
            Index           =   2
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "Hide only this layer"
            Index           =   3
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "Show all layers"
            Index           =   5
         End
         Begin VB.Menu MnuLayerVisibility 
            Caption         =   "Hide all layers"
            Index           =   6
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Crop"
         Index           =   9
         Begin VB.Menu MnuLayerCrop 
            Caption         =   "Crop to selection"
            Index           =   0
         End
         Begin VB.Menu MnuLayerCrop 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuLayerCrop 
            Caption         =   "Fit to canvas"
            Index           =   2
         End
         Begin VB.Menu MnuLayerCrop 
            Caption         =   "Trim empty borders"
            Index           =   3
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Orientation"
         Index           =   10
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
         Index           =   11
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
         Begin VB.Menu MnuLayerSize 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuLayerSize 
            Caption         =   "Fit to image"
            Index           =   5
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Transparency"
         Index           =   13
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "From color (chroma key)..."
            Index           =   0
         End
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "From luminance..."
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
         Begin VB.Menu MnuLayerTransparency 
            Caption         =   "Threshold..."
            Index           =   4
         End
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuLayer 
         Caption         =   "Rasterize"
         Index           =   15
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
         Caption         =   "Split"
         Index           =   16
         Begin VB.Menu MnuLayerSplit 
            Caption         =   "Current layer into standalone image"
            Index           =   0
         End
         Begin VB.Menu MnuLayerSplit 
            Caption         =   "All layers into standalone images"
            Index           =   1
         End
         Begin VB.Menu MnuLayerSplit 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuLayerSplit 
            Caption         =   "Other open images into this image (as layers)..."
            Index           =   3
         End
      End
   End
   Begin VB.Menu MnuSelectTop 
      Caption         =   "Select"
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
         Caption         =   "Fill selected area..."
         Index           =   11
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Heal selected area..."
         Index           =   12
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Stroke selection outline..."
         Index           =   13
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Load selection..."
         Index           =   15
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Save current selection..."
         Index           =   16
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Export"
         Index           =   17
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
      Caption         =   "Adjustments"
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
         Begin VB.Menu MnuChannels 
            Caption         =   "Channel mixer..."
            Index           =   0
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "Rechannel..."
            Index           =   1
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "Maximum channel"
            Index           =   3
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "Minimum channel"
            Index           =   4
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuChannels 
            Caption         =   "Shift left"
            Index           =   6
         End
         Begin VB.Menu MnuChannels 
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
            Caption         =   "Color lookup..."
            Index           =   9
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Colorize..."
            Index           =   10
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Photo filter..."
            Index           =   11
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Replace color..."
            Index           =   12
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Sepia..."
            Index           =   13
         End
         Begin VB.Menu MnuColor 
            Caption         =   "Split toning..."
            Index           =   14
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
            Caption         =   "Dehaze..."
            Index           =   2
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Exposure..."
            Index           =   3
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Gamma..."
            Index           =   4
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "HDR..."
            Index           =   5
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Levels..."
            Index           =   6
         End
         Begin VB.Menu MnuLighting 
            Caption         =   "Shadows and highlights..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Map"
         Index           =   17
         Begin VB.Menu MnuMap 
            Caption         =   "Gradient map..."
            Index           =   0
         End
         Begin VB.Menu MnuMap 
            Caption         =   "Palette map..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuAdjustments 
         Caption         =   "Monochrome"
         Index           =   18
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Color to monochrome..."
            Index           =   0
         End
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Monochrome to gray..."
            Index           =   1
         End
      End
   End
   Begin VB.Menu MnuEffectsTop 
      Caption         =   "Effects"
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
         Begin VB.Menu MnuBlur 
            Caption         =   "Box blur..."
            Index           =   0
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Gaussian blur..."
            Index           =   1
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Surface blur..."
            Index           =   2
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Motion blur..."
            Index           =   4
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Radial blur..."
            Index           =   5
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Zoom blur..."
            Index           =   6
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Distort"
         Index           =   2
         Begin VB.Menu MnuDistort 
            Caption         =   "Correct existing distortion..."
            Index           =   0
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Donut..."
            Index           =   2
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Droste..."
            Index           =   3
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Lens..."
            Index           =   4
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Pinch and whirl..."
            Index           =   5
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Poke..."
            Index           =   6
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Ripple..."
            Index           =   7
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Squish..."
            Index           =   8
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Swirl..."
            Index           =   9
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Waves..."
            Index           =   10
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu MnuDistort 
            Caption         =   "Miscellaneous..."
            Index           =   12
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
            Caption         =   "Gradient flow..."
            Index           =   3
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Range filter..."
            Index           =   4
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Trace contour..."
            Index           =   5
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
            Caption         =   "Bump map..."
            Index           =   1
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Cross-screen..."
            Index           =   2
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Rainbow..."
            Index           =   3
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Sunshine..."
            Index           =   4
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Dilate..."
            Index           =   6
         End
         Begin VB.Menu MnuLightShadow 
            Caption         =   "Erode..."
            Index           =   7
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
            Caption         =   "Dust and scratches..."
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
         Begin VB.Menu MnuNoise 
            Caption         =   "Symmetric nearest-neighbor..."
            Index           =   8
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
         Begin VB.Menu MnuPixelate 
            Caption         =   "Pointillize..."
            Index           =   5
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Render"
         Index           =   8
         Begin VB.Menu MnuRender 
            Caption         =   "Clouds..."
            Index           =   0
         End
         Begin VB.Menu MnuRender 
            Caption         =   "Fibers..."
            Index           =   1
         End
         Begin VB.Menu MnuRender 
            Caption         =   "Truchet..."
            Index           =   2
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Sharpen"
         Index           =   9
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen..."
            Index           =   0
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Unsharp mask..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Stylize"
         Index           =   10
         Begin VB.Menu MnuStylize 
            Caption         =   "Antique..."
            Index           =   0
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Diffuse..."
            Index           =   1
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Kuwahara..."
            Index           =   2
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Outline..."
            Index           =   3
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Palette..."
            Index           =   4
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Portrait glow..."
            Index           =   5
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Solarize..."
            Index           =   6
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Twins..."
            Index           =   7
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Vignetting..."
            Index           =   8
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Transform"
         Index           =   11
         Begin VB.Menu MnuEffectTransform 
            Caption         =   "Offset and zoom..."
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
         Index           =   12
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Animation"
         Index           =   13
         Begin VB.Menu MnuEffectAnimation 
            Caption         =   "Background..."
            Index           =   0
         End
         Begin VB.Menu MnuEffectAnimation 
            Caption         =   "Foreground..."
            Index           =   1
         End
         Begin VB.Menu MnuEffectAnimation 
            Caption         =   "Playback speed..."
            Index           =   2
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Custom filter..."
         Index           =   14
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Photoshop (8bf) plugin..."
         Index           =   15
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "Tools"
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
            Caption         =   "empty"
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
         Caption         =   "Animated screen capture..."
         Index           =   9
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Keyboard shortcuts..."
         Index           =   11
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Options..."
         Index           =   12
      End
      Begin VB.Menu MnuTool 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuTool 
         Caption         =   "Developers"
         Index           =   14
         Begin VB.Menu MnuDevelopers 
            Caption         =   "View debug log for this session..."
            Index           =   0
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Theme editor..."
            Index           =   2
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Build theme package..."
            Index           =   3
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuDevelopers 
            Caption         =   "Build standalone package..."
            Index           =   5
         End
         Begin VB.Menu MnuTest 
            Caption         =   "Test"
         End
      End
   End
   Begin VB.Menu MnuViewTop 
      Caption         =   "View"
      Begin VB.Menu MnuView 
         Caption         =   "Fit image on screen"
         Index           =   0
      End
      Begin VB.Menu MnuView 
         Caption         =   "Center image in viewport"
         Index           =   1
      End
      Begin VB.Menu MnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom in"
         Index           =   3
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom out"
         Index           =   4
      End
      Begin VB.Menu MnuView 
         Caption         =   "Zoom to value"
         Index           =   5
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
         Index           =   6
      End
      Begin VB.Menu MnuView 
         Caption         =   "Show rulers"
         Index           =   7
      End
      Begin VB.Menu MnuView 
         Caption         =   "Show status bar"
         Index           =   8
      End
      Begin VB.Menu MnuView 
         Caption         =   "Show extras"
         Index           =   9
         Begin VB.Menu MnuShow 
            Caption         =   "Layer edges"
            Index           =   0
         End
         Begin VB.Menu MnuShow 
            Caption         =   "Smart guides"
            Index           =   1
         End
      End
      Begin VB.Menu MnuView 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuView 
         Caption         =   "Snap"
         Index           =   11
      End
      Begin VB.Menu MnuView 
         Caption         =   "Snap to"
         Index           =   12
         Begin VB.Menu MnuSnap 
            Caption         =   "Canvas edges"
            Index           =   0
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "Centerlines"
            Index           =   1
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "Layers"
            Index           =   2
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "Angle 90"
            Index           =   4
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "Angle 45"
            Index           =   5
         End
         Begin VB.Menu MnuSnap 
            Caption         =   "Angle 30"
            Index           =   6
         End
      End
   End
   Begin VB.Menu MnuWindowTop 
      Caption         =   "Window"
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
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu MnuWindowToolbox 
            Caption         =   "Medium buttons"
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
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu MnuWindowOpen 
         Caption         =   "empty"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuHelpTop 
      Caption         =   "Help"
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
         Caption         =   "Ask a question..."
         Index           =   3
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Check for updates..."
         Index           =   4
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit bug report or feedback..."
         Index           =   5
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "PhotoDemon forum..."
         Index           =   7
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "PhotoDemon license and terms of use..."
         Index           =   8
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "PhotoDemon source code..."
         Index           =   9
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "PhotoDemon website..."
         Index           =   10
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Third-party libraries..."
         Index           =   12
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "About..."
         Index           =   14
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PhotoDemon is Copyright 1999-2026 by Tanner Helland, tannerhelland.com
'
'PhotoDemon's INSTALL file provides important information for developers:
' https://github.com/tannerhelland/PhotoDemon/blob/master/INSTALL.md
'
'PhotoDemon's LICENSE file provides important information on code licensing and redistribution:
' https://github.com/tannerhelland/PhotoDemon/blob/master/LICENSE.md
'
'For further information, visit https://photodemon.org or https://github.com/tannerhelland/PhotoDemon
'
'***************************************************************************
'Primary PhotoDemon Interface
'Copyright 2002-2026 by Tanner Helland
'Created: 15/September/02
'Last updated: 06/October/21
'Last update: remove all hotkey code - it's now handled elsewhere!
'
'This is PhotoDemon's main window.  In actuality, it contains relatively little code.  Its primary purpose
' is sending parameters to other, more interesting sections of the program.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This main dialog houses a few timer objects; these can be started and/or stopped by external functions.
' See the timer start/stop functions for additional details.
Private WithEvents m_MetadataTimer As pdTimer
Attribute m_MetadataTimer.VB_VarHelpID = -1

'If the user has set the preference for single-session behavior, subsequent PD sessions will forward
' their command-lines to us via a named pipe.  Pipe messages will be raised as events, and the pdStream
' object is automatically filled with data from the second instance as it arrives.
Private WithEvents m_OtherSessions As pdPipe
Attribute m_OtherSessions.VB_VarHelpID = -1
Private m_SessionStream As pdStream
Private m_uniqueSessionName As String

'Focus detection is used to correct some hotkey behavior.  (Specifically, when this form loses focus,
' PD resets its hotkey tracker; this solves problems created by Alt+Tabbing away from the program,
' and PD thinking the Alt-key is still down when the user returns.)
Private WithEvents m_FocusDetector As pdFocusDetector
Attribute m_FocusDetector.VB_VarHelpID = -1

'FormMain is loaded by PDMain.Main().  Look there for the *true* start of the program.
Private Sub Form_Load()
    
    'PhotoDemon is always developed with the VB6 IDE set to "Break on All Errors"
    ' (Tools > Options > General > Error Trapping > Break on All Errors)
    '
    'VB6-raised errors are *never* intended behavior in PhotoDemon.  The program is designed to check
    ' potential problem-states in advance, and deal with them preemptively instead of plowing ahead
    ' and waiting for errors to raise.  Built-in VB6 error handling constructs exist only as an
    ' absolute last resort, and they are only used in a select few places - like here, where I've
    ' included an error clause in case new developers want to "play around" with PhotoDemon and
    ' accidentally break something.
    '
    'From this point forward in the program, do not rely on VB6 error handling.  Validate incoming
    ' data and input properly, and deal with problem states *before* they cause errors.
    On Error GoTo FormMainLoadError
    
    '*************************************************************************************************************************************
    ' Start by rerouting control to "ContinueLoadingProgram", which initializes all key PD systems
    '*************************************************************************************************************************************
    
    'The bulk of the loading code actually takes place inside the main module's ContinueLoadingProgram() function
    Dim suspendAdditionalMessages As Boolean
    If PDMain.ContinueLoadingProgram(suspendAdditionalMessages) Then
        
        'With the program successfully initialized, we now need to (optionally) start listening
        ' for other active PhotoDemon sessions.  The user can set an option (in Tools > Options)
        ' to only allow one PhotoDemon instance at a time.  Attempting to load an image into a
        ' new PhotoDemon instance will instead route that image here, to this instance.
        '
        '(Note that this check *also* occurs in the IDE, contingent on a compile-time constant
        ' in the Mutex module.)
        If Mutex.IsThisOnlyInstance() Then
            
            'Write a unique session name to the user prefs file; other instances will use this
            ' to connect to our named pipe.
            m_uniqueSessionName = OS.GetArbitraryGUID()
            UserPrefs.WritePreference "Core", "SessionID", m_uniqueSessionName
            
            'Prep a pdStream object.  It will receive any incoming pipe data.
            Set m_SessionStream = New pdStream
            m_SessionStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
            
            'Start listening for other PhotoDemon sessions.  (Note that this function will
            ' check the user's preference before initializing single-session mode, but we
            ' initialize the session stream in advance in case the user changes the option
            ' mid-session - we'll already have everything ready to go for them!)
            Me.ChangeSessionListenerState True, True
            
        End If
        
        '*************************************************************************************************************************************
        ' Now that all program engines are initialized, we can finally display this window
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Registering toolbars with the window manager..."
        
        'Now that the main form has been correctly positioned on-screen, position all toolbars and the
        ' primary canvas to match, then display the window.
        g_WindowManager.SetAutoRefreshMode True
        FormMain.UpdateMainLayout
        g_WindowManager.SetAutoRefreshMode False
        
        'DWM may cause issues inside the IDE, so forcibly refresh the main form after displaying it.
        ' (The DoEvents fixes an unpleasant flickering issue on Windows Vista/7 when the DWM isn't
        ' running full Aero, e.g. "Classic Mode".)
        FormMain.Show vbModeless
        FormMain.Refresh
        DoEvents
        
        'Visibility for the options toolbox is automatically set according to the current tool;
        ' this is different from the left and right toolboxes. (Note that the .ResetToolButtonStates
        ' function checks the relevant preference prior to changing the window state, so all cases
        ' are covered correctly.)
        toolbar_Toolbox.ResetToolButtonStates
        
        'With all toolboxes loaded, we can safely reactivate automatic syncing of toolboxes and the
        ' main window.
        g_WindowManager.SetAutoRefreshMode True
        
        
        '*************************************************************************************************************************************
        ' Make sure the user's previous session (if any) terminated successfully
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Checking for old autosave data..."
        Autosaves.InitializeAutosave
        
        'PD's internal debug logger will now be active if...
        ' 1) this is a nightly build, or...
        ' 2) the last session crashed, or...
        ' 3) the user manually activated debug logging
        '
        'Activate the Tools > Developer > View debug log for current session menu accordingly
        Menus.SetMenuEnabled "tools_viewdebuglog", UserPrefs.GenerateDebugLogs()
        
        '*************************************************************************************************************************************
        ' Next, analyze the command line and load any passed image files
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
        
        'Update checks only work in (default) portable mode, because we require write access to our
        ' own folder.  Note that PhotoDemon will still attempt to run even if the user places us in
        ' a restricted folder (e.g. /Program Files), but we suspend updates and silently remap PD's
        ' user folder to the local Users folder instead.
        If (Not UserPrefs.IsNonPortableModeActive()) Then Updates.StandardUpdateChecks
        
        
        '*************************************************************************************************************************************
        ' Display any final messages and/or warnings
        '*************************************************************************************************************************************
        
        Message vbNullString
        FormMain.Refresh
        DoEvents
        
        '*************************************************************************************************************************************
        ' Next, see if we need to display the language/theme selection dialog
        '*************************************************************************************************************************************
        
        'In v7.0, a new "choose your language and UI theme" dialog was added to the project.
        ' We automatically show it to first-time users to help them set everything up just the way they want.
        If (Not UserPrefs.GetPref_Boolean("Themes", "HasSeenThemeDialog", False)) Then Dialogs.PromptUITheme
        
        '*************************************************************************************************************************************
        'For developers only, calculate some debug information and warn about running from the IDE
        '*************************************************************************************************************************************
        
        PDDebug.LogAction "Current PD custom control count: " & UserControls.GetPDControlCount
        
        'Because people may be using this code in the IDE, warn them about the consequences of doing so
        If (Not OS.IsProgramCompiled) Then
            If (UserPrefs.GetPref_Boolean("Core", "Display IDE Warning", True)) Then Dialogs.DisplayIDEWarning
        End If
        
        'Because various user preferences may have been modified during the load process (to account for
        ' failure states, system configurations, etc), write a copy of our potentially-modified
        ' preference list out to file.
        UserPrefs.ForceWriteToFile False
        
        'In debug mode, note that we are about to turn control over to the user
        PDDebug.LogAction "Program initialization complete.  Second baseline memory measurement:"
        PDDebug.LogAction vbNullString, PDM_Mem_Report
        
        'Before setting focus to the main form, activate a focus tracker.
        ' (PD uses this class to catch some focus cases that VB's built-in focus events do not.)
        Set m_FocusDetector = New pdFocusDetector
        m_FocusDetector.StartFocusTracking FormMain.hWnd
        
        'Finally, return focus to the main form.  The app is ready for interaction!
        If (PDImages.GetNumOpenImages > 0) Then
            FormMain.MainCanvas(0).SetFocusToCanvasView
        Else
            
            m_FocusDetector.SetFocusManually
            
            'Focus is weird in the IDE; activate it manually to work around issues under *some* versions of Windows
            If (Not OS.IsProgramCompiled()) Then m_FocusDetector.ActivateManually
            
        End If
        
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
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

'If the user attempts to close the program, run some checks first.  Specifically, we want to notify them
' about any images with unsaved images.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'An external function handles unloading.  If it fails, we will also cancel our unload.
    Cancel = (Not CanvasManager.CloseAllImages())
    If Cancel Then
        If (PDImages.GetNumOpenImages() > 0) Then Message vbNullString  'Clear any shutdown-related messages
    Else
    
        'Set a public variable to let other functions know that the user has initiated a program-wide shutdown.
        ' (Asynchronous functions in particular need this to prep their own shutdown routines.)
        g_ProgramShuttingDown = True
    
    End If
    
End Sub

'When the main form is resized, we must re-align all toolboxes and the main canvas to match
Private Sub Form_Resize()
    If (Not g_WindowManager Is Nothing) Then
        
        If g_WindowManager.GetAutoRefreshMode Then UpdateMainLayout
        
        'In the IDE it's helpful for me to test layouts against specific screen sizes
        If (Not OS.IsProgramCompiled()) Then
            Dim doNotLocalizeThis As String
            doNotLocalizeThis = "Window size: "
            Message doNotLocalizeThis & g_WindowManager.GetClientWidth(Me.hWnd) & "x" & g_WindowManager.GetClientHeight(Me.hWnd)
        End If
        
    Else
        UpdateMainLayout
    End If
End Sub

'Do not call this function directly.  Use the standard shutdown order to ensure the user has a chance
' to save unsaved work.
Private Sub Form_Unload(Cancel As Integer)
    
    'FYI, this function includes a fair amount of debug code.  It is *critical* that we unload everything
    ' correctly and do not leave any open handles on the user's PC.
    PDDebug.LogAction "Shutdown initiated"
    
    'Store the main window's location to file.  (We will use this in the next session to restore the
    ' program to the same monitor + position.)
    UserPrefs.SetPref_Long "Core", "Last Window State", Me.WindowState
    UserPrefs.SetPref_Long "Core", "Last Window Left", Me.Left / TwipsPerPixelXFix
    UserPrefs.SetPref_Long "Core", "Last Window Top", Me.Top / TwipsPerPixelYFix
    UserPrefs.SetPref_Long "Core", "Last Window Width", Me.Width / TwipsPerPixelXFix
    UserPrefs.SetPref_Long "Core", "Last Window Height", Me.Height / TwipsPerPixelYFix
    
    'Hide the main window to make shutdown appear "faster"
    Me.Visible = False
    Interface.ReleaseResources
    
    'Cancel any pending downloads
    PDDebug.LogAction "Checking for (and terminating) any in-progress downloads..."
    Me.AsyncDownloader.ResetDownloader
    
    'Allow any objects on this form to save preferences and other user data
    PDDebug.LogAction "Asking all FormMain components to write out final user preference values..."
    FormMain.MainCanvas(0).WriteUserPreferences
    Toolboxes.SaveToolboxData
    
    'Release the clipboard manager.  If we are responsible for the current clipboard data, we must manually
    ' upload a copy of all supported formats - which can take a little while.
    PDDebug.LogAction "Shutting down clipboard manager..."
    
    If (Not g_Clipboard Is Nothing) Then
    
        If (g_Clipboard.IsPDDataOnClipboard And OS.IsProgramCompiled) Then
            PDDebug.LogAction "PD's data remains on the clipboard.  Rendering any additional formats now..."
            g_Clipboard.RenderAllClipboardFormatsManually
        End If
    
        Set g_Clipboard = Nothing
        
    End If
    
    'Most core plugins are released as a final step, but ExifTool only matters when images are loaded, and we know
    ' no images are loaded by this point.  Because it takes some time to shut down, trigger it now.
    If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) Then
        ExifTool.TerminateExifTool
        PDDebug.LogAction "ExifTool terminated"
    End If
    
    'Stop tracking hotkeys
    PDDebug.LogAction "Turning off hotkey manager..."
    If (Not HotkeyManager Is Nothing) Then HotkeyManager.ReleaseResources
    
    'Release the tooltip tracker
    PDDebug.LogAction "Releasing tooltip manager..."
    UserControls.FinalTooltipUnload
    
    'Perform any printer-related cleanup
    PDDebug.LogAction "Removing printer temp files..."
    Printing.PerformPrinterCleanup
    
    'Destroy all user-added font resources
    PDDebug.LogAction "Destroying custom fonts..."
    Fonts.ReleaseUserFonts
    
    'Destroy all custom-created icons and cursors
    PDDebug.LogAction "Destroying custom icons and cursors..."
    IconsAndCursors.UnloadAllCursors
    
    'Destroy all paint-related resources
    PDDebug.LogAction "Destroying paint tool resources..."
    Tools_Paint.FreeBrushResources
    Tools_Pencil.FreeBrushResources
    Tools_Clone.FreeBrushResources
    Tools_Fill.FreeFillResources
        
    'Save all MRU lists to the preferences file.  (I've considered doing this as files are loaded,
    ' but the only time that would be an improvement is if the program crashes, and if it does crash,
    ' the user wouldn't want to re-load the problematic image anyway... so we do it here.)
    PDDebug.LogAction "Saving recent file list..."
    If (Not g_RecentFiles Is Nothing) Then
        g_RecentFiles.WriteListToFile
        g_RecentMacros.MRU_SaveToFile
    End If
    
    'Release any Win7+ features (which typically rely on custom run-time interfaces)
    PDDebug.LogAction "Releasing custom Windows 7+ features..."
    OS.StopWin7PlusFeatures
    
    'Tool panels need to be shut down carefully because we've messed with their window longs
    ' (to embed them correctly on their parent window).
    PDDebug.LogAction vbNullString, PDM_Mem_Report
    PDDebug.LogAction "Unloading tool panels..."
    toolbar_Toolbox.FreeAllToolpanels
    
    'With all tool panels unloaded, unload all toolboxes too
    PDDebug.LogAction "Unloading toolboxes..."
    
    'Before unloading toolboxes, we need to reset their window bits.  (These window bits get
    ' toggled by the toolbox module as part of assigning parent/child relationships.)
    If PDMain.WasStartupSuccessful() Then
        Toolboxes.ReleaseToolbox toolbar_Layers.hWnd
        Unload toolbar_Layers
    End If
    Set toolbar_Layers = Nothing
    
    If PDMain.WasStartupSuccessful() Then
        Toolboxes.ReleaseToolbox toolbar_Options.hWnd
        Unload toolbar_Options
    End If
    Set toolbar_Options = Nothing
    
    If PDMain.WasStartupSuccessful() Then
        Toolboxes.ReleaseToolbox toolbar_Toolbox.hWnd
        Unload toolbar_Toolbox
    End If
    Set toolbar_Toolbox = Nothing
    
    'Release this window from the central window manager
    PDDebug.LogAction "Shutting down window manager..."
    If (Not g_WindowManager Is Nothing) Then
        Interface.ReleaseFormTheming Me
        g_WindowManager.UnregisterMainForm Me
    End If
    
    'As a final failsafe, forcibly unload any remaining forms.  (This shouldn't do anything, but it's
    ' a nice way to ensure no mistakes were made.)
    PDDebug.LogAction "Forcibly unloading any remaining forms..."
    
    Dim tmpForm As Form
    For Each tmpForm In Forms

        'Note that there is no need to unload FormMain, as we're about to unload it anyway
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
        Else
            PDDebug.LogAction "WARNING!  One or more errors were encountered while applying an update.  PD has attempted to roll everything back to its original state."
        End If
    End If
    
    'Because PD can now auto-update between runs, it's helpful to log the current program version to the preferences file.  The next time PD runs,
    ' it can compare its version against this value, to infer if an update occurred.
    PDDebug.LogAction "Writing session data to file..."
    UserPrefs.SetPref_String "Core", "LastRunVersion", Updates.GetPhotoDemonVersionCanonical()
    
    'All core PD functions appear to have terminated correctly, so notify the Autosave handler
    ' that this session was clean.  (One exception to this rule is if this is a system-initiated shutdown,
    ' and the user has enabled auto-restart-after-reboot behavior.  We *must* maintain Autosave data
    ' so that we can restore their session as-is.)
    Dim allowAutosaveClean As Boolean: allowAutosaveClean = True
    If (Not g_ThunderMain Is Nothing) Then allowAutosaveClean = Not g_ThunderMain.WasEndSessionReceived(True)
    
    If allowAutosaveClean Then
        PDDebug.LogAction "Final step: writing out new autosave checksum..."
        Autosaves.PurgeOldAutosaveData
        Autosaves.NotifyCleanShutdown
    Else
        PDDebug.LogAction "Suspending autosave purge due to system reboot"
    End If
    
    PDDebug.LogAction "Shutdown appears to be clean.  Turning final control over to pdMain.FinalShutdown()..."
    PDMain.FinalShutdown
    
End Sub

'Whenever the window is resized, we need to update the layout of the primary canvas and all toolboxes.
Public Sub UpdateMainLayout(Optional ByVal resizeToolboxesToo As Boolean = True)

    'If the main form has been minimized, don't refresh anything
    If (FormMain.WindowState = vbMinimized) Then Exit Sub
    
    'As of 7.0, a new, lightweight toolbox manager will calculate idealized window positions for us.
    Dim mainRect As winRect, canvasRect As winRect
    g_WindowManager.GetClientWinRect FormMain.hWnd, mainRect
    Toolboxes.CalculateNewToolboxRects mainRect, canvasRect
    
    'With toolbox positions successfully calculated, we can now synchronize each toolbox to its calculated rect.
    If resizeToolboxesToo Then
        Toolboxes.PositionToolbox PDT_LeftToolbox, toolbar_Toolbox.hWnd, FormMain.hWnd
        Toolboxes.PositionToolbox PDT_RightToolbox, toolbar_Layers.hWnd, FormMain.hWnd
        Toolboxes.PositionToolbox PDT_TopToolbox, toolbar_Options.hWnd, FormMain.hWnd
    End If
    
    'Similarly, we can drop the canvas into place using the helpful rect provided by the toolbox module.
    ' Note that resizing the canvas rect will automatically trigger a redraw of the viewport, as necessary.
    With canvasRect
        FormMain.MainCanvas(0).SetPositionAndSize .x1, .y1, .x2 - .x1, .y2 - .y1
    End With
    
    'If all three toolboxes are hidden, Windows may try to hide the main window as well.  Reset focus manually.
    If Toolboxes.AreAllToolboxesHidden And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI FormMain.hWnd
    
End Sub

'Whenever the asynchronous downloader completes its work, we forcibly release all resources associated
' with the download process.
Private Sub AsyncDownloader_FinishedAllItems(ByVal allDownloadsSuccessful As Boolean)
    
    'Core program updates are handled specially, so their resources can be freed without question.
    AsyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK"
    AsyncDownloader.FreeResourcesForItem "PROGRAM_UPDATE_CHECK_USER"
    
    FormMain.MainCanvas(0).SetNetworkState False
    PDDebug.LogAction "All downloads complete."
    
End Sub

'When an asynchronous download completes, deal with it here
Private Sub AsyncDownloader_FinishedOneItem(ByVal downloadSuccessful As Boolean, ByVal entryKey As String, ByVal OptionalType As Long, downloadedData() As Byte, ByVal savedToThisFile As String)
    
    'On a typical PD install updates are checked every session, but users can specify a larger interval
    ' (from the Tools > Options menu).  To honor that preference, whenever an update check completes,
    ' we write the current date out to the preferences file so subsequent sessions can limit their update
    ' frequency accordingly.
    If Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK", True) Or Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK_USER", True) Then
        
        If downloadSuccessful Then
        
            'The update file downloaded correctly.  Write today's date to the central preferences file
            ' so we can correctly calculate weekly or monthly update checks for users that request it.
            PDDebug.LogAction "Update file download complete.  Update information has been saved at " & savedToThisFile
            UserPrefs.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
            
            'Pull the file contents (which are just simple XML) into a VB string.
            Dim updateXML As String
            updateXML = StrConv(downloadedData, vbUnicode)
            
            'Offload the rest of the check to a separate update function.  It will initiate subsequent downloads as necessary.
            Dim updateAvailable As Boolean
            updateAvailable = Updates.ProcessProgramUpdateFile(updateXML)
            
            'If the user initiated the download, display a modal notification now
            If Strings.StringsEqual(entryKey, "PROGRAM_UPDATE_CHECK_USER", True) Then
                
                If updateAvailable Then
                    Message "A new version of PhotoDemon is available.  The update is automatically processing in the background..."
                    PDMsgBox "A new version of PhotoDemon is available!" & vbCrLf & vbCrLf & "The update is automatically processing in the background.  You will receive a new notification when it completes.", vbOKOnly Or vbInformation, "PhotoDemon Updates"
                Else
                    Message "This copy of PhotoDemon is up to date."
                    PDMsgBox "This copy of PhotoDemon is the newest version available." & vbCrLf & vbCrLf & "(Current version: %1.%2.%3)", vbOKOnly Or vbInformation, "PhotoDemon Updates", VBHacks.AppMajor_Safe(), VBHacks.AppMinor_Safe(), VBHacks.AppRevision_Safe()
                End If
                
                'If the update managed to download while the reader was staring at the message box,
                ' display the restart notification immediately.  (This is more likely than you'd think.)
                If Updates.IsUpdateReadyToInstall() Then Updates.DisplayUpdateNotification
                
            End If
            
        Else
            
            'If the update check fails on XP due to known secure channel issues - error 0x80072F7D -
            ' display a meaningful message to the user.  (Note that this message is deliberately left
            ' unlocalized - it's a lot of text, and it will only ever be displayed to XP users, who
            ' comprise an increasingly tiny percentage of PD users.)
            If (AsyncDownloader.GetLastErrorNumber = -2147012739) Then
                
                PDDebug.LogAction "Can't download update file; Windows XP is likely the problem."
                
                Dim xpErrorMsg As pdString
                Set xpErrorMsg = New pdString
                xpErrorMsg.AppendLine "Unfortunately, PhotoDemon is unable to check for updates.  This copy of Windows lacks the necessary security protocol (TLS 1.2) to securely connect to the update server."
                xpErrorMsg.AppendLineBreak
                xpErrorMsg.AppendLine "To protect your privacy and safety, PhotoDemon will not auto-update without a secure connection.  If you are using Windows 7 or later, please run Windows Update to ensure that all security patches have been applied to this PC."
                xpErrorMsg.AppendLineBreak
                xpErrorMsg.AppendLine "If you are using Windows XP, Microsoft has unfortunately chosen not to provide an update with these security features.  You will need to manually download new versions of PhotoDemon from photodemon.org using a secure 3rd-party web browser (like Mozilla Firefox)."
                xpErrorMsg.AppendLineBreak
                xpErrorMsg.Append "(To prevent this message from interrupting you again, PhotoDemon will now deactivate automatic updates.  You can always reactivate this feature from the Tools > Options menu.)"
                UserPrefs.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
                PDMsgBox xpErrorMsg.ToString(), vbInformation Or vbOKOnly Or vbApplicationModal, "Updates unavailable"
                
            Else
                PDDebug.LogAction "Update file was not downloaded.  asyncDownloader returned this error message: " & AsyncDownloader.GetLastErrorNumber & " - " & AsyncDownloader.GetLastErrorDescription
            End If
            
        End If
    
    'If PROGRAM_UPDATE_CHECK (above) finds updated program or plugin files, it will trigger their download.
    ' When the download arrives, we can start patching immediately.
    ElseIf (OptionalType = PD_PATCH_IDENTIFIER) Then
        
        If downloadSuccessful Then
            
            'Notify the software updater that an update package was downloaded successfully.
            ' It needs to know this so it can actually patch all necesssary files when PD closes.
            Updates.NotifyUpdatePackageAvailable savedToThisFile
            
            'Display a notification so the user can choose to restart immediately for any new features
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

'When loading an image, metadata is parsed by ExifTool in a separate process.  This timer can fire
' at a shorter interval (and it's entirely dependent on other parts of the image load process yielding)
' because it's not time-sensitive - its goal is just to collect image metadata in the background.
Public Sub StartMetadataTimer()
    
    If (m_MetadataTimer Is Nothing) Then
        Set m_MetadataTimer = New pdTimer
        m_MetadataTimer.Interval = 250
    End If
    
    If (Not m_MetadataTimer.IsActive) Then m_MetadataTimer.StartTimer
    
End Sub

Private Sub m_FocusDetector_AppGotFocusReliable()

    'Ignore focus changes at shutdown time
    If (Not g_ProgramShuttingDown) Then
        
        'Restore any relevant UI animations
        If Selections.SelectionsAllowed(False) Then
            If PDImages.GetActiveImage.IsSelectionActive() Then
                PDImages.GetActiveImage.MainSelection.NotifyAnimationsAllowed SelectionUI.GetUISetting_Animate()
            End If
        End If
        
    End If
    
End Sub

Private Sub m_FocusDetector_AppLostFocusReliable()

    'Ignore focus changes at shutdown time
    If (Not g_ProgramShuttingDown) Then
        
        'Turn off selection animations in the main window
        If Selections.SelectionsAllowed(False) Then
            If PDImages.GetActiveImage.IsSelectionActive() Then
                PDImages.GetActiveImage.MainSelection.NotifyAnimationsAllowed False
            End If
        End If
        
    End If
    
End Sub

'This metadata timer is a final failsafe for images with huge metadata collections that take a long time
' to parse.  If an image has successfully loaded but its metadata parsing is still in-progress, PD's image
' load function will activate this timer.  The timer will wait (asynchronously) for metadata parsing to finish,
' and when it does, it will copy the metadata into the active pdImage object, then turn itself off.
Private Sub m_MetadataTimer_Timer()
    
    If g_ProgramShuttingDown Then Exit Sub
    
    'Until ExifTool reports completion, we just need to poll in the background
    If ExifTool.IsMetadataFinished Then
    
        'Start by disabling this timer (as it's no longer needed)
        m_MetadataTimer.StopTimer
        
        'Cache the current UI message (if any) because we want to replace it with a metadata-centric one
        Dim prevMessage As String
        prevMessage = Interface.GetLastFullMessage()
        Message "Importing metadata..."
        
        'Ask ExifTool to send all received metadata to its proper parent image
        ExifTool.FinishAsyncMetadataLoading
        
        'Update the interface to match the active image.  (This must be done if things like GPS tags were found in the metadata,
        ' because their presence affects the enabling of certain metadata-related menu entries.)
        Interface.SyncInterfaceToCurrentImage
        
        'Restore the original on-screen message and exit
        Interface.Message prevMessage
        
    End If

End Sub

'As of 2021, hotkeys can be blindly passed to PD's high-level action processor.
' (The new action processor handles all validation and routing duties.)
Private Sub HotkeyManager_HotkeyPressed(ByVal hotkeyID As Long)
    Actions.LaunchAction_ByName Hotkeys.GetHotKeyAction(hotkeyID), pdas_Hotkey
End Sub

'This listener will raise events when PD is in single-session mode and another session is initiated.
' The other session will forward any "open this image" requests passed on its command line.
Private Sub m_OtherSessions_BytesArrived(ByVal initStreamPosition As Long, ByVal numOfBytes As Long)
    
    'Failsafe checks only
    If g_ProgramShuttingDown Then Exit Sub
    If (m_SessionStream Is Nothing) Then Exit Sub
    If (Not m_SessionStream.IsOpen()) Then Exit Sub
    
    'This pipe is used by other PD sessions to forward their command-line contents to us,
    ' if the user has enabled single-session mode.  We do not want to retrieve the pipe's
    ' data until the full pipe contents have arrived; for this reason, we don't care about
    ' the passed pipe stream position - we only care about the size of the pipe matching
    ' the long-type value at the start of the stream (which is the size of the passed string).
    m_SessionStream.SetPosition 0, FILE_BEGIN
    
    'If at least four bytes are available, retrieve them; they're the size of the command line placed
    ' into the pipe buffer.
    If (m_SessionStream.GetStreamSize() >= 4) Then
        
        Dim msgSize As Long
        msgSize = m_SessionStream.ReadLong()
        
        'If the stream has received the full pipe message, retrieve all data, then blank out
        ' the stream.
        If (m_SessionStream.GetStreamSize() >= msgSize + 4) Then
            
            'If arguments were passed, extract them to a string stack
            If (msgSize > 0) Then
            
                'Iterate strings
                Dim numArguments As Long
                numArguments = m_SessionStream.ReadLong()
                
                If (numArguments > 0) Then
                    
                    Dim i As Long, argSize As Long, argString As String
                    For i = 0 To numArguments - 1
                        
                        argSize = m_SessionStream.ReadLong()
                        If (argSize > 0) Then
                            
                            'This is a path to an image.  If PD is idle, attempt to load it.
                            argString = m_SessionStream.ReadString_UTF8(argSize, False)
                            If (Not Processor.IsProgramBusy()) Then Loading.LoadFileAsNewImage argString
                        End If
                        
                    Next i
                    
                    'Bring FormMain to the forefront to indicate that something has happened.
                    If (Not g_WindowManager Is Nothing) Then g_WindowManager.BringWindowToForeground FormMain.hWnd
                    
                End If
                
            Else
                PDDebug.LogAction "Received 0-length string from second instance; no action taken."
            End If
            
            'Reset the stream in case other sessions connect in the future
            m_SessionStream.SetPosition 0
            m_SessionStream.SetSizeExternally 0
            
            'Forcibly disconnect from the client (required regardless of client connection state,
            ' e.g. even if the client has disconnected, we still need to disconnect from this
            ' instance as well), then recreate the pipe.  (This could be avoided by switching
            ' to asynchronous I/O, but I haven't written that code... yet.)
            m_OtherSessions.DisconnectFromClient
            m_OtherSessions.ClosePipe
            If m_OtherSessions.CreatePipe(m_uniqueSessionName, m_SessionStream, pom_ClientToServer Or pom_FlagFirstPipeInstance, pm_TypeByte Or pm_ReadModeByte Or pm_RemoteClientsReject, 1) Then m_OtherSessions.Server_WaitForResponse 1000
            
        Else
            'Do nothing; we just need to chill and wait for the rest of the stream to arrive
        End If
        
    End If
    
End Sub

'Start/stop multisession listening support.  Note that this behavior can be overruled by the
' user's current preference for single-session behavior (e.g. if the user doesn't want
' single-session behavior, these functions are effectively nops.)
'
'That said, it is important to call these functions before PD engages something like a
' batch process, because opening new images in the midst of a batch process can produce
' very unpredictable behavior.  Same goes for e.g. an effect window being active, because PD
' isn't equipped to deal with the active image changing while a modal dialog is live.
'
'Pass FALSE to turn off multi-session detection; TRUE to turn it on.  Note that if the current
' state is set to OFF, parallel PD sessions will be allowed to start regardless of the user's
' current setting - this is a "convenience", as e.g. it allows the user to edit photos while
' a batch process is running in the background.
Public Sub ChangeSessionListenerState(ByVal newState As Boolean, Optional ByVal isFirstCall As Boolean = False)
    
    'Before doing anything else, check the user's preference for multi-session behavior.
    ' If the user does *not* want a multi-session listener, we can simply turn off the
    ' current listener (if any).
    If UserPrefs.GetPref_Boolean("Loading", "Single Instance", False) Then
        
        'The caller wants single-session behavior.
        If isFirstCall Then Set m_OtherSessions = New pdPipe
        
        'In the IDE, it is possible for the session listener to *never* be enabled (because I don't
        ' need async IDE shenanigans in my life), so bail if the pipe listener was never created.
        If (m_OtherSessions Is Nothing) Or (m_SessionStream Is Nothing) Then Exit Sub
        
        'Turn on multi-session listeners
        If newState Then
            
            'Create the pipe anew
            If m_OtherSessions.CreatePipe(m_uniqueSessionName, m_SessionStream, pom_ClientToServer Or pom_FlagFirstPipeInstance, pm_TypeByte Or pm_ReadModeByte Or pm_RemoteClientsReject, 1) Then
                If isFirstCall Then PDDebug.LogAction "Multi-session listener started successfully."
                m_OtherSessions.Server_WaitForResponse 1000
            Else
                If isFirstCall Then PDDebug.LogAction "WARNING!  FormMain couldn't create a named pipe for single session support!"
            End If
                
        'Suspend the pipe
        Else
            If (Not m_OtherSessions Is Nothing) Then
                m_OtherSessions.DisconnectFromClient
                m_OtherSessions.ClosePipe
            End If
        End If
    
    'The caller wants to allow multiple parallel instances.  Turn off the multi-session listener.
    Else
        If (Not m_OtherSessions Is Nothing) Then
            m_OtherSessions.DisconnectFromClient
            m_OtherSessions.ClosePipe
        End If
    End If
    
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

Private Sub MnuAdjustments_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_autocorrect"
        Case 1
            Actions.LaunchAction_ByName "adj_autoenhance"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "adj_blackandwhite"
        Case 4
            Actions.LaunchAction_ByName "adj_bandc"
        Case 5
            Actions.LaunchAction_ByName "adj_colorbalance"
        Case 6
            Actions.LaunchAction_ByName "adj_curves"
        Case 7
            Actions.LaunchAction_ByName "adj_levels"
        Case 8
            Actions.LaunchAction_ByName "adj_sandh"
        Case 9
            Actions.LaunchAction_ByName "adj_vibrance"
        Case 10
            Actions.LaunchAction_ByName "adj_whitebalance"
        Case Else
            'All remaining commands in this menu are parent menus only
    End Select
End Sub

Private Sub MnuArtistic_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_colorpencil"
        Case 1
            Actions.LaunchAction_ByName "effects_comicbook"
        Case 2
            Actions.LaunchAction_ByName "effects_figuredglass"
        Case 3
            Actions.LaunchAction_ByName "effects_filmnoir"
        Case 4
            Actions.LaunchAction_ByName "effects_glasstiles"
        Case 5
            Actions.LaunchAction_ByName "effects_kaleidoscope"
        Case 6
            Actions.LaunchAction_ByName "effects_modernart"
        Case 7
            Actions.LaunchAction_ByName "effects_oilpainting"
        Case 8
            Actions.LaunchAction_ByName "effects_plasticwrap"
        Case 9
            Actions.LaunchAction_ByName "effects_posterize"
        Case 10
            Actions.LaunchAction_ByName "effects_relief"
        Case 11
            Actions.LaunchAction_ByName "effects_stainedglass"
    End Select
End Sub

Private Sub MnuBatch_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "file_batch_process"
        Case 1
            Actions.LaunchAction_ByName "file_batch_repair"
    End Select
End Sub

Private Sub MnuBlur_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_boxblur"
        Case 1
            Actions.LaunchAction_ByName "effects_gaussianblur"
        Case 2
            Actions.LaunchAction_ByName "effects_surfaceblur"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "effects_motionblur"
        Case 5
            Actions.LaunchAction_ByName "effects_radialblur"
        Case 6
            Actions.LaunchAction_ByName "effects_zoomblur"
    End Select
End Sub

Private Sub MnuColor_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_colorbalance"
        Case 1
            Actions.LaunchAction_ByName "adj_whitebalance"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "adj_hsl"
        Case 4
            Actions.LaunchAction_ByName "adj_temperature"
        Case 5
            Actions.LaunchAction_ByName "adj_tint"
        Case 6
            Actions.LaunchAction_ByName "adj_vibrance"
        Case 7
            '(separator)
        Case 8
            Actions.LaunchAction_ByName "adj_blackandwhite"
        Case 9
            Actions.LaunchAction_ByName "adj_colorlookup"
        Case 10
            Actions.LaunchAction_ByName "adj_colorize"
        Case 11
            Actions.LaunchAction_ByName "adj_photofilters"
        Case 12
            Actions.LaunchAction_ByName "adj_replacecolor"
        Case 13
            Actions.LaunchAction_ByName "adj_sepia"
        Case 14
            Actions.LaunchAction_ByName "adj_splittone"
    End Select
End Sub

Private Sub MnuChannels_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_channelmixer"
        Case 1
            Actions.LaunchAction_ByName "adj_rechannel"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "adj_maxchannel"
        Case 4
            Actions.LaunchAction_ByName "adj_minchannel"
        Case 5
            '(separator)
        Case 6
            Actions.LaunchAction_ByName "adj_shiftchannelsleft"
        Case 7
            Actions.LaunchAction_ByName "adj_shiftchannelsright"
    End Select
End Sub

Private Sub MnuClearRecentMacros_Click()
    g_RecentMacros.MRU_ClearList
End Sub

'The Developer Tools menu is automatically hidden in production builds, so (obviously)
' DO NOT PUT ANYTHING HERE that end-users should be able to access.
Private Sub MnuDevelopers_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "tools_viewdebuglog"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "tools_themeeditor"
        Case 3
            Actions.LaunchAction_ByName "tools_themepackage"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "tools_standalonepackage"
    End Select
End Sub

Private Sub MnuDistort_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_fixlensdistort"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "effects_donut"
        Case 3
            Actions.LaunchAction_ByName "effects_droste"
        Case 4
            Actions.LaunchAction_ByName "effects_lens"
        Case 5
            Actions.LaunchAction_ByName "effects_pinchandwhirl"
        Case 6
            Actions.LaunchAction_ByName "effects_poke"
        Case 7
            Actions.LaunchAction_ByName "effects_ripple"
        Case 8
            Actions.LaunchAction_ByName "effects_squish"
        Case 9
            Actions.LaunchAction_ByName "effects_swirl"
        Case 10
            Actions.LaunchAction_ByName "effects_waves"
        Case 11
            '(separator)
        Case 12
            Actions.LaunchAction_ByName "effects_miscdistort"
    End Select
End Sub

Private Sub MnuEdge_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_emboss"
        Case 1
            Actions.LaunchAction_ByName "effects_enhanceedges"
        Case 2
            Actions.LaunchAction_ByName "effects_findedges"
        Case 3
            Actions.LaunchAction_ByName "effects_gradientflow"
        Case 4
            Actions.LaunchAction_ByName "effects_rangefilter"
        Case 5
            Actions.LaunchAction_ByName "effects_tracecontour"
    End Select
End Sub

Private Sub MnuEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "edit_undo"
        Case 1
            Actions.LaunchAction_ByName "edit_redo"
        Case 2
            Actions.LaunchAction_ByName "edit_history"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "edit_repeat"
        Case 5
            Actions.LaunchAction_ByName "edit_fade"
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "edit_cutlayer"
        Case 8
            Actions.LaunchAction_ByName "edit_cutmerged"
        Case 9
            Actions.LaunchAction_ByName "edit_copylayer"
        Case 10
            Actions.LaunchAction_ByName "edit_copymerged"
        Case 11
            Actions.LaunchAction_ByName "edit_pasteaslayer"
        Case 12
            Actions.LaunchAction_ByName "edit_pasteasimage"
        Case 13
            'Top-level "cut/copy/paste special"
        Case 14
            '(separator)
        Case 15
            Actions.LaunchAction_ByName "edit_clear"
        Case 16
            Actions.LaunchAction_ByName "edit_contentawarefill"
        Case 17
            Actions.LaunchAction_ByName "edit_fill"
        Case 18
            Actions.LaunchAction_ByName "edit_stroke"
    End Select
End Sub

Private Sub MnuEditSpecial_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "edit_specialcut"
        Case 1
            Actions.LaunchAction_ByName "edit_specialcopy"
        Case 2
            'TODO
            'Actions.LaunchAction_ByName "edit_specialpaste"
        Case 3
            'separator
        Case 4
            Actions.LaunchAction_ByName "edit_emptyclipboard"
    End Select
End Sub

Private Sub MnuEffectAnimation_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_animation_background"
        Case 1
            Actions.LaunchAction_ByName "effects_animation_foreground"
        Case 2
            Actions.LaunchAction_ByName "effects_animation_speed"
    End Select
End Sub

Private Sub MnuEffectTransform_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_panandzoom"
        Case 1
            Actions.LaunchAction_ByName "effects_perspective"
        Case 2
            Actions.LaunchAction_ByName "effects_polarconversion"
        Case 3
            Actions.LaunchAction_ByName "effects_rotate"
        Case 4
            Actions.LaunchAction_ByName "effects_shear"
        Case 5
            Actions.LaunchAction_ByName "effects_spherize"
    End Select
End Sub

Private Sub MnuEffectUpper_Click(Index As Integer)
    Select Case Index
        Case 0
            'Artistic
        Case 1
            'Blur
        Case 2
            'Distort
        Case 3
            'Edge
        Case 4
            'Light and Shadow
        Case 5
            'Natural
        Case 6
            'Noise
        Case 7
            'Pixelate
        Case 8
            'Render
        Case 9
            'Sharpen
        Case 10
            'Stylize
        Case 11
            'Transform
        Case 12
            '(separator)
        Case 13
            'Animation
        Case 14
            Actions.LaunchAction_ByName "effects_customfilter"
        Case 15
            Actions.LaunchAction_ByName "effects_8bf"
    End Select
End Sub

Private Sub MnuFile_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "file_new"
        Case 1
            Actions.LaunchAction_ByName "file_open"
        Case 2
            '<Open Recent top-level>
        Case 3
            '<Import top-level>
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "file_close"
        Case 6
            Actions.LaunchAction_ByName "file_closeall"
        Case 7
            '(separator)
        Case 8
            Actions.LaunchAction_ByName "file_save"
        Case 9
            Actions.LaunchAction_ByName "file_savecopy"
        Case 10
            Actions.LaunchAction_ByName "file_saveas"
        Case 11
            Actions.LaunchAction_ByName "file_revert"
        Case 12
            'Export top-level
        Case 13
            '(separator)
        Case 14
            'Batch top-level
        Case 15
            '(separator)
        Case 16
            Actions.LaunchAction_ByName "file_print"
        Case 17
            '(separator)
        Case 18
            Actions.LaunchAction_ByName "file_quit"
    End Select
End Sub

Private Sub MnuFileExport_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "file_export_image"
        Case 1
            Actions.LaunchAction_ByName "file_export_layers"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "file_export_animation"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "file_export_colorlookup"
        Case 6
            Actions.LaunchAction_ByName "file_export_colorprofile"
        Case 7
            Actions.LaunchAction_ByName "file_export_palette"
    End Select
End Sub

Private Sub MnuFileImport_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "file_import_paste"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "file_import_scanner"
        Case 3
            Actions.LaunchAction_ByName "file_import_selectscanner"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "file_import_web"
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "file_import_screenshot"
    End Select
End Sub

Private Sub MnuHelp_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "help_patreon"
        Case 1
            Actions.LaunchAction_ByName "help_donate"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "help_forum"
        Case 4
            Actions.LaunchAction_ByName "help_checkupdates"
        Case 5
            Actions.LaunchAction_ByName "help_reportbug"
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "help_forum"
        Case 8
            Actions.LaunchAction_ByName "help_license"
        Case 9
            Actions.LaunchAction_ByName "help_sourcecode"
        Case 10
            Actions.LaunchAction_ByName "help_website"
        Case 11
            '(separator)
        Case 12
            Actions.LaunchAction_ByName "help_3rdpartylibs"
        Case 13
            '(separator)
        Case 14
            Actions.LaunchAction_ByName "help_about"
    End Select
End Sub

Private Sub MnuHistogram_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_histogramdisplay"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "adj_histogramequalize"
        Case 3
            Actions.LaunchAction_ByName "adj_histogramstretch"
    End Select
End Sub

Private Sub MnuImage_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "image_duplicate"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "image_resize"
        Case 3
            Actions.LaunchAction_ByName "image_contentawareresize"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "image_canvassize"
        Case 6
            Actions.LaunchAction_ByName "image_fittolayer"
        Case 7
            Actions.LaunchAction_ByName "image_fitalllayers"
        Case 8
            '(separator)
        Case 9
            Actions.LaunchAction_ByName "image_crop"
        Case 10
            Actions.LaunchAction_ByName "image_trim"
        Case 11
            '(separator)
        Case 12
            'top-level rotate
        Case 13
            Actions.LaunchAction_ByName "image_fliphorizontal"
        Case 14
            Actions.LaunchAction_ByName "image_flipvertical"
        Case 15
            '(separator)
        Case 16
            Actions.LaunchAction_ByName "image_mergevisible"
        Case 17
            Actions.LaunchAction_ByName "image_flatten"
        Case 18
            '(separator)
        Case 19
            Actions.LaunchAction_ByName "image_animation"
        Case 20
            'Compare top-level
        Case 21
            'Metadata top-level
        Case 22
            Actions.LaunchAction_ByName "image_showinexplorer"
    End Select
End Sub

Private Sub MnuImageCompare_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "image_createlut"
        Case 1
            Actions.LaunchAction_ByName "image_similarity"
    End Select
End Sub

Private Sub MnuInvert_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_invertcmyk"
        Case 1
            Actions.LaunchAction_ByName "adj_inverthue"
        Case 2
            Actions.LaunchAction_ByName "adj_invertrgb"
    End Select
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    
    'Because loading a language can take some time, display a wait screen to discourage interactions
    DisplayWaitScreen g_Language.TranslateMessage("Please wait while the new language is applied..."), Me
    
    'Remove the existing translation from any visible windows
    Message "Removing existing translation..."
    g_Language.UndoTranslations FormMain
    g_Language.UndoTranslations toolbar_Toolbox
    g_Language.UndoTranslations toolbar_Options
    g_Language.UndoTranslations toolbar_Layers
    
    'That may have taken a second or two, so display the reverted text so the user knows what's happening
    DoEvents
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.ActivateNewLanguage Index
    g_Language.ApplyLanguage True, True
    
    Message "Language changed successfully."
    
    HideWaitScreen
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub MnuLayer_Click(Index As Integer)
    Select Case Index
        Case 0
            'Add submenu
        Case 1
            'Delete submenu
        Case 2
            'Replace submenu
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "layer_mergeup"
        Case 5
            Actions.LaunchAction_ByName "layer_mergedown"
        Case 6
            'Order submenu
        Case 7
            'Visibility submenu
        Case 8
            '(separator)
        Case 9
            'Crop submenu
        Case 10
            'Orientation submenu
        Case 11
            'Size submenu
        Case 12
            '(separator)
        Case 13
            'Transparency submenu
        Case 14
            '(separator)
        Case 15
            'Rasterize submenu
    End Select
End Sub

Private Sub MnuLayerDelete_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_deletecurrent"
        Case 1
            Actions.LaunchAction_ByName "layer_deletehidden"
    End Select
End Sub

Private Sub MnuLayerNew_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_addbasic"
        Case 1
            Actions.LaunchAction_ByName "layer_addblank"
        Case 2
            Actions.LaunchAction_ByName "layer_duplicate"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "layer_addfromclipboard"
        Case 5
            Actions.LaunchAction_ByName "layer_addfromfile"
        Case 6
            Actions.LaunchAction_ByName "layer_addfromvisiblelayers"
        Case 7
            '(separator)
        Case 8
            Actions.LaunchAction_ByName "layer_addviacopy"
        Case 9
            Actions.LaunchAction_ByName "layer_addviacut"
    End Select
End Sub

Private Sub MnuLayerCrop_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_cropselection"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "layer_pad"
        Case 3
            Actions.LaunchAction_ByName "layer_trim"
    End Select
End Sub

Private Sub MnuLayerOrder_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_gotop"
        Case 1
            Actions.LaunchAction_ByName "layer_goup"
        Case 2
            Actions.LaunchAction_ByName "layer_godown"
        Case 3
            Actions.LaunchAction_ByName "layer_gobottom"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "layer_movetop"
        Case 6
            Actions.LaunchAction_ByName "layer_moveup"
        Case 7
            Actions.LaunchAction_ByName "layer_movedown"
        Case 8
            Actions.LaunchAction_ByName "layer_movebottom"
        Case 9
            '(separator)
        Case 10
            Actions.LaunchAction_ByName "layer_reverse"
    End Select
End Sub

Private Sub MnuLayerOrientation_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_straighten"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "layer_rotate90"
        Case 3
            Actions.LaunchAction_ByName "layer_rotate270"
        Case 4
            Actions.LaunchAction_ByName "layer_rotate180"
        Case 5
            Actions.LaunchAction_ByName "layer_rotatearbitrary"
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "layer_fliphorizontal"
        Case 8
            Actions.LaunchAction_ByName "layer_flipvertical"
    End Select
End Sub

Private Sub MnuLayerRasterize_Click(Index As Integer)
     Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_rasterizecurrent"
        Case 1
            Actions.LaunchAction_ByName "layer_rasterizeall"
    End Select
End Sub

Private Sub MnuLayerReplace_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_replacefromclipboard"
        Case 1
            Actions.LaunchAction_ByName "layer_replacefromfile"
        Case 2
            Actions.LaunchAction_ByName "layer_replacefromvisiblelayers"
    End Select
End Sub

Private Sub MnuLayerSize_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_resetsize"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "layer_resize"
        Case 3
            Actions.LaunchAction_ByName "layer_contentawareresize"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "layer_fittoimage"
    End Select
End Sub

Private Sub MnuLayerSplit_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_splitlayertoimage"
        Case 1
            Actions.LaunchAction_ByName "layer_splitalllayerstoimages"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "layer_splitimagestolayers"
    End Select
End Sub

Private Sub MnuLayerTransparency_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_colortoalpha"
        Case 1
            Actions.LaunchAction_ByName "layer_luminancetoalpha"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "layer_removealpha"
        Case 4
            Actions.LaunchAction_ByName "layer_thresholdalpha"
    End Select
End Sub

Private Sub MnuLayerVisibility_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "layer_show"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "layer_showonly"
        Case 3
            Actions.LaunchAction_ByName "layer_hideonly"
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "layer_showall"
        Case 6
            Actions.LaunchAction_ByName "layer_hideall"
    End Select
End Sub

Private Sub MnuLighting_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_bandc"
        Case 1
            Actions.LaunchAction_ByName "adj_curves"
        Case 2
            Actions.LaunchAction_ByName "adj_dehaze"
        Case 3
            Actions.LaunchAction_ByName "adj_exposure"
        Case 4
            Actions.LaunchAction_ByName "adj_gamma"
        Case 5
            Actions.LaunchAction_ByName "adj_hdr"
        Case 6
            Actions.LaunchAction_ByName "adj_levels"
        Case 7
            Actions.LaunchAction_ByName "adj_sandh"
    End Select
End Sub

Private Sub MnuLightShadow_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_blacklight"
        Case 1
            Actions.LaunchAction_ByName "effects_bumpmap"
        Case 2
            Actions.LaunchAction_ByName "effects_crossscreen"
        Case 3
            Actions.LaunchAction_ByName "effects_rainbow"
        Case 4
            Actions.LaunchAction_ByName "effects_sunshine"
        Case 5
            '(separator)
        Case 6
            Actions.LaunchAction_ByName "effects_dilate"
        Case 7
            Actions.LaunchAction_ByName "effects_erode"
    End Select
End Sub

Private Sub MnuMacroCreate_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "tools_macrofromhistory"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "tools_recordmacro"
        Case 3
            Actions.LaunchAction_ByName "tools_stopmacro"
    End Select
End Sub

Private Sub MnuMap_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_gradientmap"
        Case 1
            Actions.LaunchAction_ByName "adj_palettemap"
    End Select
End Sub

Private Sub MnuMetadata_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "image_editmetadata"
        Case 1
            Actions.LaunchAction_ByName "image_removemetadata"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "image_countcolors"
        Case 4
            Actions.LaunchAction_ByName "image_maplocation"
    End Select
    
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "adj_colortomonochrome"
        Case 1
            Actions.LaunchAction_ByName "adj_monochrometogray"
    End Select
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_atmosphere"
        Case 1
            Actions.LaunchAction_ByName "effects_fog"
        Case 2
            Actions.LaunchAction_ByName "effects_ignite"
        Case 3
            Actions.LaunchAction_ByName "effects_lava"
        Case 4
            Actions.LaunchAction_ByName "effects_metal"
        Case 5
            Actions.LaunchAction_ByName "effects_snow"
        Case 6
            Actions.LaunchAction_ByName "effects_underwater"
    End Select
End Sub

Private Sub MnuNoise_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_filmgrain"
        Case 1
            Actions.LaunchAction_ByName "effects_rgbnoise"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "effects_anisotropic"
        Case 4
            Actions.LaunchAction_ByName "effects_dustandscratches"
        Case 5
            Actions.LaunchAction_ByName "effects_harmonicmean"
        Case 6
            Actions.LaunchAction_ByName "effects_meanshift"
        Case 7
            Actions.LaunchAction_ByName "effects_median"
        Case 8
            Actions.LaunchAction_ByName "effects_snn"
    End Select
End Sub

Private Sub MnuPixelate_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_colorhalftone"
        Case 1
            Actions.LaunchAction_ByName "effects_crystallize"
        Case 2
            Actions.LaunchAction_ByName "effects_fragment"
        Case 3
            Actions.LaunchAction_ByName "effects_mezzotint"
        Case 4
            Actions.LaunchAction_ByName "effects_mosaic"
        Case 5
            Actions.LaunchAction_ByName "effects_pointillize"
    End Select
End Sub

Public Sub MnuRecentFileList_Click(Index As Integer)
    Actions.LaunchAction_ByName COMMAND_FILE_OPEN_RECENT & Trim$(Str$(Index))
End Sub

Private Sub MnuRecentFiles_Click(Index As Integer)
    Select Case Index
        Case 0
            '(separator)
        Case 1
            Actions.LaunchAction_ByName "file_open_allrecent"
        Case 2
            Actions.LaunchAction_ByName "file_open_clearrecent"
    End Select
End Sub

Private Sub MnuRecentMacros_Click(Index As Integer)
    Actions.LaunchAction_ByName COMMAND_TOOLS_MACRO_RECENT & Trim$(Str$(Index))
End Sub

Private Sub MnuRender_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_clouds"
        Case 1
            Actions.LaunchAction_ByName "effects_fibers"
        Case 2
            Actions.LaunchAction_ByName "effects_truchet"
    End Select
End Sub

Private Sub MnuRotate_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "image_straighten"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "image_rotate90"
        Case 3
            Actions.LaunchAction_ByName "image_rotate270"
        Case 4
            Actions.LaunchAction_ByName "image_rotate180"
        Case 5
            Actions.LaunchAction_ByName "image_rotatearbitrary"
    End Select
End Sub

Private Sub MnuSelect_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "select_all"
        Case 1
            Actions.LaunchAction_ByName "select_none"
        Case 2
            Actions.LaunchAction_ByName "select_invert"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "select_grow"
        Case 5
            Actions.LaunchAction_ByName "select_shrink"
        Case 6
            Actions.LaunchAction_ByName "select_border"
        Case 7
            Actions.LaunchAction_ByName "select_feather"
        Case 8
            Actions.LaunchAction_ByName "select_sharpen"
        Case 9
            '(separator)
        Case 10
            Actions.LaunchAction_ByName "select_erasearea"
        Case 11
            Actions.LaunchAction_ByName "select_fill"
        Case 12
            Actions.LaunchAction_ByName "select_heal"
        Case 13
            Actions.LaunchAction_ByName "select_stroke"
        Case 14
            '(separator)
        Case 15
            Actions.LaunchAction_ByName "select_load"
        Case 16
            Actions.LaunchAction_ByName "select_save"
        Case 17
            'Top-level "Export selection as..." menu
    End Select
End Sub

Private Sub MnuSelectExport_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "select_exportarea"
        Case 1
            Actions.LaunchAction_ByName "select_exportmask"
    End Select
End Sub

Private Sub MnuSharpen_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_sharpen"
        Case 1
            Actions.LaunchAction_ByName "effects_unsharp"
    End Select
End Sub

Private Sub MnuShow_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "show_layeredges"
        Case 1
            Actions.LaunchAction_ByName "show_smartguides"
    End Select
End Sub

Private Sub MnuSnap_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "snap_canvasedge"
        Case 1
            Actions.LaunchAction_ByName "snap_centerline"
        Case 2
            Actions.LaunchAction_ByName "snap_layer"
        Case 3
            'separator
        Case 4
            Actions.LaunchAction_ByName "snap_angle_90"
        Case 5
            Actions.LaunchAction_ByName "snap_angle_45"
        Case 6
            Actions.LaunchAction_ByName "snap_angle_30"
    End Select
End Sub

Private Sub MnuSpecificZoom_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "zoom_16_1"
        Case 1
            Actions.LaunchAction_ByName "zoom_8_1"
        Case 2
            Actions.LaunchAction_ByName "zoom_4_1"
        Case 3
            Actions.LaunchAction_ByName "zoom_2_1"
        Case 4
            Actions.LaunchAction_ByName "zoom_actual"
        Case 5
            Actions.LaunchAction_ByName "zoom_1_2"
        Case 6
            Actions.LaunchAction_ByName "zoom_1_4"
        Case 7
            Actions.LaunchAction_ByName "zoom_1_8"
        Case 8
            Actions.LaunchAction_ByName "zoom_1_16"
    End Select
End Sub

Private Sub MnuStylize_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "effects_antique"
        Case 1
            Actions.LaunchAction_ByName "effects_diffuse"
        Case 2
            Actions.LaunchAction_ByName "effects_kuwahara"
        Case 3
            Actions.LaunchAction_ByName "effects_outline"
        Case 4
            Actions.LaunchAction_ByName "effects_palette"
        Case 5
            Actions.LaunchAction_ByName "effects_portraitglow"
        Case 6
            Actions.LaunchAction_ByName "effects_solarize"
        Case 7
            Actions.LaunchAction_ByName "effects_twins"
        Case 8
            Actions.LaunchAction_ByName "effects_vignetting"
    End Select
End Sub

Private Sub mnuTool_Click(Index As Integer)
    Select Case Index
        Case 0
            'Language top-level
        Case 1
            Actions.LaunchAction_ByName "tools_languageeditor"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "tools_theme"
        Case 4
            '(separator)
        Case 5
            'Create macro top-level
        Case 6
            Actions.LaunchAction_ByName "tools_playmacro"
        Case 7
            'Recent macros top-level
        Case 8
            '(separator)
        Case 9
            Actions.LaunchAction_ByName "tools_screenrecord"
        Case 10
            '(separator)
        Case 11
            Actions.LaunchAction_ByName "tools_hotkeys"
        Case 12
            Actions.LaunchAction_ByName "tools_options"
        Case 13
            '(separator)
        Case 14
            'Developer menu top-level
    End Select
End Sub

Private Sub MnuView_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "view_fit"
        Case 1
            Actions.LaunchAction_ByName "view_center_on_screen"
        Case 2
            '(separator)
        Case 3
            Actions.LaunchAction_ByName "view_zoomin"
        Case 4
            Actions.LaunchAction_ByName "view_zoomout"
        Case 5
            'zoom-to-value top-level
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "view_rulers"
        Case 8
            Actions.LaunchAction_ByName "view_statusbar"
        Case 9
            'show extras top-level
        Case 10
            '(separator)
        Case 11
            Actions.LaunchAction_ByName "snap_global"
        Case 12
            'snap-to top-level
    End Select
End Sub

Private Sub MnuWindow_Click(Index As Integer)
    Select Case Index
        Case 0
            'Toolbox top-level
        Case 1
            Actions.LaunchAction_ByName "window_tooloptions"
        Case 2
            Actions.LaunchAction_ByName "window_layers"
        Case 3
            'Tab-strip top level
        Case 4
            '(separator)
        Case 5
            Actions.LaunchAction_ByName "window_resetsettings"
        Case 6
            '(separator)
        Case 7
            Actions.LaunchAction_ByName "window_next"
        Case 8
            Actions.LaunchAction_ByName "window_previous"
    End Select
End Sub

Private Sub MnuWindowOpen_Click(Index As Integer)

    'Open the current document corresponding to the index in the menu
    Dim listOfOpenImages As pdStack
    PDImages.GetListOfActiveImageIDs listOfOpenImages
    
    If (Index < listOfOpenImages.GetNumOfInts) Then
        If PDImages.IsImageActive(listOfOpenImages.GetInt(Index)) Then CanvasManager.ActivatePDImage listOfOpenImages.GetInt(Index), "window menu"
    End If

End Sub

Private Sub MnuWindowTabstrip_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "window_imagetabstrip_alwaysshow"
        Case 1
            Actions.LaunchAction_ByName "window_imagetabstrip_shownormal"
        Case 2
            Actions.LaunchAction_ByName "window_imagetabstrip_nevershow"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "window_imagetabstrip_alignleft"
        Case 5
            Actions.LaunchAction_ByName "window_imagetabstrip_aligntop"
        Case 6
            Actions.LaunchAction_ByName "window_imagetabstrip_alignright"
        Case 7
            Actions.LaunchAction_ByName "window_imagetabstrip_alignbottom"
    End Select
End Sub

Private Sub MnuWindowToolbox_Click(Index As Integer)
    Select Case Index
        Case 0
            Actions.LaunchAction_ByName "window_displaytoolbox"
        Case 1
            '(separator)
        Case 2
            Actions.LaunchAction_ByName "window_displaytoolcategories"
        Case 3
            '(separator)
        Case 4
            Actions.LaunchAction_ByName "window_smalltoolbuttons"
        Case 5
            Actions.LaunchAction_ByName "window_mediumtoolbuttons"
        Case 6
            Actions.LaunchAction_ByName "window_largetoolbuttons"
    End Select
End Sub

'Test TEST menu is a special, developer-only menu for easy access to whatever feature you're currently testing.
' Do NOT expose this menu in public builds.
Private Sub MnuTest_Click()
    
    On Error GoTo StopTestImmediately
    
    'Typically, you'll want to ensure an image is loaded...
    If (PDImages.GetNumOpenImages <= 0) Then Exit Sub
    
    'Perf testing can be initialized like this...
    Dim startTime As Currency, lastTime As Currency
    VBHacks.GetHighResTime startTime
    lastTime = startTime
    
'    'Test code goes here
'    PDDebug.LogAction "Convert to HDR..."
'    Dim tmpSurface As pdSurfaceF
'    Set tmpSurface = New pdSurfaceF
'    PDDebug.LogAction tmpSurface.CreateFromPDDib(PDImages.GetActiveImage.GetActiveDIB)
'    PDDebug.LogAction "Alpha status: " & tmpSurface.GetAlphaPremultiplication
'    PDDebug.LogAction VBHacks.GetTimeDiffNowAsString(lastTime)
'    VBHacks.GetHighResTime lastTime
'
'    PDDebug.LogAction "Downsample..."
'    Dim rsSurface As pdSurfaceF
'    Set rsSurface = New pdSurfaceF
'    tmpSurface.SetAlphaPremultiplication False
'    PDDebug.LogAction Resampling.ResampleImageF(rsSurface, tmpSurface, tmpSurface.GetWidth / 10, tmpSurface.GetHeight / 10, rf_Box, True)
'    PDDebug.LogAction "Alpha status: " & tmpSurface.GetAlphaPremultiplication
'    PDDebug.LogAction VBHacks.GetTimeDiffNowAsString(lastTime)
'    VBHacks.GetHighResTime lastTime
'
'    PDDebug.LogAction "Upsample..."
'    tmpSurface.Reset
'    PDDebug.LogAction Resampling.ResampleImageF(tmpSurface, rsSurface, tmpSurface.GetWidth, tmpSurface.GetHeight, rf_Lanczos, True)
'    tmpSurface.SetAlphaPremultiplication True
'    ProgressBars.ReleaseProgressBar
'    PDDebug.LogAction "Alpha status: " & tmpSurface.GetAlphaPremultiplication
'    PDDebug.LogAction VBHacks.GetTimeDiffNowAsString(lastTime)
'    VBHacks.GetHighResTime lastTime
'
'    PDDebug.LogAction "Convert back to SDR..."
'    Dim tmpDIB As pdDIB
'    PDDebug.LogAction tmpSurface.ConvertToPDDib(tmpDIB)
'    PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.CreateFromExistingDIB tmpDIB
'    PDDebug.LogAction "Alpha status: " & PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetAlphaPremultiplication
'    PDDebug.LogAction VBHacks.GetTimeDiffNowAsString(lastTime)
'    VBHacks.GetHighResTime lastTime
'
'    PDDebug.LogAction "Done."
'    Set tmpSurface = Nothing
'
'    'Want to display the test results?  Copy the processed image into PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB,
'    ' then uncomment these two lines:
'    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
'    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage, FormMain.MainCanvas(0)
    
    'Want to test a new dialog?  Call it here, using a line like the following:
    'ShowPDDialog vbModal, FormToTest
    
    'Report timing results:
    PDDebug.LogAction "Test function time: " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    Exit Sub
    
StopTestImmediately:
    PDDebug.LogAction "Error in test sub: " & Err.Number & ", " & Err.Description

End Sub
