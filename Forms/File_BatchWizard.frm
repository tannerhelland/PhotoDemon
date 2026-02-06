VERSION 5.00
Begin VB.Form FormBatchWizard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Batch process"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13200
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   880
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdPrevious 
      Height          =   615
      Left            =   6060
      TabIndex        =   0
      Top             =   7530
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1085
      Caption         =   "Previous"
      Enabled         =   0   'False
   End
   Begin PhotoDemon.pdButton cmdNext 
      Height          =   615
      Left            =   8820
      TabIndex        =   1
      Top             =   7530
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1085
      Caption         =   "Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   11760
      TabIndex        =   2
      Top             =   7530
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   6705
      Index           =   0
      Left            =   120
      Top             =   720
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   11827
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWizardTitle 
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   714
      Caption         =   "Step 1: select the photo editing action(s) to apply to each image"
      FontBold        =   -1  'True
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   2
      Left            =   3300
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdSpinner spnVectorImport 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   45
         Top             =   6180
         Width           =   2055
         _ExtentX        =   5741
         _ExtentY        =   661
         DefaultValue    =   1920
         Min             =   1
         Max             =   32000
         Value           =   1920
      End
      Begin PhotoDemon.pdButtonStrip btsVectorImport 
         Height          =   975
         Left            =   720
         TabIndex        =   44
         Top             =   4800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1720
         Caption         =   "when importing vector images"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   4320
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   661
         Caption         =   "(optional) advanced import settings"
         FontSize        =   12
      End
      Begin PhotoDemon.pdCheckBox chkExportAnimation 
         Height          =   375
         Left            =   720
         TabIndex        =   42
         Top             =   3000
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         Caption         =   "auto-detect animated images"
      End
      Begin PhotoDemon.pdButton cmdExportSettings 
         Height          =   735
         Left            =   720
         TabIndex        =   24
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1296
         Caption         =   "set export settings for this format..."
      End
      Begin PhotoDemon.pdDropDown cmbOutputFormat 
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   1680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdRadioButton optFormat 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   661
         Caption         =   "keep images in their original format"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdRadioButton optFormat 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   661
         Caption         =   "convert all images to a new format"
      End
      Begin PhotoDemon.pdLabel lblExplanationFormat 
         Height          =   600
         Left            =   720
         Top             =   420
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   1058
         Caption         =   ""
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdButton cmdExportSettingsAnimated 
         Height          =   735
         Left            =   720
         TabIndex        =   41
         Top             =   3420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1296
         Caption         =   "set export settings for animated images..."
      End
      Begin PhotoDemon.pdLabel lblVectorImport 
         Height          =   300
         Index           =   0
         Left            =   840
         Top             =   5880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         Caption         =   "width"
         FontSize        =   11
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdSpinner spnVectorImport 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   46
         Top             =   6180
         Width           =   2055
         _ExtentX        =   5741
         _ExtentY        =   661
         DefaultValue    =   1080
         Min             =   1
         Max             =   32000
         Value           =   1080
      End
      Begin PhotoDemon.pdLabel lblVectorImport 
         Height          =   300
         Index           =   1
         Left            =   3240
         Top             =   5880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         Caption         =   "height"
         FontSize        =   11
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   0
      Left            =   3300
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdButtonStrip btsPhotoOps 
         Height          =   975
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   1720
         Caption         =   "apply photo editing actions"
      End
      Begin PhotoDemon.pdLabel lblExplanation 
         Height          =   720
         Index           =   1
         Left            =   240
         Top             =   1200
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   1270
         Caption         =   "if you only want to rename images or change image formats, use this option "
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdContainer picPhotoEdits 
         Height          =   5400
         Left            =   120
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9525
         Begin PhotoDemon.pdPictureBox picResizeDemo 
            Height          =   750
            Left            =   6720
            Top             =   2880
            Width           =   2865
            _ExtentX        =   0
            _ExtentY        =   0
         End
         Begin PhotoDemon.pdDropDown cmbResizeFit 
            Height          =   615
            Left            =   720
            TabIndex        =   33
            Top             =   2850
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1085
            Caption         =   "resize image by"
            FontSizeCaption =   10
         End
         Begin PhotoDemon.pdButton cmdSelectMacro 
            Height          =   615
            Left            =   6960
            TabIndex        =   34
            Top             =   4170
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1085
            Caption         =   "Select macro..."
            FontSize        =   9
         End
         Begin PhotoDemon.pdTextBox txtMacro 
            Height          =   315
            Left            =   600
            TabIndex        =   35
            Top             =   4320
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
            Text            =   "no macro selected"
         End
         Begin PhotoDemon.pdCheckBox chkActions 
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   3840
            Width           =   9540
            _ExtentX        =   16828
            _ExtentY        =   529
            Caption         =   "apply other actions from a saved macro file"
            Value           =   0   'False
         End
         Begin PhotoDemon.pdCheckBox chkActions 
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   10020
            _ExtentX        =   17674
            _ExtentY        =   582
            Caption         =   "resize images"
            Value           =   0   'False
         End
         Begin PhotoDemon.pdCheckBox chkActions 
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Width           =   10020
            _ExtentX        =   17674
            _ExtentY        =   582
            Caption         =   "fix exposure and lighting problems"
            Value           =   0   'False
         End
         Begin PhotoDemon.pdResize ucResize 
            Height          =   1650
            Left            =   360
            TabIndex        =   40
            Top             =   960
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   2910
            UnknownSizeMode =   -1  'True
         End
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   1
      Left            =   3300
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdPictureBox picPreview 
         Height          =   2925
         Left            =   6600
         Top             =   465
         Width           =   3135
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdLabel lblCurrentFile 
         Height          =   285
         Left            =   330
         Top             =   3570
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   503
         Caption         =   ""
         FontSize        =   9
      End
      Begin PhotoDemon.pdCheckBox chkAddSubfoldersToo 
         Height          =   375
         Left            =   225
         TabIndex        =   3
         Top             =   5670
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Caption         =   "include subfolders"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdListBox lstFiles 
         Height          =   3405
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   6006
         Caption         =   "current batch list"
         FontSize        =   9
      End
      Begin PhotoDemon.pdButton cmdSaveList 
         Height          =   615
         Left            =   6960
         TabIndex        =   19
         Top             =   4995
         Width           =   2775
         _ExtentX        =   5318
         _ExtentY        =   1085
         Caption         =   "save list..."
      End
      Begin PhotoDemon.pdButton cmdLoadList 
         Height          =   615
         Left            =   6960
         TabIndex        =   20
         Top             =   4290
         Width           =   2775
         _ExtentX        =   5318
         _ExtentY        =   1085
         Caption         =   "load list..."
      End
      Begin PhotoDemon.pdButton cmdRemoveAll 
         Height          =   615
         Left            =   3600
         TabIndex        =   21
         Top             =   6045
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "erase entire list"
      End
      Begin PhotoDemon.pdButton cmdRemove 
         Height          =   615
         Left            =   3600
         TabIndex        =   22
         Top             =   4290
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "remove selected image"
      End
      Begin PhotoDemon.pdButton cmdAddFiles 
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   4290
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "add individual images..."
      End
      Begin PhotoDemon.pdCheckBox chkEnablePreview 
         Height          =   330
         Left            =   6600
         TabIndex        =   7
         Top             =   0
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   582
         Caption         =   "show image previews"
      End
      Begin PhotoDemon.pdLabel lblFiles 
         Height          =   285
         Left            =   120
         Top             =   3930
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   503
         Caption         =   "add images"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblModify 
         Height          =   285
         Left            =   3480
         Top             =   3930
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   503
         Caption         =   "modify list"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblLoadSaveList 
         Height          =   285
         Left            =   6840
         Top             =   3930
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   503
         Caption         =   "load / save list"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdButton cmdAddFolders 
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   4995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "add entire folder(s)..."
      End
      Begin PhotoDemon.pdButton cmdRemoveFolder 
         Height          =   615
         Left            =   3600
         TabIndex        =   4
         Top             =   4995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "remove all images in this folder"
      End
      Begin PhotoDemon.pdCheckBox chkRemoveSubfolders 
         Height          =   375
         Left            =   3585
         TabIndex        =   5
         Top             =   5670
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "include subfolders"
         Value           =   0   'False
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   4
      Left            =   3300
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdProgressBar pbBatch 
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   873
      End
      Begin PhotoDemon.pdLabel lblBatchProgress 
         Height          =   645
         Left            =   240
         Top             =   2400
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   1138
         Alignment       =   2
         Caption         =   "(batch conversion process will appear here at run-time)"
         ForeColor       =   -2147483640
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblTimeRemaining 
         Height          =   645
         Left            =   240
         Top             =   3840
         Width           =   9435
         _ExtentX        =   0
         _ExtentY        =   0
         Alignment       =   2
         Caption         =   ""
         ForeColor       =   -2147483640
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   3
      Left            =   3300
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdCheckBox chkOutputPreserveFolders 
         Height          =   345
         Left            =   480
         TabIndex        =   43
         Top             =   750
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   609
         Caption         =   "preserve input folder structure"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdDropDown cmbOutputOptions 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1680
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButton cmdSelectOutputPath 
         Height          =   735
         Left            =   6600
         TabIndex        =   26
         Top             =   330
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
         Caption         =   "Select destination folder..."
         FontSize        =   9
      End
      Begin PhotoDemon.pdTextBox txtRenameRemove 
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Top             =   3960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendBack 
         Height          =   315
         Left            =   5640
         TabIndex        =   28
         Top             =   3000
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendFront 
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   3000
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         Text            =   "NEW_"
      End
      Begin PhotoDemon.pdTextBox txtOutputPath 
         Height          =   315
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         Text            =   "C:\"
      End
      Begin PhotoDemon.pdRadioButton optCase 
         Height          =   330
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   4920
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Caption         =   "lowercase"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkRenamePrefix 
         Height          =   330
         Left            =   480
         TabIndex        =   11
         Top             =   2640
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         Caption         =   "add a prefix to each filename:"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkRenameSuffix 
         Height          =   330
         Left            =   5280
         TabIndex        =   12
         Top             =   2640
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   582
         Caption         =   "add a suffix to each filename:"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkRenameRemove 
         Height          =   330
         Left            =   480
         TabIndex        =   13
         Top             =   3600
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   582
         Caption         =   "remove the following text (if found) from each filename:"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkRenameCase 
         Height          =   330
         Left            =   480
         TabIndex        =   14
         Top             =   4560
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   582
         Caption         =   "force each filename, including extension, to the following case:"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdRadioButton optCase 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   16
         Top             =   4920
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         Caption         =   "UPPERCASE"
      End
      Begin PhotoDemon.pdCheckBox chkRenameSpaces 
         Height          =   330
         Left            =   480
         TabIndex        =   17
         Top             =   5520
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   582
         Caption         =   "replace spaces in filenames with underscores"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkRenameCaseSensitive 
         Height          =   330
         Left            =   5280
         TabIndex        =   18
         Top             =   3960
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         Caption         =   "use case-sensitive matching"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdLabel lblDstFilename 
         Height          =   285
         Left            =   120
         Top             =   1320
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   503
         Caption         =   "after images are processed, save them with the following name:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblOptionalText 
         Height          =   285
         Left            =   120
         Top             =   2280
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   503
         Caption         =   "additional rename options"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblDstFolder 
         Height          =   285
         Left            =   120
         Top             =   0
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   503
         Caption         =   "output images to this folder:"
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "FormBatchWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Batch Conversion Form
'Copyright 2007-2026 by Tanner Helland
'Created: 3/Nov/07
'Last updated: 09/December/22
'Last update: add overrides for vector image import size (see https://github.com/tannerhelland/PhotoDemon/issues/456)
'
'PhotoDemon's batch process wizard is one of its most unique features.  It integrates tightly with PD's
' macro recorder, which allows any combination of actions to be applied to any set of images.  Neat stuff!
'
'The current batch wizard is broken into four stages:
'
'1) Select the photo editing operations (if any) that will be applied.  This step is optional;
'    if no photo editing actions are selected, you can still convert images between formats and/or
'    batch rename them, without actually changing their pixel contents.
'
'2) Assemble the list of files to be processed.  The list can be built from any number of files or folders,
'    and several different input methods are supported.
'
'3) Select output file format.  There are two choices: retain original format (with limitations,
'    e.g. read-only formats like manufacturer-specific RAW files will be saved as JPEGs), or export to
'    some new format.  If a new format is selected, PD's standard export dialogs are available to set
'    precise export parameters (e.g. JPEG quality).
'
'4) Choose where the exported images will be saved and how they will be named.  This includes a number
'    of basic renaming options.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Currently active page in the wizard
Private m_CurrentPage As Long

'Has the current list of images been saved?  (We check this if the wizard is exited prematurely, so the user can
' have a chance to save their existing settings.)
Private m_ImageListSaved As Boolean

'The path to the image currently rendered to the "image preview" box.  (We cache this to optimize redraws; if the
' path hasn't changed since the last request, we do not redraw the preview.)
Private m_CurImagePreview As String

'Because these words are used frequently, if we have to translate them every time they're used, it slows down the
' process considerably.  So cache them in advance.
' TODO: fix this, because word order (obviously) is not consistent from language to language
Private m_wordForBatchList As String, m_wordForItem As String, m_wordForItems As String

'We maintain folder paths locally, in case the user wants to add multiple folders in succession
Private m_LastBatchFolder As String

'While we're processing the list (for example, when removing items automatically), we want to ignore any events raised
' by the underlying list UI.
Private m_ListBusy As Boolean

'Export settings were overhauled for 7.0's release.  Batch processing now uses the same export dialogs as PD's regular
' save functions.  To make sure the user actually sets export settings before progressing, we use this tracker.
Private m_ExportSettingsSet As Boolean, m_ExportSettingsFormat As String, m_ExportSettingsMetadata As String

'As of 9.0, formats that support animation can also be batch-processed, and the user can set export settings
' generically, just like they do for static formats.
Private m_ExportSettingsSetAnimation As Boolean, m_ExportSettingsFormatAnimation As String

'This dialog interacts with a lot of file-system bits.  This module-level pdFSO object is initialized at Form_Load(),
' and can be used wherever convenient.
Private m_FSO As pdFSO

Private Sub btsPhotoOps_Click(ByVal buttonIndex As Long)
    UpdatePhotoOpVisibility
End Sub

Private Sub btsVectorImport_Click(ByVal buttonIndex As Long)
    UpdateVectorImportVisibility
End Sub

'Enable/disable previewing the currently selected image.  (This is helpful for camera folders full of names like "DSC1234".)
Private Sub chkEnablePreview_Click()
    
    'If the user is enabling previews, try to display the last item the user selected in the SOURCE list box
    If chkEnablePreview.Value Then
        If (lstFiles.ListIndex >= 0) Then UpdatePreview lstFiles.List(lstFiles.ListIndex), True
        
    'If the user is disabling previews, clear the picture box and display a notice
    Else
        picPreview.PaintText g_Language.TranslateMessage("previews disabled"), 10!, False, True
    End If
    
End Sub

'By default, neither case-related export option is selected.  Default to lowercase when the RenameCase checkbox is used.
Private Sub chkRenameCase_Click()
    If (Not optCase(0).Value) And (Not optCase(1).Value) Then optCase(0).Value = True
End Sub

'Set output image format
Private Sub cmbOutputFormat_Click()
        
    'Automatically switch the option for "use fixed export format" to TRUE
    optFormat(1).Value = True
        
    'If this format doesn't support export settings, hide the "set export settings" button
    If ImageFormats.IsExportDialogSupported(ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)) Then
        m_ExportSettingsSet = False
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
        cmdExportSettings.Visible = True
    Else
        m_ExportSettingsSet = True
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
        cmdExportSettings.Visible = False
    End If
    
    'Same for animation settings
    If ImageFormats.IsAnimationSupported(ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)) Then
        chkExportAnimation.Visible = True
        cmdExportSettingsAnimated.Visible = True
        m_ExportSettingsSetAnimation = False
        m_ExportSettingsFormatAnimation = vbNullString
    Else
        chkExportAnimation.Visible = False
        cmdExportSettingsAnimated.Visible = False
        m_ExportSettingsSetAnimation = False
        m_ExportSettingsFormatAnimation = vbNullString
    End If
    
End Sub

'Show a sample for the unintuitive "how to fit resized image in canvas" option
Private Sub cmbResizeFit_Click()
    UpdateResizeSample
End Sub

Private Sub UpdateResizeSample()
    
    'Display a sample image of the selected resize method
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Load the proper sample image to our temporary DIB
    Select Case cmbResizeFit.ListIndex
    
        'Stretch
        Case 0
            LoadResourceToDIB "sample_resize_stretch", tmpDIB, 191, 50
        
        'Fit inclusive
        Case 1
            LoadResourceToDIB "sample_resize_fitinclusive", tmpDIB, 191, 50
        
        'Fit exclusive
        Case 2
            LoadResourceToDIB "sample_resize_fitexclusive", tmpDIB, 191, 50
    
    End Select
    
    'Paint the sample image to the screen
    picResizeDemo.CopyDIB tmpDIB, suspendTransparencyGrid:=True, useNeutralBackground:=True, drawBorderAroundImage:=True

End Sub

'cmdAddFiles allows the user to move files from the source image list box to the batch list box
Private Sub cmdAddFiles_Click()
    
    Dim listOfFiles As pdStringStack
    If FileMenu.PhotoDemon_OpenImageDialog(listOfFiles, Me.hWnd) Then
        
        lstFiles.SetAutomaticRedraws False
        
        Dim tmpFilename As String
        Do While listOfFiles.PopString(tmpFilename)
            lstFiles.AddItem tmpFilename
        Loop
        
        lstFiles.SetAutomaticRedraws True, True
        
        UpdateBatchListCount
        m_ImageListSaved = False
        
        'Enable the "remove all images" button if at least one image exists in the processing list
        cmdRemoveAll.Enabled = (lstFiles.ListCount > 0)
        cmdSaveList.Enabled = (lstFiles.ListCount > 0)
            
    End If
    
End Sub

'Add entire folders to the current batch list
Private Sub cmdAddFolders_Click()
    
    If (LenB(m_LastBatchFolder) = 0) Then m_LastBatchFolder = UserPrefs.GetPref_String("Paths", "Open Image", vbNullString)
    
    Dim folderPath As String
    folderPath = Files.PathBrowseDialog(Me.hWnd, m_LastBatchFolder)
    
    If (LenB(folderPath) <> 0) Then
        
        Dim listOfFiles As pdStringStack
        If m_FSO.RetrieveAllFiles(folderPath, listOfFiles, chkAddSubfoldersToo.Value, False, ImageFormats.GetListOfInputFormats("|", False)) Then
                
            lstFiles.SetAutomaticRedraws False
            
            Dim tmpFilename As String
            Do While listOfFiles.PopString(tmpFilename)
                lstFiles.AddItem tmpFilename
            Loop
            
            lstFiles.SetAutomaticRedraws True, True
            
            UpdateBatchListCount
            m_ImageListSaved = False
            
            'Enable the "remove all images" button if at least one image exists in the processing list
            cmdRemoveAll.Enabled = (lstFiles.ListCount > 0)
            cmdSaveList.Enabled = (lstFiles.ListCount > 0)
            
            'Save this folder as the last-used folder
            m_LastBatchFolder = folderPath
            UserPrefs.SetPref_String "Paths", "Open Image", folderPath
            
        End If
        
    End If

End Sub

'Cancel and exit the dialog, with optional prompts as necessary (see Form_QueryUnload)
Private Sub cmdCancel_Click()
    
    If (m_CurrentPage = picContainer.Count - 1) Then
        
        If (Macros.GetMacroStatus <> MacroSTOP) Then
            Dim msgReturn As VbMsgBoxResult
            msgReturn = PDMsgBox("Are you sure you want to cancel the current batch process?", vbYesNoCancel Or vbExclamation, "Cancel batch processing")
            If (msgReturn = vbYes) Then Macros.SetMacroStatus MacroCANCEL
            pbBatch.Visible = False
        Else
            Unload Me
        End If
        
    Else
        Unload Me
    End If
    
End Sub

'Are we allowed to exit the dialog?  Some conditions may result in modal UI prompts (e.g. "do you want to save your
' current settings before exiting?")  If the user CANCELS a modal UI dialog, the exit process must be aborted.
Private Function AllowedToExit() As Boolean
    
    AllowedToExit = True
    
    'If the user has created a list of images to process and they attempt to exit without saving the list,
    ' give them a chance to save it.
    If (m_CurrentPage < picContainer.Count - 1) Then
        If (Not m_ImageListSaved) Then
            If (lstFiles.ListCount > 0) Then
            
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("If you exit now, your batch list (the list of images to be processed) will be lost.  By saving your list, you can easily resume this batch operation at a later date." & vbCrLf & vbCrLf & "Would you like to save your batch list before exiting?", vbExclamation Or vbYesNoCancel, "Unsaved changes")
                
                Select Case msgReturn
                    Case vbYes
                        AllowedToExit = SaveCurrentBatchList()
                    Case vbNo
                        AllowedToExit = True
                    Case vbCancel
                        AllowedToExit = False
                End Select
                
            End If
        End If
    End If
    
End Function

'Raise an appropriate settings dialog for the selected export format
Private Sub cmdExportSettings_Click()
    
    'Convert the current dropdown index into a PD format constant
    Dim saveFormat As PD_IMAGE_FORMAT
    saveFormat = ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)
    
    'See if this format even supports dialogs...
    If ImageFormats.IsExportDialogSupported(saveFormat) Then
        
        'The saving module will now raise a dialog specific to the selected format.  If successful, it will fill
        ' the passed settings and metadata strings with XML data describing the user's settings.
        m_ExportSettingsSet = Saving.GetExportParamsFromDialog(Nothing, saveFormat, m_ExportSettingsFormat, m_ExportSettingsMetadata, False)
        
        'If the user cancels the dialog, exit immediately
        If (Not m_ExportSettingsSet) Then
            m_ExportSettingsSet = False
            m_ExportSettingsFormat = vbNullString
            m_ExportSettingsMetadata = vbNullString
        End If
    
    'Formats that do not support export settings do not need to raise a dialog at all
    Else
        m_ExportSettingsSet = True
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
    End If
    
End Sub

'Same as normal export settings, but an animation-centric export dialog is used instead
Private Sub cmdExportSettingsAnimated_Click()

    'Convert the current dropdown index into a PD format constant
    Dim saveFormat As PD_IMAGE_FORMAT
    saveFormat = ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)
    
    'See if this format even supports animation dialogs...
    If ImageFormats.IsExportDialogSupported(saveFormat) And ImageFormats.IsAnimationSupported(saveFormat) Then
        
        'The saving module will now raise a dialog specific to the selected format.  If successful, it will fill
        ' the passed settings and metadata strings with XML data describing the user's settings.
        m_ExportSettingsSetAnimation = Saving.GetExportParamsFromDialog(Nothing, saveFormat, m_ExportSettingsFormatAnimation, m_ExportSettingsMetadata, True)
        
        'If the user cancels the dialog, note it so we can prompt them again if they try to
        ' proceed with batch processing
        If (Not m_ExportSettingsSetAnimation) Then
            m_ExportSettingsSetAnimation = False
            m_ExportSettingsFormatAnimation = vbNullString
        End If
    
    'Formats that do not support export settings do not need to raise a dialog at all
    Else
        m_ExportSettingsSetAnimation = True
        m_ExportSettingsFormatAnimation = vbNullString
    End If
    
End Sub

'Load a list of images (previously saved from within PhotoDemon) into the current batch list
Private Sub cmdLoadList_Click()
    
    Dim sFile As String
    
    'Get the last "open/save image list" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Batch Process", "List Folder", vbNullString)
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("Batch Image List") & " (.pdl)|*.pdl"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Load a list of images")
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, cdTitle, ".pdl", FormBatchWizard.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = Files.FileGetPath(sFile)
        UserPrefs.SetPref_String "Batch Process", "List Folder", listPath
        
        Dim listOK As Boolean
        listOK = False
        
        'Load the file using pdFSO, which is Unicode-compatible
        Dim fileContents As String
        If Files.FileLoadAsString(sFile, fileContents) Then
            
            'Ensure multiple lines (there should be at least 3 - header, 1+ files, footer)
            If (InStr(1, fileContents, vbCrLf, vbBinaryCompare) > 0) Then
            
                'The file was originally delimited by vbCrLf.  Parse it now.
                Dim fileLines() As String
                fileLines = Split(fileContents, vbCrLf, , vbBinaryCompare)
                
                If (UBound(fileLines) > 0) Then
                    
                    'Validate the first line of the file
                    If Strings.StringsEqual(fileLines(0), "<PHOTODEMON BATCH CONVERSION LIST>", True) Then
                        
                        listOK = True
                        
                        'If the user has already created a list of files to process, ask if they want to replace or append
                        ' the loaded entries to their current list.
                        If (lstFiles.ListCount > 0) Then
                            Dim msgReturn As VbMsgBoxResult
                            msgReturn = PDMsgBox("You have already created a list of images for processing.  The list of images inside this file will be appended to the bottom of your current list.", vbOKCancel Or vbInformation, "Batch process notification")
                            If (msgReturn = vbCancel) Then Exit Sub
                        End If
                                    
                        Screen.MousePointer = vbHourglass
                    
                        'Now that everything is in place, load the entries from the previously saved file
                        Dim numOfEntries As Long
                        numOfEntries = CLng(fileLines(1))
                        
                        lstFiles.SetAutomaticRedraws False
                        
                        Dim i As Long
                        For i = 2 To numOfEntries + 1
                            If Files.FileExists(fileLines(i)) Then lstFiles.AddItem fileLines(i)
                        Next i
                        
                        lstFiles.SetAutomaticRedraws True, True
                        
                        'Note that the current list has NOT been saved
                        m_ImageListSaved = False
            
                        'Enable the "remove all images" button if at least one image exists in the processing list
                        If (lstFiles.ListCount > 0) Then
                            If (Not cmdRemoveAll.Enabled) Then cmdRemoveAll.Enabled = True
                            If (Not cmdSaveList.Enabled) Then cmdSaveList.Enabled = True
                        End If
                        
                        UpdateBatchListCount
                        
                        Screen.MousePointer = vbDefault
                        
                    '/bad header
                    End If
                
                '/no lines
                End If
            
            '/no vbCrLf
            End If
        
        '/failed initial load
        End If
        
        If (Not listOK) Then
            PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation Or vbOKOnly, "Invalid list file"
            Exit Sub
        End If
        
        'Note that the current list has been saved (technically it hasn't, I realize,
        ' but it exists in a file in its current state so close enough!)
        m_ImageListSaved = True
        
    End If
    
End Sub

Private Sub cmdNext_Click()
    ChangeBatchPage True
End Sub

Private Sub cmdPrevious_Click()
    ChangeBatchPage False
End Sub

'This function is used to advance (TRUE) or retreat (FALSE) the active wizard panel
Private Sub ChangeBatchPage(ByVal moveForward As Boolean)
    
    'Before doing anything else, see if the user is on the final step.  If they are, initiate the batch conversion.
    If moveForward And (m_CurrentPage = picContainer.Count - 2) Then
        m_CurrentPage = picContainer.Count - 1
        UpdateWizardText
        PrepareForBatchConversion
        Exit Sub
    End If
    
    'Before moving to the next page, validate the current one
    Select Case m_CurrentPage
    
        'Select photo editing options
        Case 0
        
            'If the user is not applying any photo editing actions, skip to the next step.  If the user IS applying photo editing
            ' actions, additional validations must be applied.
            If (btsPhotoOps.ListIndex = 1) Then
            
                'If the user wants to resize the image, make sure the width and height values are valid
                If chkActions(1).Value Then
                    If Not ucResize.IsValid(True) Then Exit Sub
                End If
                
                'If the user wants us to apply a macro, ensure that the macro text box has a macro file specified
                If chkActions(2).Value And (Strings.StringsEqual(txtMacro.Text, g_Language.TranslateMessage("no macro selected")) Or (LenB(txtMacro.Text) = 0)) Then
                    PDMsgBox "You have requested that a macro be applied to each image, but no macro file has been selected.  Please select a valid macro file.", vbExclamation Or vbOKOnly, "No macro file selected"
                    txtMacro.SelectAll
                    Exit Sub
                End If
                
            End If
            
        'Add images to batch list
        Case 1
        
            'If no images have been added to the batch list, make the user add some!
            If (moveForward And (lstFiles.ListCount = 0)) Then
                PDMsgBox "You have not selected any images to process!  Please add one or more images to the batch list.", vbExclamation Or vbOKOnly, "No images selected"
                Exit Sub
            End If
        
        'Select output format
        Case 2
            
            'If the user has asked us to convert all images to a new format, make sure they clicked the
            ' "set export options" button (to define what export settings we'll use).  There are technically
            ' two states to check here - regular format settings, and if the format supports animations and
            ' the "auto-detect animated images" checkbox is set, animation settings too.
            
            ' contains all of the user's selected image format options (JPEG quality, etc)
            If (optFormat(1).Value And moveForward) Then
                
                Dim showWarning As Boolean, showWarningAnimated As Boolean
                showWarning = (Not m_ExportSettingsSet)
                If chkExportAnimation.Visible Then
                    showWarningAnimated = chkExportAnimation.Value And (Not m_ExportSettingsSetAnimation)
                Else
                    showWarningAnimated = True
                End If
                
                'If the user clicks one box but not the other, that's okay - they probably only care about
                ' that particular type of image.  But if they haven't clicked *either* box, warn them.
                If (showWarning And showWarningAnimated) Then
                    PDMsgBox "Before proceeding, you need to click the ""set export settings for this format"" button to specify what export settings you want to use.", vbExclamation Or vbOKOnly, "Export settings required"
                    Exit Sub
                End If
                
            End If
        
        'Select output directory and file name
        Case 3
            
            'Make sure we have write access to the output folder.  If we don't, cancel and warn the user.
            If (Not Files.PathExists(txtOutputPath)) Then
                
                If (Not Files.PathCreate(txtOutputPath)) Then
                    PDMsgBox "PhotoDemon cannot access the requested output folder.  Please select a non-system, unrestricted folder for the batch process.", vbExclamation Or vbOKOnly, "Folder access unavailable"
                    txtOutputPath.SelectAll
                    Exit Sub
                End If
                
            End If
    
    End Select

    'True means move forward; false means move backward
    If moveForward Then m_CurrentPage = m_CurrentPage + 1 Else m_CurrentPage = m_CurrentPage - 1
        
    'Hide all inactive panels (and show the active one)
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = (i = m_CurrentPage)
    Next i
    
    'If we are at the beginning, disable the previous button
    cmdPrevious.Enabled = (m_CurrentPage <> 0)
    
    'If we are at the end, change the text of the "next" button; otherwise, make sure it says "next"
    If (m_CurrentPage = picContainer.Count - 2) Then
        cmdNext.Caption = g_Language.TranslateMessage("Start processing!")
    Else
        If (cmdNext.Caption <> g_Language.TranslateMessage("Next")) Then cmdNext.Caption = g_Language.TranslateMessage("Next")
    End If
    
    'Finally, update all the label captions that change according to the active panel
    UpdateWizardText
    
End Sub

'Used to display unique text for each page of the wizard.  The value of m_currentPage is used to determine what text to display.
Private Sub UpdateWizardText()

    Dim sideText As pdString
    Set sideText = New pdString
    
    Select Case m_CurrentPage
        
        'Step 1: choose what photo editing you will apply to each image
        Case 0
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 1: select the photo editing action(s) to apply to each image")
            
            sideText.AppendLine g_Language.TranslateMessage("Welcome to PhotoDemon's batch wizard.  This tool can be used to edit multiple images at once, in what is called a ""batch process"".")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("Start by selecting the photo editing action(s) you want to apply.  If multiple actions are selected, they will be applied in the order they appear on this page.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("Note: a ""macro"" is simply a list of photo editing actions.  It can include any adjustment, filter, or effect in the main program.  You can create a new macro by using the ""Tools -> Macros -> Record new macro"" menu in the main PhotoDemon window.")
            sideText.AppendLineBreak
            sideText.Append g_Language.TranslateMessage("In the next step, you will select the images you want to process.")
            
        'Step 2: add images to list
        Case 1
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 2: prepare the batch list (the list of images to be processed)")
            
            sideText.AppendLine g_Language.TranslateMessage("You can add files to the batch list in two ways:")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("1) By manually adding one or more image file(s) using a standard Open Image dialog.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("2) By adding entire folders at once.  Image file(s) inside the folder (or subfolders, if selected) will be automatically identified.")
            sideText.AppendLineBreak
            sideText.Append g_Language.TranslateMessage("In the next step, you will choose how you want the processed images saved.")
        
        'Step 3: choose the output image format
        Case 2
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 3: choose a destination image format")
            
            sideText.AppendLine g_Language.TranslateMessage("PhotoDemon needs to know which format to use when saving the images in your batch list.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("If ""keep images in their original format"" is selected, PhotoDemon will attempt to save each image in its original format.  If the original format is not supported, a standard format (JPEG or PNG, depending on color depth) will be used.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("If you choose to save images to a new format, please make sure the format you have selected is appropriate for all images in your list.  (For example, images with transparency should be saved to a format that supports transparency!)")
            sideText.AppendLineBreak
            sideText.Append g_Language.TranslateMessage("In the final step, you will choose how you want the saved files to be named.")
            
        'Step 4: choose where processed images will be placed and named
        Case 3
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 4: provide a destination folder and any renaming options")
            
            sideText.AppendLine g_Language.TranslateMessage("In this final step, PhotoDemon needs to know where to save the processed images, and what name to give the new files.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("For your convenience, a number of standard renaming options are also provided.  Note that all items under ""additional rename options"" are optional.")
            sideText.AppendLineBreak
            sideText.Append g_Language.TranslateMessage("Finally, if two or more images in the batch list have the same filename, and the ""original filenames"" option is selected, such files will automatically be given unique filenames upon saving (e.g. ""original-filename (2)"").")
        
        'Step 5: process!
        Case 4
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 5: wait for batch processing to finish")
            
            sideText.AppendLine g_Language.TranslateMessage("Batch processing is now underway.")
            sideText.AppendLineBreak
            sideText.AppendLine g_Language.TranslateMessage("Once the batch processor has processed several images, it will display an estimated time remaining.")
            sideText.AppendLineBreak
            sideText.Append g_Language.TranslateMessage("You can cancel batch processing at any time by pressing the ""Cancel"" button in the bottom-right corner.  If you choose to cancel, any processed images will still be present in the output folder, so you may need to remove them manually.")
            
    End Select
    
    lblExplanation(0).Caption = sideText.ToString()
    
End Sub

'Remove all selected items from the batch conversion list
Private Sub cmdRemove_Click()
    
    If (lstFiles.ListIndex >= 0) Then
        
        Dim prevListIndex As Long
        prevListIndex = lstFiles.ListIndex
        lstFiles.RemoveItem prevListIndex
        If (prevListIndex < lstFiles.ListCount) Then lstFiles.ListIndex = prevListIndex Else lstFiles.ListIndex = lstFiles.ListCount - 1
    
        'And if all files were removed, disable actions that require at least one image
        cmdRemoveAll.Enabled = (lstFiles.ListCount > 0)
        cmdSaveList.Enabled = (lstFiles.ListCount > 0)
        
    End If
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
    'Update the label that displays the number of items in the list
    UpdateBatchListCount
    
    'If user preferences allow, update the current image preview
    If (lstFiles.ListIndex >= 0) Then UpdatePreview lstFiles.List(lstFiles.ListIndex)
            
End Sub

'Remove ALL items from the batch conversion list
Private Sub cmdRemoveAll_Click()
    
    lstFiles.Clear
    UpdatePreview vbNullString
    
    'Because all entries have been removed, disable actions that require at least one image to be present
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
    cmdSaveList.Enabled = False
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
    'Update the label that displays the number of items in the list
    UpdateBatchListCount
    
End Sub

Private Function SaveCurrentBatchList() As Boolean

    'Get the last "open/save image list" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Batch Process", "List Folder", vbNullString)
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("Batch Image List") & " (.pdl)|*.pdl"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save the current list of images")
    
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    Dim sFile As String
    If saveDialog.GetSaveFileName(sFile, , True, cdFilter, 1, tempPathString, cdTitle, ".pdl", FormBatchWizard.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = Files.FileGetPath(sFile)
        UserPrefs.SetPref_String "Batch Process", "List Folder", listPath
        
        'Assemble the output string, which basically just contains the currently selected list of files.
        Dim outputText As pdString
        Set outputText = New pdString
        
        outputText.AppendLine "<PHOTODEMON BATCH CONVERSION LIST>"
        outputText.AppendLine Trim$(Str$(lstFiles.ListCount))
        
        Dim i As Long
        For i = 0 To lstFiles.ListCount - 1
            outputText.AppendLine lstFiles.List(i)
        Next i
        
        outputText.AppendLine "<END OF LIST>"
        outputText.AppendLineBreak
        
        'Write the text out to file using a pdFSO instance
        SaveCurrentBatchList = Files.FileSaveAsText(outputText.ToString(), sFile)
                
    Else
        SaveCurrentBatchList = False
    End If

End Function

Private Sub cmdRemoveFolder_Click()

    If (lstFiles.ListIndex >= 0) Then
        
        m_ListBusy = True
        
        'Retrieve the target path from the currently selected list item
        Dim srcPath As String
        srcPath = m_FSO.FileGetPath(lstFiles.List(lstFiles.ListIndex))
        
        'We now want to iterate through the list, removing items as we go.  Note that the removal criteria varies depending on whether
        ' the user wants subfolders removed as well.
        Dim removeSubfolders As Boolean
        removeSubfolders = chkRemoveSubfolders.Value
        
        Dim testPath As String, removeFile As Boolean
        
        lstFiles.SetAutomaticRedraws False, False
        
        Dim i As Long: i = 0
        Do While (i < lstFiles.ListCount)
            
            removeFile = False
            
            If removeSubfolders Then
                testPath = lstFiles.List(i)
                removeFile = (InStr(1, testPath, srcPath, vbBinaryCompare) <> 0)
            Else
                testPath = m_FSO.FileGetPath(lstFiles.List(i))
                removeFile = Strings.StringsEqual(testPath, srcPath, True)
            End If
            
            If removeFile Then lstFiles.RemoveItem i Else i = i + 1
            
        Loop
        
        lstFiles.SetAutomaticRedraws True, True
        
        m_ListBusy = False
        If (lstFiles.ListIndex >= 0) Then UpdatePreview lstFiles.List(lstFiles.ListIndex) Else UpdatePreview vbNullString
        
        UpdateBatchListCount
        m_ImageListSaved = False
        
    End If

End Sub

Private Sub cmdSaveList_Click()
    
    'Before attempting to save, make sure at least one image has been placed in the list
    If (lstFiles.ListCount = 0) Then
        PDMsgBox "You haven't selected any image files.  Please add one or more files to the batch list before saving.", vbExclamation Or vbOKOnly, "Empty image list"
    Else
        SaveCurrentBatchList
        m_ImageListSaved = True
    End If
    
End Sub

'Open a common dialog and allow the user to select a macro file (to apply to each image in the batch list)
Private Sub cmdSelectMacro_Click()
    
    'Get the last macro-related path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Macro", vbNullString)
    
    Dim cdFilter As String
    cdFilter = "PhotoDemon " & g_Language.TranslateMessage("Macro") & " (." & MACRO_EXT & ")|*." & MACRO_EXT & ";*.thm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    'Prepare a common dialog object
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    Dim sFile As String
   
    'If the user provides a valid macro file, use that as part of the batch process
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, g_Language.TranslateMessage("Open macro"), "." & MACRO_EXT, Me.hWnd) Then
        
        'As a convenience to the user, save this directory as the default macro path
        tempPathString = Files.FileGetPath(sFile)
        UserPrefs.SetPref_String "Paths", "Macro", tempPathString
        
        'Display the selected macro location in the relevant text box
        txtMacro.Text = sFile
        
        'Also, select the macro option button by default
        chkActions(2).Value = True
        
    End If

End Sub

Private Sub cmdSelectOutputPath_Click()
    
    Dim tString As String
    tString = PathBrowseDialog(FormBatchWizard.hWnd)
    
    If (LenB(tString) <> 0) Then
        txtOutputPath.Text = Files.PathAddBackslash(tString)
    
        'Save this new directory as the default path for future usage
        UserPrefs.SetPref_String "Batch Process", "Output Folder", tString
    End If
    
End Sub

Private Sub Form_Load()
        
    Set m_FSO = New pdFSO
        
    Dim i As Long
    
    'Populate all photo-editing-action-related combo boxes, tooltip, and options
        
        'Yes/No for photo edits
        btsPhotoOps.AddItem "no", 0
        btsPhotoOps.AddItem "yes", 1
        btsPhotoOps.ListIndex = 0
        UpdatePhotoOpVisibility
            
        'Resize fit types
        cmbResizeFit.Clear
        cmbResizeFit.AddItem "stretching to fit", 0
        cmbResizeFit.AddItem "fit inclusively", 1
        cmbResizeFit.AddItem "fit exclusively", 2
        cmbResizeFit.ListIndex = 0
        UpdateResizeSample
        
        'For convenience, change the default resize width and height to the current screen resolution
        If (Not g_Displays Is Nothing) Then
            If (Not g_Displays.PrimaryDisplay Is Nothing) Then ucResize.SetInitialDimensions g_Displays.PrimaryDisplay.GetWidth, g_Displays.PrimaryDisplay.GetHeight
        End If
            
        'By default, select "apply no photo editing actions"
        For i = 0 To chkActions.Count - 1
            chkActions(i).Value = False
        Next i
        
        'Populate all file-format-related combo boxes, tooltips, and options
        m_ExportSettingsSet = False
        For i = 0 To ImageFormats.GetNumOfOutputFormats()
            cmbOutputFormat.AddItem ImageFormats.GetOutputFormatDescription(i), i
        Next i
        
        'Save JPEGs by default
        cmbOutputFormat.ListIndex = ImageFormats.GetIndexOfOutputPDIF(PDIF_JPEG)
        
        'Advanced import options are now available for vector formats (e.g. SVG)
        btsVectorImport.AddItem "use embedded size", 0
        btsVectorImport.AddItem "use custom size", 1
        btsVectorImport.ListIndex = 0
        UpdateVectorImportVisibility
        
    'Build default paths from preference file values
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Batch Process", "Output Folder", vbNullString)
    If (LenB(tempPathString) <> 0) And (Files.PathExists(tempPathString)) Then txtOutputPath.Text = tempPathString Else txtOutputPath.Text = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'By default, offer to save processed images in their original format
    optFormat(0).Value = True
    
    'Populate the combo box for file rename options
    cmbOutputOptions.AddItem "original filenames"
    cmbOutputOptions.AddItem "ascending numbers (1, 2, 3, etc.)"
    cmbOutputOptions.ListIndex = 0
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    Dim btnIconSize As Long
    btnIconSize = Interface.FixDPI(32)
    cmdNext.AssignImage "generic_next", , btnIconSize, btnIconSize
    cmdPrevious.AssignImage "generic_previous", , btnIconSize, btnIconSize
    
    'Set the current page number to 0
    m_CurrentPage = 0
    
    'Mark the current image list as "not saved"
    m_ImageListSaved = False
    
    'Display appropriate help text and wizard title
    UpdateWizardText
    
    'Display some text manually to make sure translations are handled correctly
    txtMacro.Text = g_Language.TranslateMessage("no macro selected")
    lblExplanationFormat.Caption = g_Language.TranslateMessage("if PhotoDemon does not support an image's original format, a standard format will be used")
    lblExplanationFormat.Caption = lblExplanationFormat.Caption & vbCrLf & " " & g_Language.TranslateMessage("( specifically, JPEG at 92% quality for photographs, and lossless PNG for non-photographs )")
    
    'Hide all inactive wizard panes
    For i = 1 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Switch the batch file listview to file mode (which will automatically truncate long paths)
    lstFiles.SetDisplayMode_Files True
    
    'Apply visual themes and translations
    ApplyThemeAndTranslations Me
    
    'Cache the translations for words used in high-performance processes
    m_wordForBatchList = g_Language.TranslateMessage("batch list")
    m_wordForItem = g_Language.TranslateMessage("item")
    m_wordForItems = g_Language.TranslateMessage("items")
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not AllowedToExit()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstFiles_Click()

    If (Not m_ListBusy) Then
        
        'Perform a quick check to make sure the selected image hasn't been removed
        Dim targetFile As String
        targetFile = lstFiles.List(lstFiles.ListIndex)
        
        cmdRemove.Enabled = Files.FileExists(targetFile)
        If cmdRemove.Enabled Then UpdatePreview targetFile Else lstFiles.RemoveItem lstFiles.ListIndex
        
    End If
    
End Sub

Private Sub lstFiles_CustomDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'File lists (from e.g. Explorer) can be drag-dropped onto the batch listview
    If (Not Data Is Nothing) Then
    If Data.GetFormat(vbCFFiles) Then
        
        Dim cFiles As pdStringStack
        If VBHacks.GetDragDropFileListW(Data, cFiles) Then

            'As a convenience, sort the list alphabetically and remove any duplicate entries
            cFiles.SortLogically True
            
            'Bulk add the entire file collection to the list box
            lstFiles.SetAutomaticRedraws False, False
            
            Dim i As Long
            For i = 0 To cFiles.GetNumOfStrings - 1
                lstFiles.AddItem cFiles.GetString(i)
            Next i
            
            Set cFiles = Nothing
            lstFiles.SetAutomaticRedraws True, True
            
            'Update any related UI elements to reflect the new list
            UpdateBatchListCount
            m_ImageListSaved = False
            cmdRemoveAll.Enabled = (lstFiles.ListCount > 0)
            cmdSaveList.Enabled = (lstFiles.ListCount > 0)
            
        End If
        
    
    'End validation checks
    End If
    End If
    
End Sub

Private Sub lstFiles_CustomDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub

'Update the active image preview in the top-right
Private Sub UpdatePreview(ByVal srcImagePath As String, Optional ByVal forceUpdate As Boolean = False)
    
    lblCurrentFile.Caption = srcImagePath
    
    'Only redraw the preview if it doesn't match the last image we previewed
    If (chkEnablePreview.Value And (Strings.StringsNotEqual(m_CurImagePreview, srcImagePath, True) Or forceUpdate)) Then
    
        'Use PD's central load function to load a copy of the requested image
        Dim tmpDIB As pdDIB: Set tmpDIB = New pdDIB
        Dim loadSuccessful As Boolean: loadSuccessful = False
        If (LenB(srcImagePath) <> 0) Then loadSuccessful = Loading.QuickLoadImageToDIB(srcImagePath, tmpDIB, False, False)
        
        'If the image load failed, display a placeholder message; otherwise, render the image to the picture box
        If loadSuccessful Then
            picPreview.CopyDIB tmpDIB, True, True
        Else
            picPreview.PaintText g_Language.TranslateMessage("previews disabled"), 10!, False, True
        End If
        
        'Remember the name of the current preview; this saves us having to reload the preview any more than
        ' is absolutely necessary
        m_CurImagePreview = srcImagePath
    
    End If
    
End Sub

Private Sub UpdateBatchListCount()
    
    Select Case lstFiles.ListCount
    
        Case 0
            lstFiles.Caption = m_wordForBatchList & ":"
        Case 1
            lstFiles.Caption = m_wordForBatchList & " (" & lstFiles.ListCount & " " & m_wordForItem & "):"
        Case Else
            lstFiles.Caption = m_wordForBatchList & " (" & lstFiles.ListCount & " " & m_wordForItems & "):"
            
    End Select
    
End Sub

Private Sub optCase_Click(Index As Integer)
    chkRenameCase.Value = True
End Sub

'When the user presses "Start Conversion", this routine is triggered.
Private Sub PrepareForBatchConversion()

    BatchConvertMessage g_Language.TranslateMessage("Preparing batch processing engine...")
    
    Dim i As Long
    
    'Display the progress panel
    picContainer(picContainer.Count - 1).Visible = True
    For i = 0 To picContainer.Count - 2
        picContainer(i).Visible = False
    Next i
    
    'Hide back/forward buttons
    cmdPrevious.Visible = False
    cmdNext.Visible = False
    
    'Let the rest of the program know that batch processing has begun
    Macros.SetMacroStatus MacroBATCH
    
    Dim curBatchFile As Long
    Dim srcFilename As String, dstFilename As String
    
    Dim totalNumOfFiles As Long
    totalNumOfFiles = lstFiles.ListCount
    
    'Prepare the folder that will receive the processed images
    Dim outputPath As String
    outputPath = Files.PathAddBackslash(txtOutputPath)
    If (Not Files.PathExists(outputPath)) Then Files.PathCreate outputPath, True
    
    'Prepare the progress bar, which will keep the user updated on our progress.
    pbBatch.Max = totalNumOfFiles
    pbBatch.Value = 0
    
    'Let's also give the user an estimate of how long this is going to take.  We estimate time by determining an
    ' approximate "time-per-image" value, then multiplying that by the number of images remaining.  The progress bar
    ' will display this, automatically updated, as each image is completed.
    Dim timeMsg As String
    timeMsg = vbNullString
    
    Dim lastTimeCalculation As Long
    lastTimeCalculation = &H7FFFFFFF
    
    Dim timeStarted As Currency
    VBHacks.GetHighResTime timeStarted
    
    Dim numFilesTimeNotUpdated As Long
    
    'We're now gonna build two stacks of strings:
    ' 1) all input files
    ' 2) all output files
    '
    'Note that (2) won't include *all* the correct filename information yet (because the final wizard page allows for
    ' all kinds of custom filename settings) - but it will include each image's final *path*, including all subfolders,
    ' with the original filename tacked on (with its original file extension).
    '
    'A later function handles all optional output filename settings, but we need the correct folder structure established
    ' beforehand so we can create intermediary subfolders.
    Dim srcListFiles As pdStringStack, dstListFiles As pdStringStack
    Set srcListFiles = New pdStringStack
    
    'Some normal import behaviors (e.g. displaying rasterize size prompts for vector images) must be overridden
    ' during a batch conversion.  Build a param string with all override settings, which we can then blindly forward
    ' to the image importer, and individual import components can strip out whichever override settings may be
    'relevant to them.
    Dim overrideParams As pdSerialize
    Set overrideParams = New pdSerialize
    
    overrideParams.AddParam "vector-size-use-default", (btsVectorImport.ListIndex = 0)
    overrideParams.AddParam "vector-size-x", spnVectorImport(0).Value
    overrideParams.AddParam "vector-size-y", spnVectorImport(1).Value
    
    For i = 0 To lstFiles.ListCount - 1
        srcListFiles.AddString lstFiles.List(i)
    Next i
    
    'Create a matching list of destination files
    If chkOutputPreserveFolders.Value Then
        
        'Make sure we can create valid destination paths for each output image.  (This is not expected
        ' to fail, but combining paths is complicated and difficult to test exhaustively.)
        If (Not Files.PathRebaseListOnNewPath(srcListFiles, dstListFiles, outputPath)) Then
            
            'Path merging failed.  Fall back to "do not preserve output subfolders" mode.
            chkOutputPreserveFolders.Value = False
            
        End If
        
    End If
    
    'This is where the fun begins.  Loop through every file in the list, and process them one-by-one using the options requested
    ' by the user.
    For curBatchFile = 0 To totalNumOfFiles - 1
    
        'Pause for keypresses - this allows the user to press "Escape" to cancel the operation
        DoEvents
        If (Macros.GetMacroStatus = MacroCANCEL) Then GoTo MacroCanceled
        
        'Give the user a progress update
        BatchConvertMessage g_Language.TranslateMessage("Processing image # %1 of %2", (curBatchFile + 1), totalNumOfFiles)
        pbBatch.Value = curBatchFile
        
        'Retrieve the source file and validate it
        srcFilename = srcListFiles.GetString(curBatchFile)
        If Files.FileExists(srcFilename) Then
            
            Dim importDialogResults As VbMsgBoxResult
            importDialogResults = vbYes
            
            'Load the target image
            If Loading.LoadFileAsNewImage(srcFilename, vbNullString, False, importDialogResults, False, overrideParams.GetParamString()) Then
                
                'Manually activate the just-loaded image
                Dim tmpStack As pdStack
                Set tmpStack = Nothing
                PDImages.GetListOfActiveImageIDs tmpStack
                CanvasManager.ActivatePDImage tmpStack.GetInt(tmpStack.GetNumOfInts - 1), newImageJustLoaded:=True
                
                'With the image loaded, it is time to apply any requested photo editing actions.
                If (btsPhotoOps.ListIndex = 1) Then ApplyEditOperations
                
                'With the macro complete, prepare the file for saving.  (This function will determine both
                ' a final filename and a proper file extension.)
                dstFilename = GetFinalFilename(srcFilename, outputPath, dstListFiles, curBatchFile)
                
                'Request a save from the PhotoDemon_SaveImage method, and pass it the parameter string created by the user
                ' on the matching wizard panel.  Note that we need to silently swap-in animation parameters instead of
                ' static ones, if the source image is animated.
                Dim finalSaveParams As String
                finalSaveParams = m_ExportSettingsFormat
                If (PDImages.GetActiveImage.IsAnimated() And chkExportAnimation.Value) Then finalSaveParams = m_ExportSettingsFormatAnimation
                
                ' TODO: metadata for animated images
                ' TODO: track success/fail results and collate any failures into a list that we can report to the user
                Saving.PhotoDemon_BatchSaveImage PDImages.GetActiveImage(), dstFilename, PDImages.GetActiveImage.GetCurrentFileFormat, finalSaveParams, m_ExportSettingsMetadata
                
                'Unload the finished image
                CanvasManager.FullPDImageUnload PDImages.GetActiveImageID()
            
            End If
            
            'Update our running time estimate
            If UpdateTimeEstimate(timeMsg, curBatchFile + 1, totalNumOfFiles - (curBatchFile + 1), timeStarted, lastTimeCalculation, numFilesTimeNotUpdated) Then BatchTimeMessage timeMsg
            
        End If
                
    'Carry on
    Next curBatchFile
    
    Macros.SetMacroStatus MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    'Max out the progess bar and display a success message
    pbBatch.Value = pbBatch.Max
    BatchConvertMessage g_Language.TranslateMessage("%1 files were successfully processed!", totalNumOfFiles)
    BatchTimeMessage vbNullString
    
    'Finally, there is no longer any need for the user to save their batch list, as the batch process is complete.
    m_ImageListSaved = True
    
    Exit Sub
    
MacroCanceled:

    Macros.SetMacroStatus MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Reset the progress bar
    pbBatch.Value = 0
    
    Dim cancelMsg As String
    cancelMsg = g_Language.TranslateMessage("Batch conversion canceled.  %1 image(s) were processed before cancelation.  Last processed image was ""%2"".", curBatchFile, lstFiles.List(curBatchFile))
    BatchConvertMessage cancelMsg
    BatchTimeMessage vbNullString
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    m_ImageListSaved = True
    
End Sub

'Apply all selected photo editing operations to the image
Private Sub ApplyEditOperations()

    'If the user has requested automatic lighting fixes, apply it now
    If chkActions(0).Value Then Process "Auto correct", , , UNDO_Layer
    
    'If the user has requested an image resize, apply it now
    If chkActions(1).Value Then
        
        'Generate a compatible list of options for PD's resampling engine
        Dim resizeParams As pdSerialize
        Set resizeParams = New pdSerialize
        With resizeParams
            .SetParamVersion 3#
            .AddParam "width", ucResize.ResizeWidth
            .AddParam "height", ucResize.ResizeHeight
            .AddParam "unit", ucResize.UnitOfMeasurement
            .AddParam "ppi", ucResize.ResizeDPIAsPPI
            .AddParam "resample", Resampling.GetResamplerName(rf_Automatic)
            .AddParam "approximations-ok", True
            .AddParam "fit", cmbResizeFit.ListIndex
            .AddParam "fillcolor", vbWhite
            .AddParam "target", pdat_Image
        End With
        
        Process "Resize image", , resizeParams.GetParamString
        
    End If
    
    'If the user has requested a macro, play it now
    If chkActions(2).Value Then Macros.PlayMacroFromFile txtMacro
    
End Sub

'Do not pass invalid files or paths to this function.  It does not validate inputs.
Private Function GetFinalFilename(ByRef originalFilename As String, ByVal outputPath As String, ByRef dstListFiles As pdStringStack, ByVal curBatchFile As Long) As String
    
    'Before we even think about output path, start by stripping the incoming filename
    ' down to just its filename.
    Dim tmpFilename As String
    tmpFilename = Files.FileGetName(originalFilename, True)
    
    'Start working on building a filename that matches the user's output settings.
    
    'First, append any prefix/suffix text
    If (cmbOutputOptions.ListIndex = 0) Then
        If chkRenamePrefix.Value Then tmpFilename = txtAppendFront & tmpFilename
        If chkRenameSuffix.Value Then tmpFilename = tmpFilename & txtAppendBack
    Else
        tmpFilename = curBatchFile + 1
        If chkRenamePrefix.Value Then tmpFilename = txtAppendFront & tmpFilename
        If chkRenameSuffix.Value Then tmpFilename = tmpFilename & txtAppendBack
    End If
    
    'If requested, remove any specified text from the filename
    If chkRenameRemove.Value And (LenB(txtRenameRemove) <> 0) Then
    
        'Use case-sensitive or case-insensitive matching as requested
        If chkRenameCaseSensitive.Value Then
            If (InStr(1, tmpFilename, txtRenameRemove, vbBinaryCompare) <> 0) Then
                tmpFilename = Replace(tmpFilename, txtRenameRemove, vbNullString, , , vbBinaryCompare)
            End If
        Else
            If (InStr(1, tmpFilename, txtRenameRemove, vbTextCompare) <> 0) Then
                tmpFilename = Replace(tmpFilename, txtRenameRemove, vbNullString, , , vbTextCompare)
            End If
        End If
        
    End If
    
    'Replace spaces with underscores if requested
    If chkRenameSpaces.Value Then
        If (InStr(1, tmpFilename, " ") <> 0) Then tmpFilename = Replace$(tmpFilename, " ", "_")
    End If
    
    'Change the full filename's case if requested
    If chkRenameCase.Value Then
        If optCase(0).Value Then tmpFilename = LCase$(tmpFilename) Else tmpFilename = UCase$(tmpFilename)
    End If
    
    'Attach a proper image format file extension and save format ID number based off the user's
    ' requested output format
    Dim tmpFileExtension As String
    
    'Possibility 1: use original file format
    If optFormat(0).Value Then
        
        'See if this image's file format is supported by the export engine.
        If (ImageFormats.GetIndexOfOutputPDIF(PDImages.GetActiveImage.GetCurrentFileFormat) = -1) Then
            
            'The current format isn't supported.  Use PNG as it's the best compromise of
            ' lossless, well-supported, and reasonably well-compressed.
            tmpFileExtension = ImageFormats.GetExtensionFromPDIF(PDIF_PNG)
            PDImages.GetActiveImage.SetCurrentFileFormat PDIF_PNG
            
        Else
            
            'This format IS supported, so use the default extension
            tmpFileExtension = ImageFormats.GetExtensionFromPDIF(PDImages.GetActiveImage.GetCurrentFileFormat)
        
        End If
        
    'Possibility 2: force all images to a single file format
    Else
        tmpFileExtension = ImageFormats.GetOutputFormatExtension(cmbOutputFormat.ListIndex)
        PDImages.GetActiveImage.SetCurrentFileFormat ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)
    End If
    
    'If the user has requested lower- or upper-case, we now need to convert the extension as well
    If chkRenameCase.Value Then
        If optCase(0).Value Then tmpFileExtension = LCase$(tmpFileExtension) Else tmpFileExtension = UCase$(tmpFileExtension)
    End If
    
    'We now have a finished filename, but we still need to deal with output path.
    
    'The base output path we use varies depending on the "preserve subfolders" option
    If chkOutputPreserveFolders.Value Then
        
        'The user wants output folders preserved.  Silently replace the passed output path with the one
        ' calculated by PD's path-matcher.
        outputPath = Files.FileGetPath(dstListFiles.GetString(curBatchFile))
        
        'Ensure all folders in this output path exist.
        Files.PathCreate outputPath, True
    
    'No /Else branch required - the outputPath string already contains the flat path all images are being saved to.
    End If
    
    'Because removing specified text from filenames may lead to files with the same name, call the incrementFilename
    ' function to find a unique filename of the "filename (n+1)" variety if necessary.  This will also prepend the
    ' drive and directory structure determined by the previous step.
    tmpFilename = outputPath & Files.IncrementFilename(outputPath, tmpFilename, tmpFileExtension) & "." & tmpFileExtension
    
    'Return the final result
    GetFinalFilename = tmpFilename
    
End Function

'Update the current "time remaining" estimate
Private Function UpdateTimeEstimate(ByRef dstMessage As String, ByVal numFilesProcessed As Long, ByVal numFilesRemaining As Long, ByVal timeStarted As Currency, ByRef lastTimeCalculation As Long, ByRef numFilesTimeNotUpdated As Long) As Boolean
    
    UpdateTimeEstimate = True
    
    Dim timeElapsed As Double, timeRemaining As Double, timePerFile As Double
    Dim minutesRemaining As Long, secondsRemaining As Long
    
    If (numFilesProcessed >= 10) Then
        
        timeElapsed = VBHacks.GetTimerDifferenceNow(timeStarted)
        timePerFile = timeElapsed / numFilesProcessed
        timeRemaining = timePerFile * numFilesRemaining
        
        minutesRemaining = Int(timeRemaining / 60#)
        secondsRemaining = Int(timeRemaining) Mod 60
        If (minutesRemaining > 10) Then secondsRemaining = (secondsRemaining \ 5) * 5
        
        'If there are a *ton* of images left to process, reduce our update frequency to minimize
        ' the potential for very poor time estimates.
        Dim okToUpdate As Boolean
        okToUpdate = (numFilesRemaining < 250) Or ((numFilesProcessed Mod 5) = 0)
        
        'Normally, we only want to update the screen if our current time estimate is less than our previous
        ' time estimates.  (We do this because it's frustrating if time estimates jump around instead of
        ' keeping a steady downward trend.)  However, if many images pass and our time estimates are still
        ' too low, then we concede defeat and update the screen accordingly.
        If okToUpdate Then
            If (timeRemaining < lastTimeCalculation) Or (numFilesTimeNotUpdated > 4) Then
                numFilesTimeNotUpdated = 0
                lastTimeCalculation = timeRemaining
                dstMessage = g_Language.TranslateMessage("Estimated time remaining: %1:%2", minutesRemaining, Format$(secondsRemaining, "00"))
            Else
                numFilesTimeNotUpdated = numFilesTimeNotUpdated + 1
                UpdateTimeEstimate = False
            End If
        Else
            UpdateTimeEstimate = False
        End If

    Else
        dstMessage = g_Language.TranslateMessage("Estimating time remaining...")
    End If
            
End Function

'Display time and progress updates to the user
Private Sub BatchConvertMessage(ByRef newMessage As String)
    lblBatchProgress.Caption = newMessage
    lblBatchProgress.RequestRefresh
End Sub

Private Sub BatchTimeMessage(ByRef newMessage As String)
    lblTimeRemaining.Caption = newMessage
    lblTimeRemaining.RequestRefresh
End Sub

Private Sub UpdatePhotoOpVisibility()
    lblExplanation(1).Visible = (btsPhotoOps.ListIndex = 0)
    picPhotoEdits.Visible = (btsPhotoOps.ListIndex <> 0)
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not chkEnablePreview.Value) Then picPreview.PaintText g_Language.TranslateMessage("previews disabled"), 10!, False, True
End Sub

Private Sub picResizeDemo_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    UpdateResizeSample
End Sub

Private Sub UpdateVectorImportVisibility()
    
    Dim vecSettingsVisible As Boolean
    vecSettingsVisible = (btsVectorImport.ListIndex = 1)
    
    Dim i As Long
    For i = lblVectorImport.lBound To lblVectorImport.UBound
        lblVectorImport(i).Visible = vecSettingsVisible
        spnVectorImport(i).Visible = vecSettingsVisible
    Next i
    
End Sub
