VERSION 5.00
Begin VB.Form FormBatchWizard 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Batch Process Wizard"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      Caption         =   "&Previous"
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
      Caption         =   "&Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   11760
      TabIndex        =   2
      Top             =   7530
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      Caption         =   "&Cancel"
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   6705
      Index           =   0
      Left            =   120
      Top             =   720
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   11827
      Caption         =   "(text populated at run-time)"
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
      Index           =   3
      Left            =   3300
      TabIndex        =   8
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdDropDown cmbOutputOptions 
         Height          =   375
         Left            =   480
         TabIndex        =   46
         Top             =   1680
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButton cmdSelectOutputPath 
         Height          =   615
         Left            =   6600
         TabIndex        =   26
         Top             =   330
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "Select destination folder..."
         FontSize        =   9
      End
      Begin PhotoDemon.pdTextBox txtRenameRemove 
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Top             =   4440
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendBack 
         Height          =   315
         Left            =   5640
         TabIndex        =   28
         Top             =   3360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendFront 
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   3360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         Text            =   "NEW_"
      End
      Begin PhotoDemon.pdTextBox txtOutputPath 
         Height          =   315
         Left            =   480
         TabIndex        =   30
         Top             =   480
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
         Top             =   5520
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
         Top             =   2880
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         Caption         =   "add a prefix to each filename:"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkRenameSuffix 
         Height          =   330
         Left            =   5280
         TabIndex        =   12
         Top             =   2880
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   582
         Caption         =   "add a suffix to each filename:"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkRenameRemove 
         Height          =   330
         Left            =   480
         TabIndex        =   13
         Top             =   3960
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   582
         Caption         =   "remove the following text (if found) from each filename:"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkRenameCase 
         Height          =   330
         Left            =   480
         TabIndex        =   14
         Top             =   5040
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   582
         Caption         =   "force each filename, including extension, to the following case:"
         Value           =   0
      End
      Begin PhotoDemon.pdRadioButton optCase 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   16
         Top             =   5520
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         Caption         =   "UPPERCASE"
      End
      Begin PhotoDemon.pdCheckBox chkRenameSpaces 
         Height          =   330
         Left            =   480
         TabIndex        =   17
         Top             =   6060
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   582
         Caption         =   "replace spaces in filenames with underscores"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkRenameCaseSensitive 
         Height          =   330
         Left            =   5760
         TabIndex        =   18
         Top             =   4440
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   582
         Caption         =   "use case-sensitive matching"
         Value           =   0
      End
      Begin PhotoDemon.pdLabel lblDstFilename 
         Height          =   285
         Left            =   120
         Top             =   1200
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   503
         Caption         =   "after images are processed, save them with the following name:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblOptionalText 
         Height          =   285
         Left            =   120
         Top             =   2400
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
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   2
      Left            =   3300
      TabIndex        =   5
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin PhotoDemon.pdButton cmdExportSettings 
         Height          =   735
         Left            =   720
         TabIndex        =   48
         Top             =   2400
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1296
         Caption         =   "set export settings for this format..."
      End
      Begin PhotoDemon.pdDropDown cmbOutputFormat 
         Height          =   375
         Left            =   720
         TabIndex        =   47
         Top             =   1800
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
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   0
      Left            =   3300
      TabIndex        =   4
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
         TabIndex        =   32
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9525
         Begin VB.PictureBox picResizeDemo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Left            =   6720
            ScaleHeight     =   50
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   191
            TabIndex        =   36
            Top             =   2880
            Width           =   2865
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
            Value           =   0
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
            Value           =   0
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
            Value           =   0
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
      TabIndex        =   3
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
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
         TabIndex        =   44
         Top             =   5670
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Caption         =   "include subfolders"
         Value           =   0
      End
      Begin PhotoDemon.pdListBox lstFiles 
         Height          =   3405
         Left            =   120
         TabIndex        =   42
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
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2925
         Left            =   6600
         ScaleHeight     =   193
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   6
         Top             =   465
         Width           =   3135
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
         TabIndex        =   41
         Top             =   4995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "add entire folder(s)..."
      End
      Begin PhotoDemon.pdButton cmdRemoveFolder 
         Height          =   615
         Left            =   3600
         TabIndex        =   43
         Top             =   4995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "remove all images in this folder"
      End
      Begin PhotoDemon.pdCheckBox chkRemoveSubfolders 
         Height          =   375
         Left            =   3585
         TabIndex        =   45
         Top             =   5670
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "include subfolders"
         Value           =   0
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6780
      Index           =   4
      Left            =   3300
      TabIndex        =   24
      Top             =   660
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      Begin VB.PictureBox picBatchProgress 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   629
         TabIndex        =   25
         Top             =   3120
         Width           =   9435
      End
      Begin PhotoDemon.pdLabel lblBatchProgress 
         Height          =   645
         Left            =   285
         Top             =   2400
         Width           =   9435
         _ExtentX        =   0
         _ExtentY        =   0
         Alignment       =   2
         Caption         =   "(batch conversion process will appear here at run-time)"
         ForeColor       =   -2147483640
         Layout          =   1
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
'Copyright 2007-2018 by Tanner Helland
'Created: 3/Nov/07
'Last updated: 09/September/16
'Last update: complete overhaul of UI and underlying logic
'
'PhotoDemon's batch process wizard is one of its most unique features.  It integrates tightly with PD's
' macro recorder, which allows any combination of actions to be applied to any set of images.  Neat stuff!
'
'The current batch wizard is broken into four steps.
'
'1) Select which photo editing operations (if any) to apply to these images.  This step is optional;
'    if no photo editing actions are selected, only format conversion and/or renaming will be applied.
'
'2) Build the batch list, e.g. the list of files to be processed.  This is currently the most complicated
'    page of the wizard, but it allows some nice features, like constructing a list of images from any
'    number of source directories.  (Many batch tools limit you to just one source folder, ugh.)
'
'3) Select output file format.  There are two choices: retain original format (with limitations, e.g.
'    read-only formats like manufacturer-specific RAW files will be saved as JPEGS), or save to some new
'    format.  If a new format is selected, PD's standard export dialogs are available to set whatever
'    export parameters the user wants.
'
'4) Choose where the new images will go and how they will be named.  This includes a number of basic
'    renaming options.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Current active page in the wizard
Private m_CurrentPage As Long

'Has the current list of images been saved?
Private m_ImageListSaved As Boolean

'Current list of image format parameters
Private m_FormatParams As String

'The path to the image currently rendered to the "image preview" box.  (We cache this to optimize redraws; if the
' path hasn't changed since the last request, we do not redraw the preview.)
Private m_CurImagePreview As String

'Because these words are used frequently, if we have to translate them every time they're used, it slows down the
' process considerably.  So cache them in advance.
' TODO: fix this, because word order (obviously) is not consistent from language to language
Private m_wordForBatchList As String, m_wordForItem As String, m_wordForItems As String

'We maintain folder paths locally, in case the user wants to add multiple folders in succession
Private m_LastBatchFolder As String

'While we're processing the list (for example, when removing items automatically), we want to ignore any events raised by the list
Private m_ListBusy As Boolean

'Export settings were overhauled for 7.0's release.  Batch processing now uses the same export dialogs as PD's regular
' save functions.  To make sure the user actually sets export settings before progressing, we use this tracker.
Private m_ExportSettingsSet As Boolean, m_ExportSettingsFormat As String, m_ExportSettingsMetadata As String

'System progress bar control
Private sysProgBar As cProgressBarOfficial

'This dialog interacts with a lot of file-system bits.  This module-level pdFSO object is initialized at Form_Load(),
' and can be used wherever convenient.
Private m_FSO As pdFSO

Private Sub btsPhotoOps_Click(ByVal buttonIndex As Long)
    UpdatePhotoOpVisibility
End Sub

Private Sub chkEnablePreview_Click()
    
    picPreview.Picture = LoadPicture("")
        
    'If the user is enabling previews, try to display the last item the user selected in the SOURCE list box
    If CBool(chkEnablePreview) Then
        
        If (lstFiles.ListIndex >= 0) Then UpdatePreview lstFiles.List(lstFiles.ListIndex), True
        
    'If the user is disabling previews, clear the picture box and display a notice
    Else
        Dim strToPrint As String
        strToPrint = g_Language.TranslateMessage("Previews disabled")
        picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.textWidth(strToPrint)) \ 2
        picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.textHeight(strToPrint)) \ 2
        picPreview.Print strToPrint
    End If
    
End Sub

'By default, neither case-related option button is selected.  Default to lowercase when the RenameCase checkbox is used.
Private Sub chkRenameCase_Click()
    If (Not optCase(0).Value) And (Not optCase(1).Value) Then optCase(0).Value = True
End Sub

Private Sub cmbOutputFormat_Click()
    
    'If this format doesn't support export settings, hide the "set export settings" button
    If g_ImageFormats.IsExportDialogSupported(g_ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)) Then
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
    
End Sub

Private Sub cmbResizeFit_Click()
    
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
    picResizeDemo.Picture = LoadPicture("")
    tmpDIB.AlphaBlendToDC picResizeDemo.hDC
    picResizeDemo.Picture = picResizeDemo.Image

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

Private Sub cmdAddFolders_Click()
    
    If (Len(m_LastBatchFolder) = 0) Then m_LastBatchFolder = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
    
    Dim folderPath As String
    folderPath = Files.PathBrowseDialog(Me.hWnd, m_LastBatchFolder)
    
    If (Len(folderPath) <> 0) Then
        
        m_LastBatchFolder = folderPath
        
        Dim listOfFiles As pdStringStack
        If m_FSO.RetrieveAllFiles(folderPath, listOfFiles, CBool(chkAddSubfoldersToo.Value), False, g_ImageFormats.GetListOfInputFormats("|", False)) Then
                
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
        
    End If

End Sub

'Cancel and exit the dialog, with optional prompts as necessary (see Form_QueryUnload)
Private Sub cmdCancel_Click()
    
    If (m_CurrentPage = picContainer.Count - 1) Then
        
        If (Macros.GetMacroStatus <> MacroSTOP) Then
        
            Dim msgReturn As VbMsgBoxResult
            msgReturn = PDMsgBox("Are you sure you want to cancel the current batch process?", vbYesNoCancel Or vbExclamation, "Cancel batch processing")
            If (msgReturn = vbYes) Then Macros.SetMacroStatus MacroCANCEL
            
        Else
            Unload Me
        End If
        
    Else
        Unload Me
    End If
    
End Sub

Private Function AllowedToExit() As Boolean
    
    AllowedToExit = True
    
    'If the user has created a list of images to process and they attempt to exit without saving the list,
    ' give them a chance to save it.
    If (m_CurrentPage < picContainer.Count - 1) Then
    
        If (Not m_ImageListSaved) Then
        
            If (lstFiles.ListCount > 0) Then
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("If you exit now, your batch list (the list of images to be processed) will be lost.  By saving your list, you can easily resume this batch operation at a later date." & vbCrLf & vbCrLf & "Would you like to save your batch list before exiting?", vbExclamation Or vbYesNoCancel, "Unsaved image list")
                
                Select Case msgReturn
                    
                    Case vbYes
                        If SaveCurrentBatchList() Then AllowedToExit = True Else AllowedToExit = False
                    
                    Case vbNo
                        AllowedToExit = True
                    
                    Case vbCancel
                        AllowedToExit = False
                            
                End Select
                
            End If
            
        End If
        
    End If
    
End Function

Private Sub cmdExportSettings_Click()
    
    'Convert the current dropdown index into a PD format constant
    Dim saveFormat As PD_IMAGE_FORMAT
    saveFormat = g_ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)
    
    'See if this format even supports dialogs...
    If g_ImageFormats.IsExportDialogSupported(saveFormat) Then
        
        'The saving module will now raise a dialog specific to the selected format.  If successful, it will fill
        ' the passed settings and metadata strings with XML data describing the user's settings.
        If Saving.GetExportParamsFromDialog(Nothing, saveFormat, m_ExportSettingsFormat, m_ExportSettingsMetadata) Then
            m_ExportSettingsSet = True
             
        'If the user cancels the dialog, exit immediately
        Else
            m_ExportSettingsSet = False
            m_ExportSettingsFormat = vbNullString
            m_ExportSettingsMetadata = vbNullString
        End If
    
    Else
        m_ExportSettingsSet = True
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
    End If
    
End Sub

'Load a list of images (previously saved from within PhotoDemon) to the batch list
Private Sub cmdLoadList_Click()
    
    Dim sFile As String
    
    'Get the last "open/save image list" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "List Folder", "")
    
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
        g_UserPreferences.SetPref_String "Batch Process", "List Folder", listPath
        
        'Load the file using pdFSO, which is Unicode-compatible
        Dim fileContents As String
        If Files.FileLoadAsString(sFile, fileContents) And (InStr(1, fileContents, vbCrLf, vbBinaryCompare) > 0) Then
            
            'The file was originally delimited by vbCrLf.  Parse it now.
            Dim fileLines() As String
            fileLines = Split(fileContents, vbCrLf, , vbBinaryCompare)
            
            If (UBound(fileLines) > 0) Then
                
                'Validate the first line of the file
                If Strings.StringsEqual(fileLines(0), "<PHOTODEMON BATCH CONVERSION LIST>", True) Then
                    
                    'If the user has already created a list of files to process, ask if they want to replace or append
                    ' the loaded entries to their current list.
                    If (lstFiles.ListCount > 0) Then
                
                    Dim msgReturn As VbMsgBoxResult
                    msgReturn = PDMsgBox("You have already created a list of images for processing.  The list of images inside this file will be appended to the bottom of your current list.", vbOKCancel Or vbInformation, "Batch process notification")
                    
                    If msgReturn = vbCancel Then Exit Sub
                    
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
                        
                Else
                    PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation Or vbOKOnly, "Invalid list file"
                    Exit Sub
                End If
                
            Else
                PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation Or vbOKOnly, "Invalid list file"
                Exit Sub
            End If
            
        Else
            PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation Or vbOKOnly, "Invalid list file"
            Exit Sub
        End If
        
        'Note that the current list has been saved (technically it hasn't, I realize, but it exists in a file in this exact state
        ' so close enough!)
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
    If moveForward And m_CurrentPage = picContainer.Count - 2 Then
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
                If CBool(chkActions(1)) Then
                    If Not ucResize.IsValid(True) Then Exit Sub
                End If
                
                'If the user wants us to apply a macro, ensure that the macro text box has a macro file specified
                If CBool(chkActions(2)) And ((txtMacro.Text = "no macro selected") Or (Len(txtMacro.Text) = 0)) Then
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
            ' "set export options" button (to define what export settings we'll use).
            
            ' contains all of the user's selected image format options (JPEG quality, etc)
            If (optFormat(1) And moveForward) Then
            
                If (Not m_ExportSettingsSet) Then
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

    Dim sideText As String
    sideText = "(description forthcoming)"

    Select Case m_CurrentPage
        
        'Step 1: choose what photo editing you will apply to each image
        Case 0
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 1: select the photo editing action(s) to apply to each image")
            
            sideText = g_Language.TranslateMessage("Welcome to PhotoDemon's batch wizard.  This tool can be used to edit multiple images at once, in what is called a ""batch process"".")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Start by selecting the photo editing action(s) you want to apply.  If multiple actions are selected, they will be applied in the order they appear on this page.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Note: a ""macro"" is simply a list of photo editing actions.  It can include any adjustment, filter, or effect in the main program.  You can create a new macro by using the ""Tools -> Macros -> Record new macro"" menu in the main PhotoDemon window.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the next step, you will select the images you want to process.")
            
        'Step 2: add images to list
        Case 1
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 2: prepare the batch list (the list of images to be processed)")
            
            sideText = g_Language.TranslateMessage("You can add files to the batch list in two ways:")
            sideText = sideText & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("1) By manually adding one or more image file(s) using a standard Open Image dialog.")
            sideText = sideText & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("2) By adding entire folders at once.  Image file(s) inside the folder (or subfolders, if selected) will be automatically identified.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the next step, you will choose how you want the processed images saved.")
        
        'Step 3: choose the output image format
        Case 2
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 3: choose a destination image format")
            
            sideText = g_Language.TranslateMessage("PhotoDemon needs to know which format to use when saving the images in your batch list.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("If ""keep images in their original format"" is selected, PhotoDemon will attempt to save each image in its original format.  If the original format is not supported, a standard format (JPEG or PNG, depending on color depth) will be used.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("If you choose to save images to a new format, please make sure the format you have selected is appropriate for all images in your list.  (For example, images with transparency should be saved to a format that supports transparency!)")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the final step, you will choose how you want the saved files to be named.")
            
        'Step 4: choose where processed images will be placed and named
        Case 3
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 4: provide a destination folder and any renaming options")
            
            sideText = g_Language.TranslateMessage("In this final step, PhotoDemon needs to know where to save the processed images, and what name to give the new files.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("For your convenience, a number of standard renaming options are also provided.  Note that all items under ""additional rename options"" are optional.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Finally, if two or more images in the batch list have the same filename, and the ""original filenames"" option is selected, such files will automatically be given unique filenames upon saving (e.g. ""original-filename (2)"").")
        
        'Step 5: process!
        Case 4
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 5: wait for batch processing to finish")
            
            sideText = g_Language.TranslateMessage("Batch processing is now underway.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Once the batch processor has processed several images, it will display an estimated time remaining.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("You can cancel batch processing at any time by pressing the ""Cancel"" button in the bottom-right corner.  If you choose to cancel, any processed images will still be present in the output folder, so you may need to remove them manually.")
            
    End Select
    
    lblExplanation(0).Caption = sideText
    
End Sub

'Remove all selected items from the batch conversion list
Private Sub cmdRemove_Click()
    
    If (lstFiles.ListIndex >= 0) Then
        
        Dim prevListIndex As Long
        prevListIndex = lstFiles.ListIndex
        lstFiles.RemoveItem prevListIndex
        If (prevListIndex < lstFiles.ListCount) Then lstFiles.ListIndex = prevListIndex Else lstFiles.ListIndex = lstFiles.ListCount - 1
    
        'And if all files were removed, disable actions that require at least one image
        If (lstFiles.ListCount = 0) Then
            cmdRemoveAll.Enabled = False
            cmdSaveList.Enabled = False
        End If
        
    End If
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
    'Update the label that displays the number of items in the list
    UpdateBatchListCount
            
    If (lstFiles.ListIndex >= 0) Then UpdatePreview lstFiles.List(lstFiles.ListIndex)
            
End Sub

'Remove ALL items from the batch conversion list
Private Sub cmdRemoveAll_Click()
    
    lstFiles.Clear
    UpdatePreview ""
    
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
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "List Folder", "")
    
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
        g_UserPreferences.SetPref_String "Batch Process", "List Folder", listPath
        
        'Assemble the output string, which basically just contains the currently selected list of files.
        Dim outputText As pdString 'As String
        Set outputText = New pdString
        
        outputText.AppendLine "<PHOTODEMON BATCH CONVERSION LIST>"
        outputText.AppendLine Trim$(Str(lstFiles.ListCount))
        
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
        removeSubfolders = CBool(chkRemoveSubfolders.Value)
        
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
            
            If removeFile Then
                lstFiles.RemoveItem i
            Else
                i = i + 1
            End If
            
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
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Macro", "")
    
    Dim cdFilter As String
    cdFilter = "PhotoDemon " & g_Language.TranslateMessage("Macro Data") & " (." & MACRO_EXT & ")|*." & MACRO_EXT & ";*.thm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    'Prepare a common dialog object
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    Dim sFile As String
   
    'If the user provides a valid macro file, use that as part of the batch process
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, g_Language.TranslateMessage("Open Macro File"), "." & MACRO_EXT, Me.hWnd) Then
        
        'As a convenience to the user, save this directory as the default macro path
        tempPathString = Files.FileGetPath(sFile)
        g_UserPreferences.SetPref_String "Paths", "Macro", tempPathString
        
        'Display the selected macro location in the relevant text box
        txtMacro.Text = sFile
        
        'Also, select the macro option button by default
        chkActions(2).Value = vbChecked
        
    End If

End Sub

Private Sub cmdSelectOutputPath_Click()
    
    Dim tString As String
    tString = PathBrowseDialog(FormBatchWizard.hWnd)
    
    If (Len(tString) <> 0) Then
        txtOutputPath.Text = Files.PathAddBackslash(tString)
    
        'Save this new directory as the default path for future usage
        g_UserPreferences.SetPref_String "Batch Process", "Output Folder", tString
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
            If MainModule.IsProgramRunning() Then picResizeDemo.BackColor = g_Themer.GetGenericUIColor(UI_Background)
            cmbResizeFit.Clear
            cmbResizeFit.AddItem "stretching to fit", 0
            cmbResizeFit.AddItem "fit inclusively", 1
            cmbResizeFit.AddItem "fit exclusively", 2
            cmbResizeFit.ListIndex = 0
        
        'For convenience, change the default resize width and height to the current screen resolution
            ucResize.SetInitialDimensions Screen.Width / TwipsPerPixelXFix, Screen.Height / TwipsPerPixelYFix
            
        'By default, select "apply no photo editing actions"
            For i = 0 To chkActions.Count - 1
                chkActions(i).Value = vbUnchecked
            Next i
                
    'Populate all file-format-related combo boxes, tooltips, and options
        m_ExportSettingsSet = False
        For i = 0 To g_ImageFormats.GetNumOfOutputFormats()
            cmbOutputFormat.AddItem g_ImageFormats.GetOutputFormatDescription(i), i
        Next i
        
        'Save JPEGs by default
        For i = 0 To cmbOutputFormat.ListCount
            If (StrComp(LCase$(g_ImageFormats.GetOutputFormatExtension(i)), "jpg", vbBinaryCompare) = 0) Then
                cmbOutputFormat.ListIndex = i
                Exit For
            End If
        Next i
    
    'Build default paths from preference file values
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "Output Folder", "")
    If (Len(tempPathString) <> 0) And (Files.PathExists(tempPathString)) Then txtOutputPath.Text = tempPathString Else txtOutputPath.Text = g_UserPreferences.GetPref_String("Paths", "Save Image", "")
    
    'By default, offer to save processed images in their original format
    optFormat(0).Value = True
    
    'Populate the combo box for file rename options
    cmbOutputOptions.AddItem "Original filenames"
    cmbOutputOptions.AddItem "Ascending numbers (1, 2, 3, etc.)"
    cmbOutputOptions.ListIndex = 0
        
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    Dim btnIconSize As Long
    btnIconSize = FixDPI(32)
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
        
        If Files.FileExists(targetFile) Then
            cmdRemove.Enabled = True
            UpdatePreview targetFile
        Else
            cmdRemove.Enabled = False
            lstFiles.RemoveItem lstFiles.ListIndex
        End If
        
    End If
    
End Sub

'Update the active image preview in the top-right
Private Sub UpdatePreview(ByVal srcImagePath As String, Optional ByVal forceUpdate As Boolean = False)
    
    lblCurrentFile.Caption = srcImagePath
    
    'Only redraw the preview if it doesn't match the last image we previewed
    If (CBool(chkEnablePreview) And (Strings.StringsNotEqual(m_CurImagePreview, srcImagePath, True) Or forceUpdate)) Then
    
        'Use PD's central load function to load a copy of the requested image
        Dim tmpDIB As pdDIB: Set tmpDIB = New pdDIB
        Dim loadSuccessful As Boolean: loadSuccessful = False
        If (Len(srcImagePath) <> 0) Then loadSuccessful = Loading.QuickLoadImageToDIB(srcImagePath, tmpDIB, False, False, True)
        
        'If the image load failed, display a placeholder message; otherwise, render the image to the picture box
        If loadSuccessful Then
            tmpDIB.RenderToPictureBox picPreview
        Else
            picPreview.Picture = LoadPicture("")
            Dim strToPrint As String
            strToPrint = g_Language.TranslateMessage("Preview not available")
            picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.textWidth(strToPrint)) \ 2
            picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.textHeight(strToPrint)) \ 2
            picPreview.Print strToPrint
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
    chkRenameCase.Value = vbChecked
End Sub

'When the user presses "Start Conversion", this routine is triggered.
Private Sub PrepareForBatchConversion()

    BatchConvertMessage g_Language.TranslateMessage("Preparing batch processing engine...")
    
    'Display the progress panel
    Dim i As Long
    
    picContainer(picContainer.Count - 1).Visible = True
    
    For i = 0 To picContainer.Count - 2
        picContainer(i).Visible = False
    Next i
    
    'Hide the back/forward buttons
    cmdPrevious.Visible = False
    cmdNext.Visible = False
    
    'Let the rest of the program know that batch processing has begun
    Macros.SetMacroStatus MacroBATCH
    
    Dim curBatchFile As Long
    Dim tmpFilename As String, tmpFileExtension As String
    
    Dim totalNumOfFiles As Long
    totalNumOfFiles = lstFiles.ListCount
    
    'Prepare the folder that will receive the processed images
    Dim outputPath As String
    outputPath = Files.PathAddBackslash(txtOutputPath)
    If (Not Files.PathExists(outputPath)) Then Files.PathCreate outputPath, True
    
    'Prepare the progress bar, which will keep the user updated on our progress.
    Set sysProgBar = New cProgressBarOfficial
    sysProgBar.CreateProgressBar picBatchProgress.hWnd, 0, 0, picBatchProgress.ScaleWidth, picBatchProgress.ScaleHeight, True, True, True, True
    sysProgBar.Max = totalNumOfFiles
    sysProgBar.Min = 0
    sysProgBar.Value = 0
    sysProgBar.Refresh
    
    'Let's also give the user an estimate of how long this is going to take.  We estimate time by determining an
    ' approximate "time-per-image" value, then multiplying that by the number of images remaining.  The progress bar
    ' will display this, automatically updated, as each image is completed.
    Dim timeStarted As Double, timeElapsed As Double, timeRemaining As Double, timePerFile As Double
    Dim numFilesProcessed As Long, numFilesRemaining As Long
    Dim minutesRemaining As Long, secondsRemaining As Long
    Dim timeMsg As String
    Dim lastTimeCalculation As Long
    lastTimeCalculation = &H7FFFFFFF
    
    timeStarted = GetTickCount
    timeMsg = ""
    
    'This is where the fun begins.  Loop through every file in the list, and process them one-by-one using the options requested
    ' by the user.
    For curBatchFile = 0 To totalNumOfFiles
    
        'Pause for keypresses - this allows the user to press "Escape" to cancel the operation
        DoEvents
        If (Macros.GetMacroStatus = MacroCANCEL) Then GoTo MacroCanceled
    
        tmpFilename = lstFiles.List(curBatchFile)
        
        'Give the user a progress update
        BatchConvertMessage g_Language.TranslateMessage("Processing image # %1 of %2. %3", (curBatchFile + 1), totalNumOfFiles, timeMsg)
        sysProgBar.Value = curBatchFile
        sysProgBar.Refresh
        
        'As a failsafe, check to make sure the current input file exists before attempting to load it
        If Files.FileExists(tmpFilename) Then
            
            'Check to see if the image file is a multipage file
            Dim howManyPages As Long
            howManyPages = Plugin_FreeImage.IsMultiImage(tmpFilename)
            
            'TODO: integrate this with future support for exporting multipage files.  At present, to avoid complications,
            ' PD will only load the first page/frame of a multipage file during conversion.
            
            'Load the current image
            If Loading.LoadFileAsNewImage(tmpFilename, , False, True, False) Then
            
                'With the image loaded, it is time to apply any requested photo editing actions.
                If (btsPhotoOps.ListIndex = 1) Then
                
                    'If the user has requested automatic lighting fixes, apply it now
                    If CBool(chkActions(0)) Then
                        Process "White balance", , TextSupport.BuildParamList("threshold", "0.1"), UNDO_Layer
                    End If
                
                    'If the user has requested an image resize, apply it now
                    If CBool(chkActions(1)) Then
                        
                        Dim resizeParams As pdParamXML
                        Set resizeParams = New pdParamXML
                        With resizeParams
                            .AddParam "width", ucResize.ResizeWidth
                            .AddParam "height", ucResize.ResizeHeight
                            .AddParam "unit", ucResize.UnitOfMeasurement
                            .AddParam "ppi", ucResize.ResizeDPIAsPPI
                            .AddParam "algorithm", ResizeSincLanczos
                            .AddParam "fit", cmbResizeFit.ListIndex
                            .AddParam "fillcolor", vbWhite
                            .AddParam "target", PD_AT_WHOLEIMAGE
                        End With
                        
                        Process "Resize image", , resizeParams.GetParamString
                        
                    End If
                    
                    'If the user has requested a macro, play it now
                    If CBool(chkActions(2)) Then Macros.PlayMacroFromFile txtMacro
                    
                End If
                
                'With the macro complete, prepare the file for saving
                tmpFilename = Files.FileGetName(lstFiles.List(curBatchFile), True)
                
                'Build a full file path using the options the user specified
                If (cmbOutputOptions.ListIndex = 0) Then
                    If CBool(chkRenamePrefix) Then tmpFilename = txtAppendFront & tmpFilename
                    If CBool(chkRenameSuffix) Then tmpFilename = tmpFilename & txtAppendBack
                Else
                    tmpFilename = curBatchFile + 1
                    If CBool(chkRenamePrefix) Then tmpFilename = txtAppendFront & tmpFilename
                    If CBool(chkRenameSuffix) Then tmpFilename = tmpFilename & txtAppendBack
                End If
                
                'If requested, remove any specified text from the filename
                If CBool(chkRenameRemove) And (Len(txtRenameRemove) <> 0) Then
                
                    'Use case-sensitive or case-insensitive matching as requested
                    If CBool(chkRenameCaseSensitive) Then
                        If InStr(1, tmpFilename, txtRenameRemove, vbBinaryCompare) Then
                            tmpFilename = Replace(tmpFilename, txtRenameRemove, "", , , vbBinaryCompare)
                        End If
                    Else
                        If InStr(1, tmpFilename, txtRenameRemove, vbTextCompare) Then
                            tmpFilename = Replace(tmpFilename, txtRenameRemove, "", , , vbTextCompare)
                        End If
                    End If
                    
                End If
                
                'Replace spaces with underscores if requested
                If CBool(chkRenameSpaces) Then
                    If InStr(1, tmpFilename, " ") Then
                        tmpFilename = Replace(tmpFilename, " ", "_")
                    End If
                End If
                
                'Change the full filename's case if requested
                If CBool(chkRenameCase) Then
                    If optCase(0) Then tmpFilename = LCase(tmpFilename) Else tmpFilename = UCase(tmpFilename)
                End If
                
                'Attach a proper image format file extension and save format ID number based off the user's
                ' requested output format
                
                'Possibility 1: use original file format
                If optFormat(0) Then
                    
                    m_FormatParams = ""
                    
                    'See if this image's file format is supported by the export engine
                    If (g_ImageFormats.GetIndexOfOutputPDIF(pdImages(g_CurrentImage).GetCurrentFileFormat) = -1) Then
                        
                        'The current format isn't supported.  Use PNG as it's the best compromise of
                        ' lossless, well-supported, and reasonably well-compressed.
                        tmpFileExtension = g_ImageFormats.GetExtensionFromPDIF(PDIF_PNG)
                        pdImages(g_CurrentImage).SetCurrentFileFormat PDIF_PNG
                        
                    Else
                        
                        'This format IS supported, so use the default extension
                        tmpFileExtension = g_ImageFormats.GetExtensionFromPDIF(pdImages(g_CurrentImage).GetCurrentFileFormat)
                    
                    End If
                    
                'Possibility 2: force all images to a single file format
                Else
                    tmpFileExtension = g_ImageFormats.GetOutputFormatExtension(cmbOutputFormat.ListIndex)
                    pdImages(g_CurrentImage).SetCurrentFileFormat g_ImageFormats.GetOutputPDIF(cmbOutputFormat.ListIndex)
                End If
                
                'If the user has requested lower- or upper-case, we now need to convert the extension as well
                If CBool(chkRenameCase) Then
                    If optCase(0) Then tmpFileExtension = LCase$(tmpFileExtension) Else tmpFileExtension = UCase$(tmpFileExtension)
                End If
                
                'Because removing specified text from filenames may lead to files with the same name, call the incrementFilename
                ' function to find a unique filename of the "filename (n+1)" variety if necessary.  This will also prepend the
                ' drive and directory structure.
                tmpFilename = outputPath & Files.IncrementFilename(outputPath, tmpFilename, tmpFileExtension) & "." & tmpFileExtension
                                
                'Request a save from the PhotoDemon_SaveImage method, and pass it the parameter string created by the user
                ' on the matching wizard panel.
                ' TODO: track success/fail results and collate any failures into a list that we can report to the user
                Saving.PhotoDemon_BatchSaveImage pdImages(g_CurrentImage), tmpFilename, pdImages(g_CurrentImage).GetCurrentFileFormat, m_ExportSettingsFormat, m_ExportSettingsMetadata
                
                'Unload the finished image
                CanvasManager.FullPDImageUnload g_CurrentImage, (Not (curBatchFile < totalNumOfFiles - 1))
            
            End If
            
            'If a good number of images have been processed, start estimating the amount of time remaining
            If (curBatchFile > 10) Then
            
                timeElapsed = GetTickCount - timeStarted
                numFilesProcessed = curBatchFile + 1
                numFilesRemaining = totalNumOfFiles - numFilesProcessed
                timePerFile = timeElapsed / numFilesProcessed
                timeRemaining = timePerFile * numFilesRemaining
                
                'Convert timeRemaining to seconds (it is currently in milliseconds)
                timeRemaining = timeRemaining / 1000#
                
                minutesRemaining = Int(timeRemaining / 60#)
                secondsRemaining = Int(timeRemaining) Mod 60
                
                'Only update the time remaining message if it is LESS than our previous result, the seconds are a multiple
                ' of 5, or there is 0 minutes remaining (in which case we can display an exact seconds estimate).
                If (timeRemaining < lastTimeCalculation) And ((secondsRemaining Mod 5 = 0) Or (minutesRemaining = 0)) Then
                
                    lastTimeCalculation = timeRemaining
                
                    'This lets us format our time nicely (e.g. "minute" vs "minutes")
                    Select Case minutesRemaining
                        'No minutes remaining - only seconds
                        Case 0
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": "
                        Case 1
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": " & minutesRemaining
                            timeMsg = timeMsg & " " & g_Language.TranslateMessage("minute") & " "
                        Case Else
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": " & minutesRemaining
                            timeMsg = timeMsg & " " & g_Language.TranslateMessage("minutes") & " "
                    End Select
                    
                    Select Case secondsRemaining
                        Case 1
                            timeMsg = timeMsg & "1 " & g_Language.TranslateMessage("second")
                        Case Else
                            timeMsg = timeMsg & secondsRemaining & " " & g_Language.TranslateMessage("seconds")
                    End Select
                
                End If

            ElseIf (curBatchFile > 20) And (totalNumOfFiles > 50) Then
                timeMsg = g_Language.TranslateMessage("Estimating time remaining") & "..."
            End If
        
        End If
                
    'Carry on
    Next curBatchFile
    
    Macros.SetMacroStatus MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    'Max out the progess bar and display a success message
    sysProgBar.Value = sysProgBar.Max
    sysProgBar.Refresh
    BatchConvertMessage g_Language.TranslateMessage("%1 files were successfully processed!", totalNumOfFiles)
    
    'Finally, there is no longer any need for the user to save their batch list, as the batch process is complete.
    m_ImageListSaved = True
    
    Exit Sub
    
MacroCanceled:

    Macros.SetMacroStatus MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Reset the progress bar
    sysProgBar.Value = 0
    sysProgBar.Refresh
    
    Dim cancelMsg As String
    cancelMsg = g_Language.TranslateMessage("Batch conversion canceled.") & " " & curBatchFile & " "
    
    'Properly display "image" or "images" depending on how many files were processed
    If (curBatchFile <> 1) Then
        cancelMsg = cancelMsg & g_Language.TranslateMessage("images were")
    Else
        cancelMsg = cancelMsg & g_Language.TranslateMessage("image was")
    End If
    
    cancelMsg = cancelMsg & " "
    cancelMsg = cancelMsg & g_Language.TranslateMessage("processed before cancelation. Last processed image was ""%1"".", lstFiles.List(curBatchFile))
    
    BatchConvertMessage cancelMsg
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    m_ImageListSaved = True
    
End Sub

'Display a progress update to the user
Private Sub BatchConvertMessage(ByVal newMessage As String)
    lblBatchProgress.Caption = newMessage
    lblBatchProgress.RequestRefresh
End Sub

Private Sub UpdatePhotoOpVisibility()
    If (btsPhotoOps.ListIndex = 0) Then
        lblExplanation(1).Visible = True
        picPhotoEdits.Visible = False
    Else
        lblExplanation(1).Visible = False
        picPhotoEdits.Visible = True
    End If
End Sub

