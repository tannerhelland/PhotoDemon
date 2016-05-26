VERSION 5.00
Begin VB.Form FormPreferences 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Options"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11505
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   42
      Top             =   6870
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonStripVertical btsvCategory 
      Height          =   6675
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   11774
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   0
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdDropDown cboMRUCaption 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboImageCaption 
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Top             =   810
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdSpinner tudRecentFiles 
         Height          =   345
         Left            =   3900
         TabIndex        =   5
         Top             =   3930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         DefaultValue    =   10
         Min             =   1
         Max             =   32
         Value           =   10
      End
      Begin PhotoDemon.pdCheckBox chkMouseHighResolution 
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   582
         Caption         =   "use high-resolution input tracking"
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   1440
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "mouse and pen input"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblRecentFileCount 
         Height          =   240
         Left            =   240
         Top             =   3960
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   423
         Caption         =   "maximum number of recent file entries: "
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   13
         Left            =   0
         Top             =   2520
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "recent files list"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblMRUText 
         Height          =   240
         Left            =   240
         Top             =   3000
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   503
         Caption         =   "recently used file shortcuts should be: "
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "captions"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblImageText 
         Height          =   240
         Left            =   240
         Top             =   450
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   423
         Caption         =   "image window titles should be: "
         ForeColor       =   4210752
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   4
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   7
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdButton cmdColorProfilePath 
         Height          =   375
         Left            =   7380
         TabIndex        =   39
         Top             =   2760
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   661
         Caption         =   "..."
      End
      Begin PhotoDemon.pdDropDown cboAlphaCheckSize 
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   5250
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboAlphaCheck 
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   4260
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboMonitors 
         Height          =   330
         Left            =   780
         TabIndex        =   12
         Top             =   1950
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdTextBox txtColorProfilePath 
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   2790
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         Text            =   "(none)"
      End
      Begin PhotoDemon.pdRadioButton optColorManagement 
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "use the system color profile"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdColorSelector csAlphaOne 
         Height          =   435
         Left            =   6240
         TabIndex        =   15
         Top             =   4230
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
         ShowMainWindowColor=   0   'False
      End
      Begin PhotoDemon.pdColorSelector csAlphaTwo 
         Height          =   435
         Left            =   7320
         TabIndex        =   16
         Top             =   4230
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
         ShowMainWindowColor=   0   'False
      End
      Begin PhotoDemon.pdRadioButton optColorManagement 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "use one or more custom color profiles"
      End
      Begin PhotoDemon.pdLabel lblColorManagement 
         Height          =   240
         Index           =   2
         Left            =   780
         Top             =   2430
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   503
         Caption         =   "color profile for selected monitor:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblColorManagement 
         Height          =   240
         Index           =   1
         Left            =   780
         Top             =   1590
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         Caption         =   "available monitors:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblColorManagement 
         Height          =   240
         Index           =   0
         Left            =   240
         Top             =   480
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   503
         Caption         =   "when rendering images to the screen:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   503
         Caption         =   "color management"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblAlphaCheckSize 
         Height          =   240
         Left            =   240
         Top             =   4860
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Caption         =   "transparency checkerboard size:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblAlphaCheck 
         Height          =   240
         Left            =   240
         Top             =   3870
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   503
         Caption         =   "transparency checkerboard colors:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   3420
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   503
         Caption         =   "transparency management"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   2
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   18
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdCheckBox chkSaveMetadata 
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   4560
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "copy any relevant metadata to the new file"
      End
      Begin PhotoDemon.pdCheckBox chkConfirmUnsaved 
         Height          =   330
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "when closing images, warn me me about unsaved changes"
      End
      Begin PhotoDemon.pdDropDown cboDefaultSaveFormat 
         Height          =   330
         Left            =   240
         TabIndex        =   20
         Top             =   3135
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboSaveBehavior 
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   1725
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   3
         Left            =   240
         Top             =   4140
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   503
         Caption         =   "default metadata behavior on first save:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   3690
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   503
         Caption         =   "metadata"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   16
         Left            =   0
         Top             =   885
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         Caption         =   "save behavior: overwrite vs make a copy"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   2
         Left            =   240
         Top             =   1335
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Caption         =   "when ""Save"" is used:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "closing unsaved images"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   1
         Left            =   240
         Top             =   2730
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   503
         Caption         =   "when using the ""Save As"" command, set the default file format according to:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   10
         Left            =   0
         Top             =   2310
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   503
         Caption         =   "default file format when saving"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdCheckBox chkAnonymizeMetadata 
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   4920
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "erase tags that might be personal (GPS, camera details, etc)"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkMetadataListPD 
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   5280
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "list PhotoDemon as the last-used editing software"
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   1
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   11
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdCheckBox chkToneMapping 
         Height          =   330
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "display tone mapping options when importing HDR and RAW images"
      End
      Begin PhotoDemon.pdCheckBox chkLoadingOrientation 
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "obey auto-rotate instructions inside image files"
      End
      Begin PhotoDemon.pdDropDown cboLargeImages 
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   4680
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   3
         Left            =   60
         Top             =   3000
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "orientation"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   6
         Left            =   60
         Top             =   0
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   503
         Caption         =   "high-dynamic range (HDR) images"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   5
         Left            =   60
         Top             =   3960
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "zoom"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblImgOpen 
         Height          =   240
         Left            =   240
         Top             =   4320
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   503
         Caption         =   "when an image is first loaded, set its viewport zoom to: "
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   0
         Left            =   60
         Top             =   960
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   503
         Caption         =   "metadata"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdCheckBox chkMetadataBinary 
         Height          =   330
         Left            =   240
         TabIndex        =   45
         Top             =   2400
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "forcibly extract binary-type tags as Base64 (slow)"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkMetadataJPEG 
         Height          =   330
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "estimate original JPEG quality settings"
      End
      Begin PhotoDemon.pdCheckBox chkMetadataUnknown 
         Height          =   330
         Left            =   240
         TabIndex        =   47
         Top             =   2040
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "extract unknown tags"
         Value           =   0
      End
      Begin PhotoDemon.pdCheckBox chkMetadataDuplicates 
         Height          =   330
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "automatically hide duplicate tags"
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   3
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdSlider sltUndoCompression 
         Height          =   405
         Left            =   240
         TabIndex        =   25
         Top             =   5730
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   873
         Max             =   9
         SliderTrackStyle=   1
         NotchPosition   =   2
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   720
         Width           =   7920
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   1980
         Width           =   7920
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   3240
         Width           =   7920
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   29
         Top             =   4470
         Width           =   7920
         _ExtentX        =   14076
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdLabel lblPerformanceTitle 
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "color management"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceSub 
         Height          =   240
         Index           =   4
         Left            =   180
         Top             =   390
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   503
         Caption         =   "when calculating color values:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceSub 
         Height          =   240
         Index           =   3
         Left            =   180
         Top             =   1650
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   503
         Caption         =   "when decorating interface elements:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceTitle 
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   1260
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   503
         Caption         =   "interface"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPNGCompression 
         Height          =   240
         Index           =   3
         Left            =   360
         Top             =   6240
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         Caption         =   "no compression (fastest)"
         FontSize        =   8
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPNGCompression 
         Height          =   240
         Index           =   2
         Left            =   4020
         Top             =   6240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "maximum compression (slowest)"
         FontSize        =   8
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceSub 
         Height          =   240
         Index           =   2
         Left            =   240
         Top             =   5370
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Caption         =   "compress undo/redo data at the following level:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceTitle 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   4980
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         Caption         =   "undo/redo"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceTitle 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   2520
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "thumbnails"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceSub 
         Height          =   240
         Index           =   1
         Left            =   180
         Top             =   2910
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   503
         Caption         =   "when generating image and layer thumbnail images:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceSub 
         Height          =   240
         Index           =   0
         Left            =   180
         Top             =   4140
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   503
         Caption         =   "when rendering the image canvas:"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPerformanceTitle 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   3750
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   503
         Caption         =   "viewport"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   6
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdButton cmdReset 
         Height          =   600
         Left            =   240
         TabIndex        =   41
         Top             =   6000
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1058
         Caption         =   "Click here to reset all options"
      End
      Begin PhotoDemon.pdButton cmdTmpPath 
         Height          =   450
         Left            =   7680
         TabIndex        =   40
         Top             =   435
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   794
         Caption         =   "..."
      End
      Begin PhotoDemon.pdButtonToolbox cmdCopyReportClipboard 
         Height          =   570
         Left            =   7650
         TabIndex        =   36
         Top             =   3315
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtHardware 
         Height          =   1785
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3942
         FontSize        =   9
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtTempPath 
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   510
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         Text            =   "automatically generated at run-time"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   6
         Left            =   0
         Top             =   1560
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   503
         Caption         =   "hardware acceleration report:"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblMemoryUsageMax 
         Height          =   540
         Left            =   240
         Top             =   4980
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   953
         Caption         =   "memory usage will be displayed here"
         ForeColor       =   8405056
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblMemoryUsageCurrent 
         Height          =   540
         Left            =   240
         Top             =   4440
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   953
         Caption         =   "memory usage will be displayed here"
         ForeColor       =   8405056
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   5
         Left            =   0
         Top             =   4080
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "memory diagnostics"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblRuntimeSettings 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "temporary file location"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTempPathWarning 
         Height          =   600
         Left            =   240
         Top             =   900
         Visible         =   0   'False
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   1058
         ForeColor       =   255
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   5520
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "start over"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   5
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   444
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      Begin PhotoDemon.pdLabel lblExplanation 
         Height          =   2535
         Left            =   240
         Top             =   3840
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         Caption         =   "(disclaimer populated at run-time)"
         FontSize        =   9
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkUpdates 
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
         Caption         =   "update language files independently"
      End
      Begin PhotoDemon.pdLabel lblUpdates 
         Height          =   240
         Index           =   0
         Left            =   240
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Caption         =   "automatically check for updates:"
      End
      Begin PhotoDemon.pdDropDown cboUpdates 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdCheckBox chkUpdates 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
         Caption         =   "update plugins independently"
      End
      Begin PhotoDemon.pdDropDown cboUpdates 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   1710
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblUpdates 
         Height          =   240
         Index           =   1
         Left            =   240
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Caption         =   "allow updates from these tracks:"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         Caption         =   "update options"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdCheckBox chkUpdates 
         Height          =   330
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   3240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
         Caption         =   "notify me when an update is ready"
      End
   End
End
Attribute VB_Name = "FormPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright 2002-2016 by Tanner Helland
'Created: 8/November/02
'Last updated: 05/May/16
'Last update: purge the entire "file formats" section; these preferences are now supplied by export-time dialogs
'
'Dialog for interfacing with the user's desired program preferences.  Handles reading/writing from/to the persistent
' XML file that actually stores all preferences.
'
'Note that this form interacts heavily with the pdPreferences class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Private userInitiatedColorSelection As Boolean
Private userInitiatedAlphaSelection As Boolean

'This dialog interacts heavily with various system-level bits.  pdSystemInfo retrieves this data for us.
Private cSysInfo As pdSystemInfo

Private Sub btsvCategory_Click(ByVal buttonIndex As Long)

    'When the preferences category is changed, only display the controls in that category
    Dim catID As Long
    For catID = 0 To btsvCategory.ListCount - 1
        
        If catID = buttonIndex Then
            picContainer(catID).Visible = True
            If Me.Visible Then picContainer(catID).SetFocus
        Else
            picContainer(catID).Visible = False
        End If
        
    Next catID

End Sub

'Alpha channel checkerboard selection; change the color selectors to match
Private Sub cboAlphaCheck_Click()

    'Only respond to user-generated events
    If userInitiatedAlphaSelection Then

        userInitiatedAlphaSelection = False

        'Redraw the sample picture boxes based on the value the user has selected
        Select Case cboAlphaCheck.ListIndex
        
            'Case 0 - Highlights
            Case 0
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(204, 204, 204)
            
            'Case 1 - Midtones
            Case 1
                csAlphaOne.Color = RGB(153, 153, 153)
                csAlphaTwo.Color = RGB(102, 102, 102)
            
            'Case 2 - Shadows
            Case 2
                csAlphaOne.Color = RGB(51, 51, 51)
                csAlphaTwo.Color = RGB(0, 0, 0)
            
            'Case 3 - Custom
            Case 3
                csAlphaOne.Color = RGB(255, 204, 246)
                csAlphaTwo.Color = RGB(255, 255, 255)
            
        End Select
        
        userInitiatedAlphaSelection = True
                
    End If

End Sub

'Whenever the Color and Transparency -> Color Management -> Monitor combo box is changed, load the relevant color profile
' path from the preferences file (if one exists)
Private Sub cboMonitors_Click()

    'One of the difficulties with tracking multiple monitors is that the user can attach/detach them at will.  They
    ' can also have multiple monitors attached with the same make and model (and retrieving EDIDs is extremely
    ' unpleasant and overwrought).  Per this article (http://www.microsoft.com/msj/0697/monitor/monitor.aspx), PD
    ' uses the HMONITOR handle to store and retrieve monitor-specific settings, because as the article states,
    ' "A physical device has the same HMONITOR value throughout its lifetime, even across changes to display settings,
    ' as long as it remains a part of the desktop."  If the user runs PD without a monitor attached, only to reattach
    ' it later, I have no idea if a new HMONITOR will be assigned or not... and at present, that's frankly not a
    ' huge concern for me.   We'll leave that level of management up to the OS.  For now, just assume that HMONITOR
    ' is a valid way to persistently track individual monitors.

    'Start by retrieving the HMONITOR value for the selected monitor
    Dim hMonitor As Long
    If Not (g_Displays.Displays(cboMonitors.ListIndex) Is Nothing) Then hMonitor = g_Displays.Displays(cboMonitors.ListIndex).GetHandle
    
    'Use that to retrieve a stored color profile (if any)
    Dim profilePath As String
    profilePath = g_UserPreferences.GetPref_String("Transparency", "MonitorProfile_" & hMonitor, "(none)")
    
    'If the returned value is "(none)", translate that into the user's language before displaying; otherwise, display
    ' whatever path we retrieved.
    If profilePath = "(none)" Then
        txtColorProfilePath.Text = g_Language.TranslateMessage("(none)")
    Else
        txtColorProfilePath.Text = profilePath
    End If
    
End Sub

Private Sub chkAnonymizeMetadata_Click()
    If CBool(chkAnonymizeMetadata.Value) Then
        chkMetadataListPD.Value = vbUnchecked
        chkMetadataListPD.Enabled = False
    Else
        chkMetadataListPD.Enabled = True
    End If
End Sub

Private Sub cmdBarMini_OKClick()

    Message "Saving preferences..."
    Me.Visible = False
    
    'After updates on 22 Oct 2014, the preference saving sequence should happen in a flash, but just in case,
    ' we'll supply a bit of processing feedback.
    FormMain.Enabled = False
    SetProgBarMax 8
    SetProgBarVal 1
    
    'Start batch preference edit mode
    g_UserPreferences.StartBatchPreferenceMode
    
    'First, make note of the active panel, so we can default to that if the user returns to this dialog
    g_UserPreferences.SetPref_Long "Core", "Last Preferences Page", btsvCategory.ListIndex
    
    'Write preferences out to file in category order.  (The preference XML file is order-agnostic, but I try to
    ' maintain the order used in the Preferences dialog itself to make changes easier.)
    
    '***************************************************************************
    
    'BEGIN Interface preferences
        
        'START/END image window caption length
            g_UserPreferences.SetPref_Long "Interface", "Window Caption Length", cboImageCaption.ListIndex
        
        'START/END high-res input tracking
            g_UserPreferences.SetPref_Boolean "Interface", "High Resolution Input", CBool(chkMouseHighResolution.Value)
            g_HighResolutionInput = CBool(chkMouseHighResolution.Value)
        
        Dim mruNeedsToBeRebuilt As Boolean
        mruNeedsToBeRebuilt = False
        
        'START MRU caption length
        
            'Check to see if the new MRU caption setting matches the old one.  If it doesn't, reload the MRU.
            If cboMRUCaption.ListIndex <> g_UserPreferences.GetPref_Long("Interface", "MRU Caption Length", 0) Then mruNeedsToBeRebuilt = True
            g_UserPreferences.SetPref_Long "Interface", "MRU Caption Length", cboMRUCaption.ListIndex
            
        'END MRU caption length
        
        'START maximum MRU count
            Dim newMaxRecentFiles As Long
            
            'Validate the user's supplied recent file limit
            If tudRecentFiles.IsValid Then
                newMaxRecentFiles = tudRecentFiles.Value
            Else
                newMaxRecentFiles = 10
            End If
            
            'If the max number of recent files has changed, update the MRU list to match
            If newMaxRecentFiles <> g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then mruNeedsToBeRebuilt = True
            g_UserPreferences.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value
            
        'END maximum MRU count
        
        'If the MRU needs to be rebuilt, do so now
        If mruNeedsToBeRebuilt Then
            g_RecentFiles.MRU_NotifyNewMaxLimit
            ' Recent files limit applies to macros as well
            g_RecentMacros.MRU_NotifyNewMaxLimit
        End If
        
    
    'END Interface preferences
    
    '***************************************************************************
    
    SetProgBarVal 2
    
    'BEGIN Loading preferences
    
        'START/END automatically tone-map HDR images
            g_UserPreferences.SetPref_Boolean "Loading", "Tone Mapping Prompt", CBool(chkToneMapping)
            
        'START metadata behavior at load-time
            g_UserPreferences.SetPref_Boolean "Loading", "Metadata Hide Duplicates", CBool(chkMetadataDuplicates.Value)
            g_UserPreferences.SetPref_Boolean "Loading", "Metadata Estimate JPEG", CBool(chkMetadataJPEG.Value)
            g_UserPreferences.SetPref_Boolean "Loading", "Metadata Extract Binary", CBool(chkMetadataBinary.Value)
            g_UserPreferences.SetPref_Boolean "Loading", "Metadata Extract Unknown", CBool(chkMetadataUnknown.Value)
        'END metadata behavior at load-time
        
        'START/END EXIF auto-rotation
            g_UserPreferences.SetPref_Boolean "Loading", "ExifAutoRotate", CBool(chkLoadingOrientation)
        
        'START initial zoom
            g_AutozoomLargeImages = cboLargeImages.ListIndex
            g_UserPreferences.SetPref_Long "Loading", "Initial Image Zoom", g_AutozoomLargeImages
        'END initial zoom
    
    
    'END Loading preferences
    
    '***************************************************************************
    
    SetProgBarVal 3
    
    'BEGIN Saving preferences
    
        'START prompt on unsaved images
            g_ConfirmClosingUnsaved = CBool(chkConfirmUnsaved.Value)
            g_UserPreferences.SetPref_Boolean "Saving", "Confirm Closing Unsaved", g_ConfirmClosingUnsaved
    
            If g_ConfirmClosingUnsaved Then
                toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "If the current image has not been saved, you will receive a prompt to save it before it closes.", "Close the current image"
            Else
                toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.", "Close the current image"
            End If
            
            
    
        'END prompt on unsaved images
        
        'START metadata-related options
            g_UserPreferences.SetPref_Boolean "Saving", "PreserveMetadata", CBool(chkSaveMetadata.Value)
            g_UserPreferences.SetPref_Boolean "Saving", "AnonymizeMetadata", CBool(chkAnonymizeMetadata.Value)
            g_UserPreferences.SetPref_Boolean "Saving", "MetadataListPD", CBool(chkMetadataListPD.Value)
        'END metadata-related options
        
        'START/END Save behavior (overwrite or copy)
            g_UserPreferences.SetPref_Long "Saving", "Overwrite Or Copy", cboSaveBehavior.ListIndex
        
        'START/END "Save As" dialog's suggested file format
            g_UserPreferences.SetPref_Long "Saving", "Suggested Format", cboDefaultSaveFormat.ListIndex
    
    'END Saving preferences
        
    '***************************************************************************
    
    SetProgBarVal 4
    
    'START Performance preferences
    
        'START/END color management performance
            g_UserPreferences.SetPref_Long "Performance", "Color Performance", cboPerformance(0).ListIndex
            g_ColorPerformance = cboPerformance(0).ListIndex
    
        'START/END interface decoration performance
            g_UserPreferences.SetPref_Long "Performance", "Interface Decoration Performance", cboPerformance(1).ListIndex
            g_InterfacePerformance = cboPerformance(1).ListIndex
        
        'START/END thumbnail render performance
            g_UserPreferences.SetPref_Long "Performance", "Thumbnail Performance", cboPerformance(2).ListIndex
            g_ThumbnailPerformance = cboPerformance(2).ListIndex
        
        'START/END viewport render performance
            g_UserPreferences.SetPref_Long "Performance", "Viewport Render Performance", cboPerformance(3).ListIndex
            g_ViewportPerformance = cboPerformance(3).ListIndex
            
        'START/END undo/redo data compression
            g_UserPreferences.SetPref_Long "Performance", "Undo Compression", sltUndoCompression.Value
            g_UndoCompressionLevel = sltUndoCompression.Value
    
    'END Performance preferences
    
    '***************************************************************************
    
    SetProgBarVal 5
    
    'START Color and Transparency preferences

        'START use system color profile
            g_UserPreferences.SetPref_Boolean "Transparency", "Use System Color Profile", optColorManagement(0)
            g_UseSystemColorProfile = optColorManagement(0)
            CacheCurrentSystemColorProfile
            ColorManagement.CheckParentMonitor False, True
        'END use system color profile

        'START alpha checkerboard colors
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Mode", CLng(cboAlphaCheck.ListIndex)
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check One", CLng(csAlphaOne.Color)
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Two", CLng(csAlphaTwo.Color)
        'END alpha checkerboard colors
            
        'START alpha checkerboard size
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Size", cboAlphaCheckSize.ListIndex
            
            'Recreate the cached pattern for the alpha background
            Drawing.CreateAlphaCheckerboardDIB g_CheckerboardPattern
            
        'END alpha checkerboard size
    
    'END Color and Transparency preferences
    
    '***************************************************************************
    
    SetProgBarVal 6
    
    'BEGIN Update preferences
        
        'START/END update frequency
            g_UserPreferences.SetPref_Long "Updates", "Update Frequency", cboUpdates(0).ListIndex
        
        'START/END update track
            g_UserPreferences.SetPref_Long "Updates", "Update Track", cboUpdates(1).ListIndex
        
        'START update language files independently
            g_UserPreferences.SetPref_Boolean "Updates", "Update Languages Independently", CBool(chkUpdates(0).Value)
            
        'START update plugins independently
            g_UserPreferences.SetPref_Boolean "Updates", "Update Plugins Independently", CBool(chkUpdates(1).Value)
            
        'START update notifications
            g_UserPreferences.SetPref_Boolean "Updates", "Update Notifications", CBool(chkUpdates(2).Value)
    
    'END Update preferences
    
    '***************************************************************************
    
    SetProgBarVal 7
    
    'BEGIN Advanced preferences
    
        'START/END store the temporary path (but only if it's changed)
            If LCase(txtTempPath) <> LCase(g_UserPreferences.GetTempPath) Then g_UserPreferences.SetTempPath txtTempPath
    
    'END Advanced preferences
    
    '***************************************************************************
    
    'End batch preference edit mode, which will force a write-to-file operation
    g_UserPreferences.EndBatchPreferenceMode
    
    'All user preferences have now been written out to file
    
    'Because some preferences affect the program's interface, redraw the active image.
    FormMain.UpdateMainLayout
        
    'TODO: the image tabstrip should be handled by the program canvas as of 7.0's release.  See if we can remove this line.
    Interface.RequestTabstripRedraw
    
    SetProgBarVal 0
    ReleaseProgressBar
    FormMain.Enabled = True
    
    Message "Preferences updated."
    
End Sub

'Allow the user to select a new color profile for the attached monitor.  Because this text box is re-used for multiple
' settings, save any changes to file immediately, rather than waiting for the user to click OK.
Private Sub cmdColorProfilePath_Click()

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    Dim sFile As String
    sFile = ""
    
    'Get the last color profile path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Color Profile", "")
    
    'If no color profile path was found, populate it with the default system color profile path
    If Len(tempPathString) = 0 Then tempPathString = GetSystemColorFolder()
    
    'Prepare a common dialog filter list with extensions of known profile types
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("ICC Profiles") & " (.icc, .icm)|*.icc;*.icm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Please select a color profile")
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, cdTitle, ".icc", FormPreferences.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPref_String "Paths", "Color Profile", listPath
        
        'Set the text box to match this color profile, and save the resulting preference out to file.
        txtColorProfilePath = sFile
        
        Dim hMonitor As Long
        If Not g_Displays.Displays(cboMonitors.ListIndex) Is Nothing Then hMonitor = g_Displays.Displays(cboMonitors.ListIndex).GetHandle
        g_UserPreferences.SetPref_String "Transparency", "MonitorProfile_" & hMonitor, TrimNull(sFile)
        
        'If the "user custom color profiles" option button isn't selected, mark it now
        If Not optColorManagement(1).Value Then optColorManagement(1).Value = True
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput

End Sub

'Copy the hardware acceleration report to the clipboard
Private Sub cmdCopyReportClipboard_Click()
    Clipboard.Clear
    Clipboard.SetText txtHardware
End Sub

'RESET will regenerate the preferences file from scratch.  This can be an effective way to
' "reset" a copy of the program.
Private Sub cmdReset_Click()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = PDMsgBox("This action will reset all preferences to their default values.  It cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbApplicationModal + vbExclamation + vbYesNo, "Reset all preferences")

    'If the user gives final permission, rewrite the preferences file from scratch and repopulate this form
    If confirmReset = vbYes Then
        g_UserPreferences.ResetPreferences
        LoadAllPreferences
        
        'Restore the currently active language to the preferences file; this prevents the language from resetting to English
        ' (a behavior that isn't made clear by this action).
        g_Language.WriteLanguagePreferencesToFile
        
    End If

End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = BrowseForFolder(Me.hWnd)
    If Len(tString) <> 0 Then txtTempPath.Text = FixPath(tString)
End Sub

'Load all relevant values from the preferences file, and populate their corresponding controls with the user's current settings
Private Sub LoadAllPreferences()
    
    'Start batch preference mode.  This will suspend any file read/write operations until the mode finishes.
    g_UserPreferences.StartBatchPreferenceMode
    
    'For the sake of order, we will load preferences by category.  (They can be loaded in any order without consequence,
    ' but there are MANY preferences, so maintaining some kind of order is helpful.)
    
    'Note also that many tooltips are manually populated throughout this section.  This is done for translation
    ' purposes; the tooltips themselves are sometimes too long to fit inside a traditional VB control, so the
    ' IDE dumps them to a separate custom .frx resource file where they are difficult to extract. Rather than
    ' mess with that, I manually add the tooltips here so that the automatic translation engine can easily find
    ' all tooltip text.
    
    '***************************************************************************
    
    'START Interface preferences
        
        'START image window caption length
            cboImageCaption.Clear
            cboImageCaption.AddItem " compact - file name only", 0
            cboImageCaption.AddItem " descriptive - full location, including folder(s)", 1
            cboImageCaption.ListIndex = g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0)
            cboImageCaption.AssignTooltip "Image windows tend to be large, so feel free to display each image's full location in the image window title bars."
        'END image window caption length
        
        'START mouse and pen input
            If g_UserPreferences.GetPref_Boolean("Interface", "High Resolution Input", True) Then chkMouseHighResolution.Value = vbChecked Else chkMouseHighResolution.Value = vbUnchecked
            chkMouseHighResolution.AssignTooltip "High-resolution tracking allows PhotoDemon to more accurately reproduce mouse and pen movement.  On some older PCs, the system may struggle to keep up with the extra tracking data, so you can disable this if necessary."
        
        'START Recent file max count
            lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent file entries: ")
            tudRecentFiles.SetLeft lblRecentFileCount.GetLeft + lblRecentFileCount.GetWidth + FixDPI(6)
            tudRecentFiles.Value = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
        'END
        
        'START MRU caption length
            cboMRUCaption.Clear
            cboMRUCaption.AddItem " compact - file names only", 0
            cboMRUCaption.AddItem " descriptive - full locations, including folder(s)", 1
            cboMRUCaption.ListIndex = g_UserPreferences.GetPref_Long("Interface", "MRU Caption Length", 0)
            cboMRUCaption.AssignTooltip "The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations."
        'END MRU caption length

        
    'END Interface preferences
    
    '***************************************************************************
    
    'START Loading preferences
    
        'START tone-mapping HDR images at load time
            If g_UserPreferences.GetPref_Boolean("Loading", "Tone Mapping Prompt", True) Then chkToneMapping.Value = vbChecked Else chkToneMapping.Value = vbUnchecked
            
            If g_ImageFormats.FreeImageEnabled Then
                chkToneMapping.Enabled = True
            Else
                chkToneMapping.Caption = g_Language.TranslateMessage("feature disabled due to missing plugin")
                chkToneMapping.Enabled = False
            End If
            
            chkToneMapping.AssignTooltip "HDR and RAW images contain more colors than PC screens can physically display.  Before displaying such images, a tone mapping operation must be applied to the original image data."
        'END tone-mapping HDR images at load time
        
        'START metadata behavior at load-time
            If g_UserPreferences.GetPref_Boolean("Loading", "Metadata Hide Duplicates", True) Then chkMetadataDuplicates.Value = vbChecked Else chkMetadataDuplicates.Value = vbUnchecked
            If g_UserPreferences.GetPref_Boolean("Loading", "Metadata Estimate JPEG", True) Then chkMetadataJPEG.Value = vbChecked Else chkMetadataJPEG.Value = vbUnchecked
            If g_UserPreferences.GetPref_Boolean("Loading", "Metadata Extract Binary", False) Then chkMetadataBinary.Value = vbChecked Else chkMetadataBinary.Value = vbUnchecked
            If g_UserPreferences.GetPref_Boolean("Loading", "Metadata Extract Unknown", False) Then chkMetadataUnknown.Value = vbChecked Else chkMetadataUnknown.Value = vbUnchecked
        'END metadata behavior at load-time
        
        'START auto-rotate according to EXIF data
            If g_UserPreferences.GetPref_Boolean("Loading", "EXIF Auto Rotate", True) Then chkLoadingOrientation.Value = vbChecked Else chkLoadingOrientation.Value = vbUnchecked
            chkLoadingOrientation.AssignTooltip "Most digital photos include rotation instructions (EXIF orientation metadata), which PhotoDemon will use to automatically rotate photos.  Some older smartphones and cameras may not write these instructions correctly, so if your photos are being imported sideways or upside-down, you can try disabling the auto-rotate feature."
        'END auto-rotate according to EXIF data
        
        'START initial image zoom
            cboLargeImages.Clear
            cboLargeImages.AddItem " automatically fit the image on-screen", 0
            cboLargeImages.AddItem " 1:1 (100% zoom, or ""actual size"")", 1
            cboLargeImages.ListIndex = g_UserPreferences.GetPref_Long("Loading", "Initial Image Zoom", 0)
            
            cboLargeImages.AssignTooltip "Any photo larger than 2 megapixels is too big to fit on an average computer monitor.  PhotoDemon can automatically zoom out on large photographs so that the entire image is viewable."
        'END initial image zoom
    
    'END Loading preferences
    
    '***************************************************************************
    
    'START Saving preferences
    
        'START/END prompt about unsaved images
            If g_ConfirmClosingUnsaved Then chkConfirmUnsaved.Value = vbChecked Else chkConfirmUnsaved.Value = vbUnchecked
            
        'START suggested save as format
            cboDefaultSaveFormat.Clear
            cboDefaultSaveFormat.AddItem " the current file format of the image being saved", 0
            cboDefaultSaveFormat.AddItem " the last image format I used in the ""Save As"" screen", 1
            cboDefaultSaveFormat.ListIndex = g_UserPreferences.GetPref_Long("Saving", "Suggested Format", 0)
            
            cboDefaultSaveFormat.AssignTooltip "Most photo editors use the format of the current image as the default in the ""Save As"" screen.  When working with RAW images that will eventually be saved to JPEG, it is useful to have PhotoDemon remember that - hence the ""last used"" option."
        'END suggested save as format
        
        'START overwrite vs copy when saving
            cboSaveBehavior.Clear
            cboSaveBehavior.AddItem " overwrite the current file (standard behavior)", 0
            cboSaveBehavior.AddItem " save a new copy, e.g. ""filename (2).jpg"" (safe behavior)", 1
            cboSaveBehavior.ListIndex = g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0)
            
            cboSaveBehavior.AssignTooltip "In most photo editors, the ""Save"" command saves the image over its original version, erasing that copy forever.  PhotoDemon provides a ""safer"" option, where each save results in a new copy of the file."
        'END overwrite vs copy when saving
               
        'START metadata export
            If g_UserPreferences.GetPref_Boolean("Saving", "PreserveMetadata", True) Then chkSaveMetadata.Value = vbChecked Else chkSaveMetadata.Value = vbUnchecked
            If g_UserPreferences.GetPref_Boolean("Saving", "AnonymizeMetadata", False) Then chkAnonymizeMetadata.Value = vbChecked Else chkAnonymizeMetadata.Value = vbUnchecked
            If g_UserPreferences.GetPref_Boolean("Saving", "MetadataListPD", True) Then chkMetadataListPD.Value = vbChecked Else chkMetadataListPD.Value = vbUnchecked
        'END metadata export
    
    'END Saving preferences
    
    '***************************************************************************
    
    'START Performance preferences
    
        'Previously, this section was used for "tools" preferences.  PhotoDemon no longer provides dedicated tool preferences
        ' in this dialog; instead, this section is used for Performance preferences.  I have left the old Tool preference code
        ' and text here so it can be re-used in the future if tool preferences are reinstated.
        
        'START Clear selections after "Crop to Selection"
            'If g_UserPreferences.GetPref_Boolean("Tools", "Clear Selection After Crop", True) Then chkSelectionClearCrop.Value = vbChecked Else chkSelectionClearCrop.Value = vbUnchecked
            'chkSelectionClearCrop.ToolTipText = g_Language.TranslateMessage("When the ""Crop to Selection"" command is used, the resulting image will always contain a selection the same size as the full image.  There is generally no need to retain this, so PhotoDemon can automatically clear it for you.")
        'END Clear selections after "Crop to Selection"
        
        'We can shortcut a bit of initialization here by populating all quality drop-downs with the same values.
        Dim i As Long
        
        For i = 0 To cboPerformance.UBound
            cboPerformance(i).Clear
            cboPerformance(i).AddItem " maximize quality", 0
            cboPerformance(i).AddItem " balance performance and quality", 1
            cboPerformance(i).AddItem " maximize performance", 2
        Next i
        
        'START Color management accuracy v performance
            cboPerformance(0).ListIndex = g_ColorPerformance
            cboPerformance(0).AssignTooltip "Like any photo editor, PhotoDemon frequently converts colors between different reference spaces.  The accuracy of these conversions can be limited to improve performance."
        'END Color management accuracy v performance
        
        'START Interface decorations performance
            cboPerformance(1).ListIndex = g_InterfacePerformance
            cboPerformance(1).AssignTooltip "Some interface elements receive custom decorations (like drop shadows).  On older PCs, these decorations can be suspended for a small performance boost."
        'END Interface decorations performance
        
        'START Thumbnail rendering performance
            cboPerformance(2).ListIndex = g_ThumbnailPerformance
            cboPerformance(2).AssignTooltip "PhotoDemon has to generate a lot of thumbnail images, especially when images contain multiple layers.  The quality of these thumbnails can be lowered in order to improve performance."
        'END Thumbnail rendering performance
        
        'START Viewport rendering performance
            cboPerformance(3).ListIndex = g_ViewportPerformance
            cboPerformance(3).AssignTooltip "Rendering the primary image canvas is a common bottleneck for PhotoDemon's performance.  The automatic setting is recommended, but for older PCs, you can manually select the Maximize Performance option to sacrifice quality for raw performance."
        'END Viewport rendering performance
        
        'START Undo data compression
            sltUndoCompression.ToolTipText = g_Language.TranslateMessage("By default, PhotoDemon's undo data is not compressed.  This makes undo operations very fast, but increases disk space usage.  Compressing undo data will reduce disk space usage, at some cost to performance.  (Note that undo data is erased when PhotoDemon exits, so this setting only affects disk space usage while PhotoDemon is running.)")
            sltUndoCompression.Value = g_UndoCompressionLevel
        'END Undo data compression
        
    'END Performance preferences
    
    '***************************************************************************
    
    'START Color and Transparency preferences
    
        'START color management preferences
            
            'Set the option buttons according to the user's preference
            If g_UserPreferences.GetPref_Boolean("Transparency", "Use System Color Profile", True) Then optColorManagement(0).Value = True Else optColorManagement(1).Value = True
            
            'Load a list of all available monitors
            cboMonitors.Clear
            
            Dim PrimaryMonitor As String, secondaryMonitor As String
            PrimaryMonitor = g_Language.TranslateMessage("Primary monitor") & ": "
            secondaryMonitor = g_Language.TranslateMessage("Secondary monitor") & ": "
            
            Dim primaryIndex As Long
            
            Dim monitorEntry As String
            
            If g_Displays.GetDisplayCount > 0 Then
                
                For i = 0 To g_Displays.GetDisplayCount - 1
                
                    monitorEntry = ""
                    
                    'Explicitly label the primary monitor
                    If g_Displays.Displays(i).IsPrimary Then
                        monitorEntry = PrimaryMonitor
                        primaryIndex = i
                    Else
                        monitorEntry = secondaryMonitor
                    End If
                    
                    'Add the monitor's physical size
                    monitorEntry = monitorEntry & g_Displays.Displays(i).GetMonitorSizeAsString
                    
                    'Add the monitor's name
                    monitorEntry = monitorEntry & " " & g_Displays.Displays(i).GetBestMonitorName
                    
                    'Add the monitor's native resolution
                    monitorEntry = monitorEntry & " (" & g_Displays.Displays(i).GetMonitorResolutionAsString & ")"
                                    
                    'Display this monitor in the list
                    cboMonitors.AddItem monitorEntry, i
                    
                Next i
                
            Else
                primaryIndex = 0
                cboMonitors.AddItem "Unknown display", 0
            End If
            
            'Display the primary monitor by default; this will also trigger a load of the matching
            ' custom profile, if one exists.
            cboMonitors.ListIndex = primaryIndex
            
            'Add tooltips to all color-profile-related controls
            optColorManagement(0).AssignTooltip "This setting is the best choice for most users.  If you have no idea what color management is, use this setting.  If you have correctly configured a display profile via the Windows Control Panel, also use this setting."
            optColorManagement(1).AssignTooltip "To configure custom color profiles on a per-monitor basis, please use this setting."
            
            cboMonitors.AssignTooltip "Please specify a color profile for each monitor currently attached to the system.  Note that the text in parentheses is the display adapter driving the named monitor."
            cmdColorProfilePath.ToolTipText = g_Language.TranslateMessage("Click this button to bring up a ""browse for color profile"" dialog.")
        
        'END color management preferences
    
        'START alpha-channel checkerboard rendering
            userInitiatedAlphaSelection = False
            cboAlphaCheck.Clear
            cboAlphaCheck.AddItem " Highlight checks", 0
            cboAlphaCheck.AddItem " Midtone checks", 1
            cboAlphaCheck.AddItem " Shadow checks", 2
            cboAlphaCheck.AddItem " Custom (click boxes to customize)", 3
            
            cboAlphaCheck.ListIndex = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Mode", 0)
            
            csAlphaOne.Color = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check One", RGB(255, 255, 255))
            csAlphaTwo.Color = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Two", RGB(204, 204, 204))
            
            cboAlphaCheck.AssignTooltip "If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's colors."
            csAlphaOne.AssignTooltip "Click to change the first checkerboard background color for alpha channels"
            csAlphaTwo.AssignTooltip "Click to change the second checkerboard background color for alpha channels"
            
            userInitiatedAlphaSelection = True
        'END alpha-channel checkerboard rendering
        
        'START alpha-channel checkerboard size
            cboAlphaCheckSize.Clear
            cboAlphaCheckSize.AddItem " Small (4x4 pixels)", 0
            cboAlphaCheckSize.AddItem " Medium (8x8 pixels)", 1
            cboAlphaCheckSize.AddItem " Large (16x16 pixels)", 2
            
            cboAlphaCheckSize.ListIndex = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Size", 1)
            
            cboAlphaCheckSize.AssignTooltip "If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's size."
        'END alpha-channel checkerboard size
        
    'END Color and Transparency preferences
    
    '***************************************************************************
    
    'START Update preferences
    
        'START update frequency
            cboUpdates(0).Clear
            cboUpdates(0).AddItem "each session", 0
            cboUpdates(0).AddItem "weekly", 1
            cboUpdates(0).AddItem "monthly", 2
            cboUpdates(0).AddItem "never (not recommended)", 3
            
            'Old versions of PD used a binary check/don't check preference.  To respect users who set the "don't check" preference in a
            ' previous version, automatically convert that preference to the new "never (not recommended)" value.
            If g_UserPreferences.DoesValueExist("Updates", "Check For Updates") Then
                
                If Not g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True) Then
                    
                    'Write a matching preference in the new format.
                    g_UserPreferences.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
                    
                    'Overwrite the old preference, so it doesn't trigger again
                    g_UserPreferences.SetPref_Boolean "Updates", "Check For Updates", True
                    
                End If
                
            End If
            
            'Retrieve the current preference
            cboUpdates(0).ListIndex = g_UserPreferences.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
            cboUpdates(0).AssignTooltip "Because PhotoDemon is a portable application, it can only check for updates when the program is running.  By default, PhotoDemon will check for updates whenever the program is launched, but you can reduce this frequency if desired."
        'END update frequency
        
        'START update track
            cboUpdates(1).Clear
            cboUpdates(1).AddItem "stable releases", 0
            cboUpdates(1).AddItem "stable and beta releases", 1
            cboUpdates(1).AddItem "stable, beta, and developer releases", 2
            
            'Retrieve the current preference
            cboUpdates(1).ListIndex = g_UserPreferences.GetPref_Long("Updates", "Update Track", PDUT_BETA)
            cboUpdates(1).AssignTooltip "One of the best ways to support PhotoDemon is to help test new releases.  By default, PhotoDemon will suggest both stable and beta releases, but the truly adventurous can also try developer releases.  (Developer releases give you immediate access to the latest program enhancements, but you might encounter some bugs.)"
        'END update track
            
        'START update language files independently
            If g_UserPreferences.GetPref_Boolean("Updates", "Update Languages Independently", True) Then chkUpdates(0).Value = vbChecked Else chkUpdates(0).Value = vbUnchecked
            chkUpdates(0).AssignTooltip "PhotoDemon's volunteer translators regularly update the program's language files.  PhotoDemon can automatically download these updates separate from the main program, ensuring that you always have the most up-to-date language files."
        'END update language files independently
            
        'START update plugins independently
            If g_UserPreferences.GetPref_Boolean("Updates", "Update Plugins Independently", True) Then chkUpdates(1).Value = vbChecked Else chkUpdates(1).Value = vbUnchecked
            chkUpdates(1).AssignTooltip "PhotoDemon uses some 3rd-party plugins.  Sometimes, the authors of these plugins fix bugs or add new features.  Instead of waiting for the next PhotoDemon release, you can receive plugin updates as soon as they become available."
        'END update plugins independently
        
        'START notify when updates are ready for patching
            If g_UserPreferences.GetPref_Boolean("Updates", "Update Notifications", True) Then chkUpdates(2).Value = vbChecked Else chkUpdates(2).Value = vbUnchecked
            chkUpdates(2).AssignTooltip "PhotoDemon can notify you when it's ready to apply an update.  This allows you to use the updated version immediately."
        'END notify when updates are ready for patching
        
        'Populate the network access disclaimer in the "Update" panel
            lblExplanation.Caption = g_Language.TranslateMessage("The developers of PhotoDemon take privacy very seriously, so no information - statistical or otherwise - is uploaded during the update process.  Updates simply involve downloading several small XML files from photodemon.org. These files contain the latest software, plugin, and language version numbers. If updated versions are found, and user preferences allow, the updated files are then downloaded and patched automatically." & vbCrLf & vbCrLf & "If you still choose to disable updates, don't forget to visit photodemon.org from time to time to check for new versions.")
    
    'END Update preferences
    
    '***************************************************************************
    
    'START Advanced preferences
            
        'Display the current temporary file path
            txtTempPath.Text = g_UserPreferences.GetTempPath
    
        'Display what we know about this PC's hardware acceleration capabilities
            txtHardware = cSysInfo.GetDeviceCapsString()
            
        '...and give the "copy to clipboard" button a tooltip
            cmdCopyReportClipboard.AssignTooltip "Copy the report to the system clipboard"
        
        'Display what we know about PD's memory usage
            lblMemoryUsageCurrent.Caption = g_Language.TranslateMessage("current PhotoDemon memory usage:") & " " & Format(Str(cSysInfo.GetPhotoDemonMemoryUsage()), "###,###,###,###") & " K"
            lblMemoryUsageMax.Caption = g_Language.TranslateMessage("max PhotoDemon memory usage this session:") & " " & Format(Str(cSysInfo.GetPhotoDemonMemoryUsage(True)), "###,###,###,###") & " K"
            If Not g_IsProgramCompiled Then
                lblMemoryUsageCurrent.Caption = lblMemoryUsageCurrent.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
                lblMemoryUsageMax.Caption = lblMemoryUsageMax.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
            End If
    
    'END Advanced preferences
    
    '***************************************************************************
    
    'End batch preference mode
    g_UserPreferences.EndBatchPreferenceMode
    
    'All preference controls are now initialized with the matching value stored in the preferences file
    
End Sub

'When new transparency checkerboard colors are selected, change the corresponding list box to match
Private Sub csAlphaOne_ColorChanged()
    
    If userInitiatedAlphaSelection Then
        userInitiatedAlphaSelection = False
        cboAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
    End If
    
End Sub

Private Sub csAlphaTwo_ColorChanged()
    
    If userInitiatedAlphaSelection Then
        userInitiatedAlphaSelection = False
        cboAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
    End If
    
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the preferences file
Private Sub Form_Load()
        
    Set cSysInfo = New pdSystemInfo
        
    Dim i As Long
    
    'Populate all controls with the corresponding values from the preferences file
    LoadAllPreferences
    
    'Load custom command button images
    cmdCopyReportClipboard.AssignImage "CLIPBOARDCPY"
    
    'Prep the category button strip
    With btsvCategory
        
        'Start by adding captions for each button.  This will also update the control's layout to match.
        .AddItem "Interface", 0
        .AddItem "Loading", 1
        .AddItem "Saving", 2
        .AddItem "Performance", 3
        .AddItem "Color and Transparency", 4
        .AddItem "Updates", 5
        .AddItem "Advanced", 6
        
        'Next, add tooltips to each button
        .AssignTooltip "Interface options include settings for the main PhotoDemon interface, including things like canvas settings, font selection, and positioning.", "Interface Options", , 0
        .AssignTooltip "Load options allow you to customize the way image files enter the application.", "Load (Import) Options", , 1
        .AssignTooltip "Save options allow you to customize the way image files leave the application.", "Save (Export) Options", , 2
        .AssignTooltip "Performance options allow you to control whether PhotoDemon emphasizes speed or quality when performing certain tasks.", "Performance Options", , 3
        .AssignTooltip "Color and transparency options include settings for color management (ICC profiles), and alpha channel handling.", "Color and Transparency Options", , 4
        .AssignTooltip "Update options control how frequently PhotoDemon checks for updated versions, and how it handles the download of missing plugins.", "Update Options", , 5
        .AssignTooltip "Advanced options can be safely ignored by regular users. Testers and developers may, however, find these settings useful.", "Advanced Options", , 6
        
        'Next, add images to each button
        .AssignImageToItem 0, "PREF_INTERFACE"
        .AssignImageToItem 1, "PREF_LOADING"
        .AssignImageToItem 2, "PREF_SAVING"
        .AssignImageToItem 3, "PREF_PERFORMANCE"
        .AssignImageToItem 4, "PREF_COLOR"
        .AssignImageToItem 5, "PREF_NETWORK"
        .AssignImageToItem 6, "PREF_ADVANCED"
        
        'Finally, synchronize the tooltip manager against the current theme
        .UpdateAgainstCurrentTheme
        
    End With
    
    'Hide all category panels (the proper one will be activated in a moment)
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Activate the last preferences panel that the user looked at
    Dim activePanel As Long
    activePanel = g_UserPreferences.GetPref_Long("Core", "Last Preferences Page", 0)
    If activePanel > picContainer.UBound Then activePanel = picContainer.UBound
    picContainer(activePanel).Visible = True
    btsvCategory.ListIndex = activePanel
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    For i = 0 To picContainer.Count - 1
        SetArrowCursor picContainer(i)
    Next i
    
    userInitiatedColorSelection = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the selected temp folder doesn't have write access, warn the user
Private Sub TxtTempPath_Change()

    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not cFile.FolderExist(txtTempPath.Text) Then
        lblTempPathWarning.Caption = g_Language.TranslateMessage("WARNING: this folder is invalid (access prohibited).  Please provide a valid folder.  If no new folder is provided, PhotoDemon will use the system's default temp location.")
        lblTempPathWarning.Visible = True
    Else
        lblTempPathWarning.Visible = False
    End If
    
End Sub

