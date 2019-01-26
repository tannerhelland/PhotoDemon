VERSION 5.00
Begin VB.Form FormOptions 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Options"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
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
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   31
      Top             =   6870
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonStripVertical btsvCategory 
      Height          =   6675
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   11774
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   6
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdButtonStrip btsMouseHighRes 
         Height          =   975
         Left            =   0
         TabIndex        =   43
         Top             =   1950
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1720
         Caption         =   "high-resolution mouse input"
      End
      Begin PhotoDemon.pdButton cmdReset 
         Height          =   600
         Left            =   240
         TabIndex        =   30
         Top             =   4665
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1058
         Caption         =   "reset all program settings"
      End
      Begin PhotoDemon.pdButton cmdTmpPath 
         Height          =   450
         Left            =   7680
         TabIndex        =   29
         Top             =   5775
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   794
         Caption         =   "..."
      End
      Begin PhotoDemon.pdTextBox txtTempPath 
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   5850
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         Text            =   "automatically generated at run-time"
      End
      Begin PhotoDemon.pdLabel lblMemoryUsageMax 
         Height          =   345
         Left            =   240
         Top             =   3855
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   609
         Caption         =   "memory usage will be displayed here"
         ForeColor       =   8405056
      End
      Begin PhotoDemon.pdLabel lblMemoryUsageCurrent 
         Height          =   345
         Left            =   240
         Top             =   3495
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   609
         Caption         =   "memory usage will be displayed here"
         ForeColor       =   8405056
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   5
         Left            =   0
         Top             =   3135
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "memory diagnostics"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   19
         Left            =   0
         Top             =   5400
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "temporary file location"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTempPathWarning 
         Height          =   480
         Left            =   240
         Top             =   6240
         Visible         =   0   'False
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   847
         ForeColor       =   255
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   4305
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "start over"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "application settings folder"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSettingsFolder 
         Height          =   285
         Left            =   240
         Top             =   360
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   503
         Caption         =   ""
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdButtonStrip btsDebug 
         Height          =   975
         Left            =   0
         TabIndex        =   44
         Top             =   780
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1720
         Caption         =   "generate debug logs"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   5
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdLabel lblExplanation 
         Height          =   3495
         Left            =   240
         Top             =   3000
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6165
         Caption         =   "(disclaimer populated at run-time)"
         FontSize        =   9
         Layout          =   1
      End
      Begin PhotoDemon.pdDropDown cboUpdates 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "automatically check for updates:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboUpdates 
         Height          =   735
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   1350
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "allow updates from these tracks:"
         FontSizeCaption =   10
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
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   2400
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
         Caption         =   "notify me when an update is ready"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdDropDown cboMRUStyle 
         Height          =   810
         Left            =   180
         TabIndex        =   1
         Top             =   1800
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   1429
         Caption         =   "recently used file shortcuts:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboTitleText 
         Height          =   810
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   1429
         Caption         =   "main window title bar text:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSpinner tudRecentFiles 
         Height          =   345
         Left            =   3840
         TabIndex        =   5
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         DefaultValue    =   10
         Min             =   1
         Max             =   32
         Value           =   10
      End
      Begin PhotoDemon.pdLabel lblRecentFileCount 
         Height          =   240
         Left            =   180
         Top             =   2670
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   423
         Caption         =   "maximum number of recent file entries: "
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   13
         Left            =   0
         Top             =   1440
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "recent files list"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "main window"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdDropDown cboAlphaCheckSize 
         Height          =   810
         Left            =   180
         TabIndex        =   39
         Top             =   4530
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1429
         Caption         =   "transparency checkerboard size:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboAlphaCheck 
         Height          =   795
         Left            =   180
         TabIndex        =   40
         Top             =   3660
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1402
         Caption         =   "transparency checkerboard colors:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdColorSelector csAlphaOne 
         Height          =   435
         Left            =   6240
         TabIndex        =   41
         Top             =   3990
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
         ShowMainWindowColor=   0   'False
      End
      Begin PhotoDemon.pdColorSelector csAlphaTwo 
         Height          =   435
         Left            =   7320
         TabIndex        =   42
         Top             =   3990
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
         ShowMainWindowColor=   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   3240
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   503
         Caption         =   "transparency"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   4
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdCheckBox chkColorManagement 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   4080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         Caption         =   "use black point compensation"
      End
      Begin PhotoDemon.pdDropDown cboDisplayRenderIntent 
         Height          =   735
         Left            =   180
         TabIndex        =   38
         Top             =   3240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1296
         Caption         =   "display rendering intent:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButton cmdColorProfilePath 
         Height          =   375
         Left            =   7380
         TabIndex        =   28
         Top             =   2760
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   661
         Caption         =   "..."
      End
      Begin PhotoDemon.pdDropDown cboMonitors 
         Height          =   690
         Left            =   780
         TabIndex        =   9
         Top             =   1590
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   1217
         Caption         =   "available displays:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdTextBox txtColorProfilePath 
         Height          =   315
         Left            =   900
         TabIndex        =   10
         Top             =   2790
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   556
         Text            =   "(none)"
      End
      Begin PhotoDemon.pdRadioButton optColorManagement 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "turn off display color management"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdRadioButton optColorManagement 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   840
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "use the current system profiles for each display"
      End
      Begin PhotoDemon.pdLabel lblColorManagement 
         Height          =   240
         Index           =   2
         Left            =   780
         Top             =   2430
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   503
         Caption         =   "color profile for this display:"
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
         Caption         =   "display policies"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdRadioButton optColorManagement 
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   1200
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "use custom profiles for each display"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   2
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdCheckBox chkConfirmUnsaved 
         Height          =   330
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "when closing images, warn me me about unsaved changes"
      End
      Begin PhotoDemon.pdDropDown cboDefaultSaveFormat 
         Height          =   690
         Left            =   180
         TabIndex        =   15
         Top             =   1455
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   582
         Caption         =   "when using the ""Save As"" command, set the default file format according to:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboSaveBehavior 
         Height          =   690
         Left            =   180
         TabIndex        =   16
         Top             =   4125
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   582
         Caption         =   "when ""Save"" is used:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   2490
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   503
         Caption         =   "metadata"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   6
         Left            =   0
         Top             =   3645
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         Caption         =   "save behavior: overwrite vs make a copy"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   503
         Caption         =   "closing unsaved images"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   8
         Left            =   0
         Top             =   990
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   503
         Caption         =   "default file format when saving"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdCheckBox chkMetadataListPD 
         Height          =   375
         Left            =   180
         TabIndex        =   36
         Top             =   3000
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "list PhotoDemon as the last-used editing software"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdCheckBox chkToneMapping 
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "display tone mapping options when importing HDR and RAW images"
      End
      Begin PhotoDemon.pdCheckBox chkLoadingOrientation 
         Height          =   330
         Left            =   180
         TabIndex        =   18
         Top             =   3360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "obey auto-rotate instructions inside image files"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   9
         Left            =   0
         Top             =   3000
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         Caption         =   "orientation"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   503
         Caption         =   "high-dynamic range (HDR) images"
         FontSize        =   12
         ForeColor       =   5263440
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   12
         Left            =   0
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
         Left            =   180
         TabIndex        =   32
         Top             =   2400
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "forcibly extract binary-type tags as Base64 (slow)"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkMetadataJPEG 
         Height          =   330
         Left            =   180
         TabIndex        =   33
         Top             =   1680
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "estimate original JPEG quality settings"
      End
      Begin PhotoDemon.pdCheckBox chkMetadataUnknown 
         Height          =   330
         Left            =   180
         TabIndex        =   34
         Top             =   2040
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "extract unknown tags"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkMetadataDuplicates 
         Height          =   330
         Left            =   180
         TabIndex        =   35
         Top             =   1320
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   582
         Caption         =   "automatically hide duplicate tags"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6720
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11853
      Begin PhotoDemon.pdSlider sltUndoCompression 
         Height          =   765
         Left            =   180
         TabIndex        =   19
         Top             =   4170
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   873
         Caption         =   "compress undo/redo data at the following level:"
         FontSizeCaption =   10
         Max             =   9
         SliderTrackStyle=   1
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   690
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   1217
         Caption         =   "when decorating interface elements:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   690
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   1620
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   1217
         Caption         =   "when generating image and layer thumbnail images:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboPerformance 
         Height          =   690
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   2850
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   1217
         Caption         =   "when rendering the image canvas:"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   16
         Left            =   0
         Top             =   0
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
         Left            =   300
         Top             =   5040
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
         Left            =   3960
         Top             =   5040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "maximum compression (slowest)"
         FontSize        =   8
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   18
         Left            =   0
         Top             =   3780
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         Caption         =   "undo/redo"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   15
         Left            =   0
         Top             =   1260
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   503
         Caption         =   "thumbnails"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   17
         Left            =   0
         Top             =   2490
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   503
         Caption         =   "viewport"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "FormOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright 2002-2019 by Tanner Helland
'Created: 8/November/02
'Last updated: 28/November/16
'Last update: overhaul a number of panels to reflect new program behavior in 7.0
'
'Dialog for interfacing with the user's desired program preferences.  Handles reading/writing from/to the persistent
' XML file that actually stores all preferences.
'
'Note that this form interacts heavily with the pdPreferences class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Private m_userInitiatedColorSelection As Boolean, m_userInitiatedAlphaSelection As Boolean

Private Sub btsvCategory_Click(ByVal buttonIndex As Long)

    'When the preferences category is changed, only display the controls in that category
    Dim catID As Long
    For catID = 0 To btsvCategory.ListCount - 1
        
        If (catID = buttonIndex) Then
            picContainer(catID).Visible = True
        Else
            picContainer(catID).Visible = False
        End If
        
    Next catID

End Sub

'Alpha channel checkerboard selection; change the color selectors to match
Private Sub cboAlphaCheck_Click()

    'Only respond to user-generated events
    If m_userInitiatedAlphaSelection Then

        m_userInitiatedAlphaSelection = False

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
        
        m_userInitiatedAlphaSelection = True
                
    End If

End Sub

'Whenever the Color and Transparency -> Color Management -> Monitor combo box is changed, load the relevant color profile
' path from the preferences file (if one exists)
Private Sub cboMonitors_Click()

    'One of the difficulties with tracking multiple monitors is that the user can attach/detach them at will.
    
    'Prior to v7.0, PD used HMONITOR handles to track displays, using the reasoning from this article:
    ' http://www.microsoft.com/msj/0697/monitor/monitor.aspx
    '...specifically the line, "A physical device has the same HMONITOR value throughout its lifetime, even across changes
    ' to display settings, as long as it remains a part of the desktop."
    
    'This worked "well enough", as long as the user never disconnected the display monitor only to attach it again at
    ' some point in the future (as is common with second monitors and a laptop, for example).
    
    'In 7.0, this system was upgraded to use monitor serial numbers, and only fall back to the HMONITOR value if a
    ' serial number (or EDID) doesn't exist.
    
    Dim uniqueMonitorID As String
    If (Not g_Displays.Displays(cboMonitors.ListIndex) Is Nothing) Then
        uniqueMonitorID = g_Displays.Displays(cboMonitors.ListIndex).GetUniqueDescriptor
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        uniqueMonitorID = tmpXML.GetXMLSafeTagName(uniqueMonitorID)
    End If
    
    'Use that to retrieve a stored color profile (if any)
    Dim profilePath As String
    profilePath = UserPrefs.GetPref_String("ColorManagement", "DisplayProfile_" & uniqueMonitorID, "(none)")
    
    'If the returned value is "(none)", translate that into the user's language before displaying; otherwise, display
    ' whatever path we retrieved.
    If Strings.StringsEqual(profilePath, "(none)", False) Then
        txtColorProfilePath.Text = g_Language.TranslateMessage("(none)")
    Else
        txtColorProfilePath.Text = profilePath
    End If
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Start by auto-validating any controls that accept user input
    Dim validateCheck As Boolean
    validateCheck = True
    
    Dim eControl As Object
    For Each eControl In FormOptions.Controls
        
        'Obviously, we can only validate our own custom objects that have built-in auto-validate functions.
        If (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSpinner) Then
            
            'Finally, ask the control to validate itself
            If (Not eControl.IsValid) Then
                validateCheck = False
                Exit For
            End If
            
        End If
    Next eControl
    
    If (Not validateCheck) Then
        cmdBarMini.DoNotUnloadForm
        Exit Sub
    End If
    
    Message "Saving preferences..."
    Me.Visible = False
    
    'After updates on 22 Oct 2014, the preference saving sequence should happen in a flash, but just in case,
    ' we'll supply a bit of processing feedback.
    FormMain.Enabled = False
    SetProgBarMax 8
    SetProgBarVal 1
    
    'First, make note of the active panel, so we can default to that if the user returns to this dialog
    UserPrefs.SetPref_Long "Core", "Last Preferences Page", btsvCategory.ListIndex
    
    'Write preferences out to file in category order.  (The preference XML file is order-agnostic, but I try to
    ' maintain the order used in the Preferences dialog itself to make changes easier.)
    
    '***************************************************************************
    
    'BEGIN Interface preferences
        
        'START/END image window caption length
            UserPrefs.SetPref_Long "Interface", "Window Caption Length", cboTitleText.ListIndex
        
        Dim mruNeedsToBeRebuilt As Boolean
        mruNeedsToBeRebuilt = False
        
        'START MRU caption length
        
            'Check to see if the new MRU caption setting matches the old one.  If it doesn't, reload the MRU.
            If (cboMRUStyle.ListIndex <> UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0)) Then mruNeedsToBeRebuilt = True
            UserPrefs.SetPref_Long "Interface", "MRU Caption Length", cboMRUStyle.ListIndex
            
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
            If (newMaxRecentFiles <> UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)) Then mruNeedsToBeRebuilt = True
            UserPrefs.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value
            
        'END maximum MRU count
        
        'If any MRUs need to be rebuilt, do so now
        If mruNeedsToBeRebuilt Then
            g_RecentFiles.NotifyMaxLimitChanged
            g_RecentMacros.MRU_NotifyNewMaxLimit
        End If
        
        'START alpha checkerboard colors
            UserPrefs.SetPref_Long "Transparency", "Alpha Check Mode", CLng(cboAlphaCheck.ListIndex)
            UserPrefs.SetPref_Long "Transparency", "Alpha Check One", CLng(csAlphaOne.Color)
            UserPrefs.SetPref_Long "Transparency", "Alpha Check Two", CLng(csAlphaTwo.Color)
        'END alpha checkerboard colors
            
        'START alpha checkerboard size
            UserPrefs.SetPref_Long "Transparency", "Alpha Check Size", cboAlphaCheckSize.ListIndex
            Drawing.CreateAlphaCheckerboardDIB g_CheckerboardPattern
        'END alpha checkerboard size
    
    'END Interface preferences
    
    '***************************************************************************
    
    SetProgBarVal 2
    
    'BEGIN Loading preferences
    
        'START/END automatically tone-map HDR images
            UserPrefs.SetPref_Boolean "Loading", "Tone Mapping Prompt", chkToneMapping.Value
            
        'START metadata behavior at load-time
            UserPrefs.SetPref_Boolean "Loading", "Metadata Hide Duplicates", chkMetadataDuplicates.Value
            UserPrefs.SetPref_Boolean "Loading", "Metadata Estimate JPEG", chkMetadataJPEG.Value
            UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Binary", chkMetadataBinary.Value
            UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Unknown", chkMetadataUnknown.Value
        'END metadata behavior at load-time
        
        'START/END EXIF auto-rotation
            UserPrefs.SetPref_Boolean "Loading", "ExifAutoRotate", chkLoadingOrientation.Value
        
    
    'END Loading preferences
    
    '***************************************************************************
    
    SetProgBarVal 3
    
    'BEGIN Saving preferences
    
        'START prompt on unsaved images
            g_ConfirmClosingUnsaved = chkConfirmUnsaved.Value
            UserPrefs.SetPref_Boolean "Saving", "Confirm Closing Unsaved", g_ConfirmClosingUnsaved
    
            If g_ConfirmClosingUnsaved Then
                toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "If the current image has not been saved, you will receive a prompt to save it before it closes.", "Close the current image"
            Else
                toolbar_Toolbox.cmdFile(FILE_CLOSE).AssignTooltip "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.", "Close the current image"
            End If
        'END prompt on unsaved images
        
        'START/END metadata-related options
            UserPrefs.SetPref_Boolean "Saving", "MetadataListPD", chkMetadataListPD.Value
        
        'START/END Save behavior (overwrite or copy)
            UserPrefs.SetPref_Long "Saving", "Overwrite Or Copy", cboSaveBehavior.ListIndex
        
        'START/END "Save As" dialog's suggested file format
            UserPrefs.SetPref_Long "Saving", "Suggested Format", cboDefaultSaveFormat.ListIndex
    
    'END Saving preferences
        
    '***************************************************************************
    
    SetProgBarVal 4
    
    'START Performance preferences
        
        'START/END interface decoration performance
            UserPrefs.SetPref_Long "Performance", "Interface Decoration Performance", cboPerformance(0).ListIndex
            g_InterfacePerformance = cboPerformance(0).ListIndex
        
        'START/END thumbnail render performance
            UserPrefs.SetPref_Long "Performance", "Thumbnail Performance", cboPerformance(1).ListIndex
            UserPrefs.SetThumbnailPerformancePref cboPerformance(1).ListIndex
        
        'START/END viewport render performance
            UserPrefs.SetPref_Long "Performance", "Viewport Render Performance", cboPerformance(2).ListIndex
            g_ViewportPerformance = cboPerformance(2).ListIndex
            
        'START/END undo/redo data compression
            UserPrefs.SetPref_Long "Performance", "Undo Compression", sltUndoCompression.Value
            g_UndoCompressionLevel = sltUndoCompression.Value
    
    'END Performance preferences
    
    '***************************************************************************
    
    SetProgBarVal 5
    
    'START Color Management preferences

        'START use system color profile
            If optColorManagement(0).Value Then
                ColorManagement.SetDisplayColorManagementPreference DCM_NoManagement
            ElseIf optColorManagement(1).Value Then
                ColorManagement.SetDisplayColorManagementPreference DCM_SystemProfile
            Else
                ColorManagement.SetDisplayColorManagementPreference DCM_CustomProfile
            End If
            
            ColorManagement.SetDisplayBPC chkColorManagement(0).Value
            ColorManagement.SetDisplayRenderingIntentPref cboDisplayRenderIntent.ListIndex
            
            'Changes to color preferences require us to re-cache any working-space-to-screen transform data.
            CacheDisplayCMMData
            ColorManagement.CheckParentMonitor False, True
        'END use system color profile
        
    'END Color Management preferences
    
    '***************************************************************************
    
    SetProgBarVal 6
    
    'BEGIN Update preferences
        
        'START/END update frequency
            UserPrefs.SetPref_Long "Updates", "Update Frequency", cboUpdates(0).ListIndex
        
        'START/END update track
            UserPrefs.SetPref_Long "Updates", "Update Track", cboUpdates(1).ListIndex
            
        'START/END update notifications
            UserPrefs.SetPref_Boolean "Updates", "Update Notifications", chkUpdates(0).Value
    
    'END Update preferences
    
    '***************************************************************************
    
    SetProgBarVal 7
    
    'BEGIN Advanced preferences
    
        'START generate debug logs
            
            'First, see if the user has changed the debug log preference
            If (UserPrefs.GetDebugLogPreference <> btsDebug.ListIndex) Then
                
                'The user has changed the current setting.  Make a note of whether debug logs are currently being generated.
                ' (If this behavior changes, we may need to create and/or terminate the debugger.)
                Dim curLogBehavior As Boolean
                curLogBehavior = UserPrefs.GenerateDebugLogs()
                
                'Store the new preference
                UserPrefs.SetDebugLogPreference btsDebug.ListIndex
                
                'Invoke and/or terminate the current debugger, as necessary
                If (curLogBehavior <> UserPrefs.GenerateDebugLogs()) Then
                    If UserPrefs.GenerateDebugLogs Then PDDebug.StartDebugger True, , False Else PDDebug.TerminateDebugger False
                End If
                
            End If
            
        'END generate debug logs
            
        'START/END store the temporary path (but only if it's changed)
            If Strings.StringsNotEqual(Trim$(txtTempPath), UserPrefs.GetTempPath, True) Then UserPrefs.SetTempPath Trim$(txtTempPath)
        
        'START/END high-resolution mouse input
            If (btsMouseHighRes.ListIndex = 1) Then UserPrefs.SetPref_Boolean "Tools", "HighResMouseInput", True Else UserPrefs.SetPref_Boolean "Tools", "HighResMouseInput", False
            Tools.SetToolSetting_HighResMouse (btsMouseHighRes.ListIndex = 1)
        
    'END Advanced preferences
    
    '***************************************************************************
    
    'Forcibly write a copy of the preference data out to file
    UserPrefs.ForceWriteToFile
    
    'All user preferences have now been written out to file
    
    'Because some preferences affect the program's interface, redraw the active image.
    FormMain.Enabled = True
    FormMain.UpdateMainLayout
        
    'TODO: color management changes need to be propagated here; otherwise, they won't trigger until the program is restarted.
    
    SetProgBarVal 0
    ReleaseProgressBar
    
    Message "Preferences updated."
        
End Sub

'Allow the user to select a new color profile for the attached monitor.  Because this text box is re-used for multiple
' settings, save any changes to file immediately, rather than waiting for the user to click OK.
Private Sub cmdColorProfilePath_Click()

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    Dim sFile As String
    
    'Get the last color profile path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Color Profile", vbNullString)
    
    'If no color profile path was found, populate it with the default system color profile path
    If (Len(tempPathString) = 0) Then tempPathString = GetSystemColorFolder()
    
    'Prepare a common dialog filter list with extensions of known profile types
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("ICC Profiles") & " (.icc, .icm)|*.icc;*.icm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Please select a color profile")
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, cdTitle, ".icc", FormOptions.hWnd) Then
        
        sFile = Strings.TrimNull(sFile)
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = Files.FileGetPath(sFile)
        UserPrefs.SetPref_String "Paths", "Color Profile", listPath
        
        'Set the text box to match this color profile, and save the resulting preference out to file.
        txtColorProfilePath = sFile
        
        Dim uniqueMonID As String
        If (Not g_Displays.Displays(cboMonitors.ListIndex) Is Nothing) Then
            uniqueMonID = g_Displays.Displays(cboMonitors.ListIndex).GetUniqueDescriptor
            Dim tmpXML As pdXML
            Set tmpXML = New pdXML
            uniqueMonID = tmpXML.GetXMLSafeTagName(uniqueMonID)
        End If
        
        UserPrefs.SetPref_String "ColorManagement", "DisplayProfile_" & uniqueMonID, sFile
        
        'If the "user custom color profiles" option button isn't selected, mark it now
        If (Not optColorManagement(2).Value) Then optColorManagement(2).Value = True
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput

End Sub

'RESET will regenerate the preferences file from scratch.  This can be an effective way to
' "reset" a copy of the program.
Private Sub cmdReset_Click()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = PDMsgBox("All settings will be restored to their default values.  This action cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbExclamation Or vbYesNo, "Reset PhotoDemon")

    'If the user gives final permission, rewrite the preferences file from scratch and repopulate this form
    If (confirmReset = vbYes) Then
    
        UserPrefs.ResetPreferences
        LoadAllPreferences
        
        'Restore the currently active language to the preferences file; this prevents the language from resetting to English
        ' (a behavior that isn't made clear by this action).
        g_Language.WriteLanguagePreferencesToFile
        
    End If

End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = Files.PathBrowseDialog(Me.hWnd, UserPrefs.GetTempPath)
    If (Len(tString) <> 0) Then txtTempPath.Text = Files.PathAddBackslash(tString)
End Sub

'Load all relevant values from the preferences file, and populate their corresponding controls with the user's current settings
Private Sub LoadAllPreferences()
    
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
            cboTitleText.Clear
            cboTitleText.AddItem "compact - image names only", 0
            cboTitleText.AddItem "descriptive - full image locations, including folder(s)", 1
            cboTitleText.ListIndex = UserPrefs.GetPref_Long("Interface", "Window Caption Length", 0)
            cboTitleText.AssignTooltip "The title bar of the main PhotoDemon window displays information about the currently loaded image.  Use this preference to control how much information is displayed."
        'END image window caption length
                
        'START Recent file max count
            lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent file entries: ")
            tudRecentFiles.SetLeft lblRecentFileCount.GetLeft + lblRecentFileCount.GetWidth + FixDPI(6)
            tudRecentFiles.Value = UserPrefs.GetPref_Long("Interface", "Recent Files Limit", 10)
        'END
        
        'START MRU caption length
            cboMRUStyle.Clear
            cboMRUStyle.AddItem "compact - image names only", 0
            cboMRUStyle.AddItem "descriptive - full image locations, including folder(s)", 1
            cboMRUStyle.ListIndex = UserPrefs.GetPref_Long("Interface", "MRU Caption Length", 0)
            cboMRUStyle.AssignTooltip "The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations."
        'END MRU caption length
        
        'START alpha-channel checkerboard rendering
            m_userInitiatedAlphaSelection = False
            cboAlphaCheck.Clear
            cboAlphaCheck.AddItem "Highlight checks", 0
            cboAlphaCheck.AddItem "Midtone checks", 1
            cboAlphaCheck.AddItem "Shadow checks", 2
            cboAlphaCheck.AddItem "Custom (click boxes to customize)", 3
            
            cboAlphaCheck.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Mode", 0)
            
            csAlphaOne.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check One", RGB(255, 255, 255))
            csAlphaTwo.Color = UserPrefs.GetPref_Long("Transparency", "Alpha Check Two", RGB(204, 204, 204))
            
            cboAlphaCheck.AssignTooltip "If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's colors."
            csAlphaOne.AssignTooltip "Click to change the first checkerboard background color for alpha channels"
            csAlphaTwo.AssignTooltip "Click to change the second checkerboard background color for alpha channels"
            
            m_userInitiatedAlphaSelection = True
        'END alpha-channel checkerboard rendering
        
        'START alpha-channel checkerboard size
            cboAlphaCheckSize.Clear
            cboAlphaCheckSize.AddItem "Small (4x4 pixels)", 0
            cboAlphaCheckSize.AddItem "Medium (8x8 pixels)", 1
            cboAlphaCheckSize.AddItem "Large (16x16 pixels)", 2
            
            cboAlphaCheckSize.ListIndex = UserPrefs.GetPref_Long("Transparency", "Alpha Check Size", 1)
            
            cboAlphaCheckSize.AssignTooltip "If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's size."
        'END alpha-channel checkerboard size
        
    'END Interface preferences
    
    '***************************************************************************
    
    'START Loading preferences
    
        'START tone-mapping HDR images at load time
            chkToneMapping.Value = UserPrefs.GetPref_Boolean("Loading", "Tone Mapping Prompt", True)
            chkToneMapping.Enabled = ImageFormats.IsFreeImageEnabled
            If (Not ImageFormats.IsFreeImageEnabled) Then chkToneMapping.Caption = g_Language.TranslateMessage("feature disabled due to missing plugin")
            chkToneMapping.AssignTooltip "HDR and RAW images contain more colors than PC screens can physically display.  Before displaying such images, a tone mapping operation must be applied to the original image data."
        'END tone-mapping HDR images at load time
        
        'START metadata behavior at load-time
            chkMetadataDuplicates.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Hide Duplicates", True)
            chkMetadataDuplicates.AssignTooltip "Older cameras and photo-editing software may not embed metadata correctly, leading to multiple metadata copies within a single file.  PhotoDemon can automatically resolve duplicate entries for you."
            chkMetadataJPEG.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Estimate JPEG", True)
            chkMetadataJPEG.AssignTooltip "The JPEG format does not provide a way to store JPEG quality settings inside image files.  PhotoDemon can work around this by inferring quality settings from other metadata (like quantization tables)."
            chkMetadataUnknown.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Unknown", False)
            chkMetadataUnknown.AssignTooltip "Some camera manufacturers store proprietary metadata tags inside image files.  These tags are not generally useful to humans, but PhotoDemon can attempt to extract them anyway."
            chkMetadataBinary.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Binary", False)
            chkMetadataBinary.AssignTooltip "By default, large binary tags (like image thumbnails) are not processed.  Instead, PhotoDemon simply reports the size of the embedded data.  If you require this data, PhotoDemon can manually convert it to Base64 for further analysis."
        'END metadata behavior at load-time
        
        'START auto-rotate according to EXIF data
            chkLoadingOrientation.Value = UserPrefs.GetPref_Boolean("Loading", "EXIF Auto Rotate", True)
            chkLoadingOrientation.AssignTooltip "Most digital photos include rotation instructions (EXIF orientation metadata), which PhotoDemon will use to automatically rotate photos.  Some older smartphones and cameras may not write these instructions correctly, so if your photos are being imported sideways or upside-down, you can try disabling the auto-rotate feature."
        'END auto-rotate according to EXIF data
        
    
    'END Loading preferences
    
    '***************************************************************************
    
    'START Saving preferences
    
        'START/END prompt about unsaved images
            chkConfirmUnsaved.Value = g_ConfirmClosingUnsaved
            chkConfirmUnsaved.AssignTooltip "By default, PhotoDemon will warn you when you attempt to close an image with unsaved changes."
            
        'START suggested save as format
            cboDefaultSaveFormat.Clear
            cboDefaultSaveFormat.AddItem "the current file format of the image being saved", 0
            cboDefaultSaveFormat.AddItem "the last image format I used in the ""Save As"" screen", 1
            cboDefaultSaveFormat.ListIndex = UserPrefs.GetPref_Long("Saving", "Suggested Format", 0)
            
            cboDefaultSaveFormat.AssignTooltip "Most photo editors use the format of the current image as the default in the ""Save As"" screen.  When working with RAW images that will eventually be saved to JPEG, it is useful to have PhotoDemon remember that - hence the ""last used"" option."
        'END suggested save as format
        
        'START overwrite vs copy when saving
            cboSaveBehavior.Clear
            cboSaveBehavior.AddItem "overwrite the current file (standard behavior)", 0
            cboSaveBehavior.AddItem "save a new copy, e.g. ""filename (2).jpg"" (safe behavior)", 1
            cboSaveBehavior.ListIndex = UserPrefs.GetPref_Long("Saving", "Overwrite Or Copy", 0)
            
            cboSaveBehavior.AssignTooltip "In most photo editors, the ""Save"" command saves the image over its original version, erasing that copy forever.  PhotoDemon provides a ""safer"" option, where each save results in a new copy of the file."
        'END overwrite vs copy when saving
               
        'START/END embed PD as the last-used program
            chkMetadataListPD.Value = UserPrefs.GetPref_Boolean("Saving", "MetadataListPD", True)
            chkMetadataListPD.AssignTooltip "The EXIF specification asks programs to correctly identify themselves as the software of origin when exporting image files.  For increased privacy, you can suspend this behavior."
        
    'END Saving preferences
    
    '***************************************************************************
    
    'START Performance preferences
        
        'We can shortcut a bit of initialization here by populating all quality drop-downs with the same values.
        Dim i As Long
        
        For i = 0 To cboPerformance.UBound
            cboPerformance(i).Clear
            cboPerformance(i).AddItem "maximize quality", 0
            cboPerformance(i).AddItem "balance performance and quality", 1
            cboPerformance(i).AddItem "maximize performance", 2
        Next i
        
        'START Interface decorations performance
            cboPerformance(0).ListIndex = g_InterfacePerformance
            cboPerformance(0).AssignTooltip "Some interface elements receive custom decorations (like drop shadows).  On older PCs, these decorations can be suspended for a small performance boost."
        'END Interface decorations performance
        
        'START Thumbnail rendering performance
            cboPerformance(1).ListIndex = UserPrefs.GetThumbnailPerformancePref()
            cboPerformance(1).AssignTooltip "PhotoDemon generates many thumbnail images, especially when images contain multiple layers.  Thumbnail quality can be lowered to improve performance."
        'END Thumbnail rendering performance
        
        'START Viewport rendering performance
            cboPerformance(2).ListIndex = g_ViewportPerformance
            cboPerformance(2).AssignTooltip "Rendering the primary image canvas is a common bottleneck for PhotoDemon's performance.  The automatic setting is recommended, but for older PCs, you can manually select the Maximize Performance option to sacrifice quality for raw performance."
        'END Viewport rendering performance
        
        'START Undo data compression
            sltUndoCompression.AssignTooltip "Low compression settings require more disk space, but undo/redo operations will be faster.  High compression settings require less disk space, but undo/redo operations will be slower.  Undo data is erased when images are closed, so this setting only affects disk space while images are actively being edited."
            sltUndoCompression.Value = g_UndoCompressionLevel
        'END Undo data compression
        
    'END Performance preferences
    
    '***************************************************************************
    
    'START Color Management preferences
            
        'Set the various buttons and dropdown according to the user's current display profile preference
        optColorManagement(ColorManagement.GetDisplayColorManagementPreference()).Value = True
        chkColorManagement(0).Value = ColorManagement.GetDisplayBPC()
        Interface.PopulateRenderingIntentDropDown cboDisplayRenderIntent, ColorManagement.GetDisplayRenderingIntentPref()
        
        'Load a list of all available monitors
        cboMonitors.Clear
        
        Dim mainMonitor As String, secondaryMonitor As String
        mainMonitor = g_Language.TranslateMessage("Primary monitor") & ": "
        secondaryMonitor = g_Language.TranslateMessage("Secondary monitor") & ": "
        
        Dim primaryIndex As Long, monitorEntry As String
        
        If (g_Displays.GetDisplayCount > 0) Then
            
            For i = 0 To g_Displays.GetDisplayCount - 1
            
                monitorEntry = vbNullString
                
                'Explicitly label the primary monitor
                If g_Displays.Displays(i).IsPrimary Then
                    monitorEntry = mainMonitor
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
        optColorManagement(0).AssignTooltip "Turning off display color management can provide a small performance boost.  If your display is not currently configured for color management, use this setting."
        optColorManagement(1).AssignTooltip "This setting is the best choice for most users.  If you have no idea what color management is, use this setting.  If you have correctly configured a display profile via the Windows Control Panel, also use this setting."
        optColorManagement(2).AssignTooltip "To configure custom color profiles on a per-monitor basis, please use this setting."
        
        cboMonitors.AssignTooltip "Please specify a color profile for each monitor currently attached to the system.  Note that the text in parentheses is the display adapter driving the named monitor."
        cmdColorProfilePath.AssignTooltip "Click this button to bring up a ""browse for color profile"" dialog."
        
        cboDisplayRenderIntent.AssignTooltip "If you do not know what this setting controls, set it to ""Perceptual"".  Perceptual rendering intent is the best choice for most users."
        chkColorManagement(0).AssignTooltip "BPC is primarily relevant in colorimetric rendering intents, where it helps preserve detail in dark (shadow) regions of images.  For most workflows, BPC should be turned ON."
        
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
            If UserPrefs.DoesValueExist("Updates", "CheckForUpdates") Then
                
                If (Not UserPrefs.GetPref_Boolean("Updates", "CheckForUpdates", True)) Then
                    
                    'Write a matching preference in the new format.
                    UserPrefs.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
                    
                    'Overwrite the old preference, so it doesn't trigger again
                    UserPrefs.SetPref_Boolean "Updates", "CheckForUpdates", True
                    
                End If
                
            End If
            
            'Retrieve the current preference
            cboUpdates(0).ListIndex = UserPrefs.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
            cboUpdates(0).AssignTooltip "Because PhotoDemon is a portable application, it can only check for updates when the program is running.  By default, PhotoDemon will check for updates whenever the program is launched, but you can reduce this frequency if desired."
        'END update frequency
        
        'START update track
            cboUpdates(1).Clear
            cboUpdates(1).AddItem "stable releases", 0
            cboUpdates(1).AddItem "stable and beta releases", 1
            cboUpdates(1).AddItem "stable, beta, and developer releases", 2
            
            'Retrieve the current preference
            cboUpdates(1).ListIndex = UserPrefs.GetPref_Long("Updates", "Update Track", ut_Beta)
            cboUpdates(1).AssignTooltip "One of the best ways to support PhotoDemon is to help test new releases.  By default, PhotoDemon will suggest both stable and beta releases, but the truly adventurous can also try developer releases.  (Developer releases give you immediate access to the latest program enhancements, but you might encounter some bugs.)"
        'END update track
        
        'START notify when updates are ready for patching
            chkUpdates(0).Value = UserPrefs.GetPref_Boolean("Updates", "Update Notifications", True)
            chkUpdates(0).AssignTooltip "PhotoDemon can notify you when it's ready to apply an update.  This allows you to use the updated version immediately."
        'END notify when updates are ready for patching
        
        'START explanation of update options
        
            'In normal (portable) mode, I like to provide a short explanation of how automatic updates work.
            ' In non-portable mode, however, we don't have write access to our own folder (because the user
            ' probably stuck us in an access-restricted folder).  When this happens, we disable all update
            ' options and use the explanation label to explain "why".
            If UserPrefs.IsNonPortableModeActive() Then
            
                'This is a non-portable install.  Disable all update controls, then explain why.
                For i = cboUpdates.lBound() To cboUpdates.UBound()
                    cboUpdates(i).Enabled = False
                Next i
                
                For i = chkUpdates.lBound() To chkUpdates.UBound()
                    chkUpdates(i).Enabled = False
                Next i
                
                lblExplanation.Caption = g_Language.TranslateMessage("You have placed PhotoDemon in a restricted system folder.  Security precautions prevent PhotoDemon from modifying this folder, so automatic updates are now disabled.  To restore them, you must move PhotoDemon to a non-admin folder, like Desktop, Documents, or Downloads." & vbCrLf & vbCrLf & "(If you leave PhotoDemon where it is, please don't forget to visit photodemon.org from time to time to check for new versions.)")
                
            'This is a normal (portable) install.  Populate the network access disclaimer in the "Update" panel.
            Else
                lblExplanation.Caption = g_Language.TranslateMessage("The developers of PhotoDemon take privacy very seriously, so no information - statistical or otherwise - is uploaded during the update process.  Updates simply involve downloading several small XML files from photodemon.org. These files contain the latest software, plugin, and language version numbers. If updated versions are found, and user preferences allow, the updated files are then downloaded and patched automatically." & vbCrLf & vbCrLf & "If you still choose to disable updates, don't forget to visit photodemon.org from time to time to check for new versions.")
            End If
            
        'END explanation of update options
    
    'END Update preferences
    
    '***************************************************************************
    
    'START Advanced preferences
        
        'Display the current program settings folder.  (This is normally a subfolder inside the PD folder,
        ' unless the user has done something dumb like install us to a restricted folder.)
        lblSettingsFolder.Caption = UserPrefs.GetDataPath()
        
        'By default, debug logs are only generated in developer and beta builds.  As of 7.2, this behavior can be forcibly
        ' toggled in production builds as well.
        btsDebug.AddItem "auto", 0
        btsDebug.AddItem "no", 1
        btsDebug.AddItem "yes", 2
        btsDebug.ListIndex = UserPrefs.GetPref_Long("Core", "GenerateDebugLogs", 0)
        btsDebug.AssignTooltip "In developer builds, debug data is automatically logged to the program's \Data\Debug folder.  If you encounter bugs in a stable release, please manually activate this setting.  This will help developers resolve your problem."
        
        'High-res mouse input only needs to be deactivated if there are obvious glitches.  This is a Windows-level
        ' problem that seems to show up on VMs and Remote Desktop (see https://forums.getpaint.net/topic/28852-line-jumpsskips-to-top-of-window-while-drawing/)
        btsMouseHighRes.AddItem "off", 0
        btsMouseHighRes.AddItem "on", 1
        btsMouseHighRes.AssignTooltip "When using Remote Desktop or a VM (Virtual Machine), high-resolution mouse input may not work correctly.  This is a long-standing Windows bug.  In these situations, you can use this setting to restore correct mouse behavior."
        If UserPrefs.GetPref_Boolean("Tools", "HighResMouseInput", True) Then btsMouseHighRes.ListIndex = 1 Else btsMouseHighRes.ListIndex = 0
        
        'Display what we know about PD's memory usage
        lblMemoryUsageCurrent.Caption = g_Language.TranslateMessage("current PhotoDemon memory usage:") & " " & Format$(Str$(OS.AppMemoryUsage()), "###,###,###,###") & " K"
        lblMemoryUsageMax.Caption = g_Language.TranslateMessage("max PhotoDemon memory usage this session:") & " " & Format$(Str$(OS.AppMemoryUsage(True)), "###,###,###,###") & " K"
        If (Not OS.IsProgramCompiled) Then
            lblMemoryUsageCurrent.Caption = lblMemoryUsageCurrent.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
            lblMemoryUsageMax.Caption = lblMemoryUsageMax.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
        End If
            
        'Display the current temporary file path
        txtTempPath.Text = UserPrefs.GetTempPath
        cmdTmpPath.AssignTooltip "Click to select a new temporary folder."
        
        'Clarify the behavior of the "reset" button
        cmdReset.AssignTooltip "This button resets all PhotoDemon settings.  If the program is behaving unexpectedly, this may resolve the problem."
        
    'END Advanced preferences
    
    '***************************************************************************
    
    'All preference controls are now initialized with the matching value stored in the preferences file
    
End Sub

'When new transparency checkerboard colors are selected, change the corresponding list box to match
Private Sub csAlphaOne_ColorChanged()
    
    If m_userInitiatedAlphaSelection Then
        m_userInitiatedAlphaSelection = False
        cboAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        m_userInitiatedAlphaSelection = True
    End If
    
End Sub

Private Sub csAlphaTwo_ColorChanged()
    
    If m_userInitiatedAlphaSelection Then
        m_userInitiatedAlphaSelection = False
        cboAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        m_userInitiatedAlphaSelection = True
    End If
    
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the preferences file
Private Sub Form_Load()
    
    Dim i As Long
    
    'Populate all controls with the corresponding values from the preferences file
    If PDMain.IsProgramRunning() Then LoadAllPreferences
    
    'Prep the category button strip
    With btsvCategory
        
        'Start by adding captions for each button.  This will also update the control's layout to match.
        .AddItem "Interface", 0
        .AddItem "Loading", 1
        .AddItem "Saving", 2
        .AddItem "Performance", 3
        .AddItem "Color management", 4
        .AddItem "Updates", 5
        .AddItem "Advanced", 6
        
        'Next, add images to each button
        Dim prefButtonSize As Long
        prefButtonSize = FixDPI(32)
        .AssignImageToItem 0, "pref_interface", , prefButtonSize, prefButtonSize
        .AssignImageToItem 1, "pref_loading", , prefButtonSize, prefButtonSize
        .AssignImageToItem 2, "pref_saving", , prefButtonSize, prefButtonSize
        .AssignImageToItem 3, "pref_performance", , prefButtonSize, prefButtonSize
        .AssignImageToItem 4, "pref_colormanagement", , prefButtonSize, prefButtonSize
        .AssignImageToItem 5, "pref_updates", , prefButtonSize, prefButtonSize
        .AssignImageToItem 6, "pref_advanced", , prefButtonSize, prefButtonSize
        
    End With
    
    'Hide all category panels (the proper one will be activated in a moment)
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Activate the last preferences panel that the user looked at
    Dim activePanel As Long
    activePanel = UserPrefs.GetPref_Long("Core", "Last Preferences Page", 0)
    If (activePanel > picContainer.UBound) Then activePanel = picContainer.UBound
    picContainer(activePanel).Visible = True
    btsvCategory.ListIndex = activePanel
    
    'Apply translations and visual themes
    Interface.ApplyThemeAndTranslations Me
    
    m_userInitiatedColorSelection = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the selected temp folder doesn't have write access, warn the user
Private Sub txtTempPath_Change()
    
    'Assign theme-specific error colors
    If Me.Visible Then
        lblTempPathWarning.ForeColor = g_Themer.GetGenericUIColor(UI_ErrorRed)
        lblTempPathWarning.UseCustomForeColor = True
    End If
    
    'Ensure the specified temp path exists.  If it doesn't (or if access is denied), let the user know that we will silently
    ' fall back to the system temp folder.
    If (Not Files.PathExists(Trim$(txtTempPath.Text), True)) Then
        lblTempPathWarning.Caption = g_Language.TranslateMessage("WARNING: this folder is invalid (access prohibited).  Please provide a valid folder.  If a new folder is not provided, PhotoDemon will use the system temp folder.")
        lblTempPathWarning.Visible = True
    Else
        lblTempPathWarning.Caption = g_Language.TranslateMessage("This new temporary folder location will not take effect until you restart the program.")
        lblTempPathWarning.Visible = Strings.StringsNotEqual(Trim$(txtTempPath.Text), UserPrefs.GetTempPath, True)
    End If
    
End Sub
