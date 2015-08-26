VERSION 5.00
Begin VB.Form FormPrintNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print image"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9975
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
   ScaleHeight     =   571
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   7875
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8460
      TabIndex        =   8
      Top             =   7875
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   7875
      Width           =   1725
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Index           =   0
      Left            =   120
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   2
      Top             =   960
      Width           =   9735
      Begin VB.PictureBox picPrintJobSample 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2550
         Left            =   2400
         ScaleHeight     =   170
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   362
         TabIndex        =   19
         Top             =   4080
         Width           =   5430
      End
      Begin PhotoDemon.smartOptionButton optPrintJob 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   2520
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "one image per page"
         Value           =   -1  'True
      End
      Begin PhotoDemon.textUpDown tudCopies 
         Height          =   345
         Left            =   5280
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Min             =   1
         Max             =   256
         Value           =   1
      End
      Begin VB.ComboBox cmbQuality 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "File_PrintModern.frx":0000
         Left            =   480
         List            =   "File_PrintModern.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1455
         Width           =   4335
      End
      Begin VB.ComboBox cmbPaperSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   4335
      End
      Begin VB.ComboBox cmbPrinter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin PhotoDemon.smartOptionButton optPrintJob 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   3000
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "multiple images per page"
      End
      Begin PhotoDemon.smartOptionButton optPrintJob 
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   18
         Top             =   3480
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "one image spread across multiple pages (poster print)"
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "type of print job"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1725
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of copies"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   5040
         TabIndex        =   13
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "print quality"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "paper size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -5640
      TabIndex        =   10
      Top             =   7800
      Width           =   17415
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In the next step, you can specify detailed layout information (margins, positioning, etc)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   9240
   End
   Begin VB.Label lblWizardTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1 of 2: basic print settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3810
   End
End
Attribute VB_Name = "FormPrintNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Combined Print / Print Preview Interface
'Copyright 2003-2015 by Tanner Helland
'Created: 4/April/03
'Last updated: 12/November/13
'Last update: rewritten from scratch.  Literally.
'
'Printing images is an inherently unpleasant process.  With this new print dialog, I hope to make it less painful.
'
'After prototyping a whole swath of potential layouts, I simply couldn't find a way to condense everything to one
' page without it being a complete UX disaster.  Instead, I have chosen to separate the print process into two
' pages.  This is still better than old "Print preview / print" paradigm, where each dialog is separate.  The
' user should not be forced to switch between dialogs just to make sure their image printed correctly.
'
'In the first page of the new print wizard, the user is asked for the basic print settings that define the rest
' of the process: most significantly, the printer and page size.  Until we know these items, we cannot provide
' other layout-related options.
'
'In the second step, the user is given additional layout options relevant to their selected print situation.
' I find this preferable to trying to shoehorn all layout varieties into a single form.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'These arrays store paper size information, specifically: names, IDs, and exact dimensions (in mm)
Private paperSizeNames() As String
Private paperSizeIDs() As Integer
Private paperSizeExact() As POINTAPI

'To help orient the user, sample output is provided for the three types of print jobs.  These pre-rendered images
' are stored in the resource section of the executable, and we load them to DIBs at run-time.
Private sampleOneImageOnePage As pdDIB
Private sampleMultipleImagesOnePage As pdDIB
Private sampleOneImageMultiplePages As pdDIB

Private Sub cmbPrinter_Click()
    updatePaperSizeList
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    'Start by preparing all relevant combo boxes on the form
    
    'First, populate all available printers
    Dim i As Long
    For i = 0 To Printers.Count - 1
        cmbPrinter.AddItem Printers(i).DeviceName, i
    Next i

    'Pre-select the default printer
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = Printer.DeviceName Then
            cmbPrinter.ListIndex = i
            Exit For
        End If
    Next i
    
    'Fill our paper size arrays with paper sizes supported by the current printer
    updatePaperSizeList
    
    'Populate the quality combo box
    cmbQuality.AddItem "draft (least ink)", 0
    cmbQuality.AddItem "good", 1
    cmbQuality.AddItem "better", 2
    cmbQuality.AddItem "best (most ink)", 3
    cmbQuality.ListIndex = 3
    
    'Load the print job type sample images from the resource file
    Set sampleOneImageOnePage = New pdDIB
    Set sampleMultipleImagesOnePage = New pdDIB
    Set sampleOneImageMultiplePages = New pdDIB
    loadResourceToDIB "PRNT_1IMAGE", sampleOneImageOnePage
    loadResourceToDIB "PRNT_MLTIMGS", sampleMultipleImagesOnePage
    loadResourceToDIB "PRNT_MLTPGS", sampleOneImageMultiplePages
    
    'Display the relevant image for the selected option button
    updatePrintTypeSampleImage
    
End Sub

'The bulk of this function is handled by the matching function in the Printer module
Private Sub updatePaperSizeList()

    'Retrieve all paper sizes
    getPaperSizes cmbPrinter.ListIndex, paperSizeNames, paperSizeIDs, paperSizeExact
    
    'Clear the combo box, then populate it with the new list of paper sizes
    cmbPaperSize.Clear
    
    Dim i As Long
    For i = 0 To UBound(paperSizeNames)
        cmbPaperSize.AddItem paperSizeNames(i) & " : " & paperSizeIDs(i), i
    Next i
    
    'Select the default paper size, which can possibly be obtained from the printer, but in my testing
    ' is always index 0.
    cmbPaperSize.ListIndex = 0

End Sub

'When the user selects a different type of print job, we display a pre-rendered sample image to help them understand
' what the various job types actually do.
Private Sub updatePrintTypeSampleImage()

    Dim xOffset As Long
    xOffset = (picPrintJobSample.ScaleWidth - sampleOneImageOnePage.getDIBWidth) \ 2

    If optPrintJob(0) Then
        sampleOneImageOnePage.renderToPictureBox picPrintJobSample
    
    ElseIf optPrintJob(1) Then
        sampleMultipleImagesOnePage.renderToPictureBox picPrintJobSample
    
    Else
        picPrintJobSample.Picture = LoadPicture("")
        sampleOneImageMultiplePages.alphaBlendToDC picPrintJobSample.hDC, 255, xOffset, 0
        picPrintJobSample.Picture = picPrintJobSample.Image
        picPrintJobSample.Refresh
    
    End If

End Sub

Private Sub optPrintJob_Click(Index As Integer)
    updatePrintTypeSampleImage
End Sub

Private Sub testCurrentPrintSettings()

End Sub
