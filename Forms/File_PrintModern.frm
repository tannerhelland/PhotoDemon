VERSION 5.00
Begin VB.Form FormPrintNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print"
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
   Begin PhotoDemon.pdButton cmdNext 
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   7875
      Width           =   1725
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   8460
      TabIndex        =   8
      Top             =   7875
      Width           =   1365
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Cancel"
   End
   Begin PhotoDemon.pdButton cmdPrevious 
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   7875
      Width           =   1725
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "&Previous"
      Enabled         =   0   'False
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
         TabIndex        =   0
         Top             =   4080
         Width           =   5430
      End
      Begin PhotoDemon.pdRadioButton optPrintJob 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   2520
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "one image per page"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudCopies 
         Height          =   345
         Left            =   5280
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         DefaultValue    =   1
         Min             =   1
         Max             =   256
         Value           =   1
      End
      Begin PhotoDemon.pdDropDown cmbQuality 
         Height          =   360
         Left            =   480
         TabIndex        =   11
         Top             =   1455
         Width           =   4335
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdDropDown cmbPaperSize 
         Height          =   360
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   4335
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdDropDown cmbPrinter 
         Height          =   360
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   4335
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin PhotoDemon.pdRadioButton optPrintJob 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "multiple images per page"
      End
      Begin PhotoDemon.pdRadioButton optPrintJob 
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   3480
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   582
         Caption         =   "one image spread across multiple pages (poster print)"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   4
         Left            =   240
         Top             =   2040
         Width           =   1725
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "type of print job"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   3
         Left            =   5040
         Top             =   1080
         Width           =   1200
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "# of copies"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   2
         Left            =   240
         Top             =   1080
         Width           =   1275
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "print quality"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   5040
         Top             =   120
         Width           =   1065
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "paper size"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   240
         Top             =   120
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "printer"
      End
   End
   Begin PhotoDemon.pdLabel lblDescription 
      Height          =   285
      Left            =   240
      Top             =   480
      Width           =   9240
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "In the next step, you can specify detailed layout information (margins, positioning, etc)"
   End
   Begin PhotoDemon.pdLabel lblWizardTitle 
      Height          =   285
      Left            =   120
      Top             =   120
      Width           =   3810
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Step 1 of 2: basic print settings"
   End
End
Attribute VB_Name = "FormPrintNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Combined Print / Print Preview Interface
'Copyright 2003-2020 by Tanner Helland
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'These arrays store paper size information, specifically: names, IDs, and exact dimensions (in mm)
Private paperSizeNames() As String
Private paperSizeIDs() As Integer
Private paperSizeExact() As PointAPI

Private Sub cmbPrinter_Click()
    UpdatePaperSizeList
End Sub

Private Sub cmdCancel_Click()
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
    UpdatePaperSizeList
    
    'Populate the quality combo box
    cmbQuality.AddItem "draft (least ink)", 0
    cmbQuality.AddItem "good", 1
    cmbQuality.AddItem "better", 2
    cmbQuality.AddItem "best (most ink)", 3
    cmbQuality.ListIndex = 3
    
End Sub

'The bulk of this function is handled by the matching function in the Printer module
Private Sub UpdatePaperSizeList()

    'Retrieve all paper sizes
    GetPaperSizes cmbPrinter.ListIndex, paperSizeNames, paperSizeIDs, paperSizeExact
    
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

Private Sub TestCurrentPrintSettings()

End Sub
