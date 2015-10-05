VERSION 5.00
Begin VB.Form FormPrint 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print Image"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8745
   ControlBox      =   0   'False
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
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.smartCheckBox chkCenter 
      Height          =   330
      Left            =   4080
      TabIndex        =   22
      Top             =   3480
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   582
      Caption         =   "center on page"
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   6030
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7230
      TabIndex        =   1
      Top             =   6030
      Width           =   1365
   End
   Begin VB.PictureBox picOut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   360
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbDPI 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox picThumbFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picThumb90 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cbOrientation 
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
      ItemData        =   "File_PrintXP.frx":0000
      Left            =   4080
      List            =   "File_PrintXP.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2550
      Width           =   2610
   End
   Begin VB.ComboBox cbQuality 
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
      ItemData        =   "File_PrintXP.frx":0004
      Left            =   4080
      List            =   "File_PrintXP.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1575
      Width           =   2610
   End
   Begin VB.ComboBox cbPrinters 
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
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtCopies 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   1575
      Width           =   1335
   End
   Begin VB.PictureBox iSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   240
      ScaleHeight     =   258
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   3300
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin PhotoDemon.smartCheckBox chkFit 
      Height          =   330
      Left            =   4080
      TabIndex        =   23
      Top             =   3960
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   582
      Caption         =   "fit on page"
   End
   Begin VB.Label lblQuality 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "quality"
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
      Left            =   3960
      TabIndex        =   21
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   20
      Top             =   5880
      Width           =   8895
   End
   Begin VB.Label lblDPIWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "note: DPI is read-only when ""fit to page"" is selected."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   5640
      TabIndex        =   18
      Top             =   4980
      Width           =   3015
   End
   Begin VB.Label lblDPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dpi (print resolution)"
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
      Left            =   3960
      TabIndex        =   17
      Top             =   4560
      Width           =   2205
   End
   Begin VB.Label lblPaperSize 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   3300
   End
   Begin VB.Label lblLayoutOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "layout options"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lblOrientation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "orientation"
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
      Left            =   3945
      TabIndex        =   10
      Top             =   2205
      Width           =   1140
   End
   Begin VB.Label lblPrinter 
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
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   705
   End
   Begin VB.Label lblCopies 
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
      Left            =   6840
      TabIndex        =   8
      Top             =   1200
      Width           =   1200
   End
End
Attribute VB_Name = "FormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Printer Interface (including Print Preview)
'Copyright 2000-2015 by Tanner Helland
'Created: 4/April/03
'Last updated: 26/June/12
'Last update: redesign from the ground up.  Print previewing via FreeImage. Manual and automatic DPI calculation support.
'              Paper size interface via PageSetupDlg.  All these calculations interact with each other, resulting in a true
'              "print preview" - WYSIWYG when it comes to this form.
'Still needs: Internal paper size support via EnumForms.  At present, PageSetupDlg is used to handle paper size, but it's
'              an inelegant solution at best.  One of its biggest problems is that it doesn't restrict page sizes to ones
'              supported by a given printer.  Lame.  The best solution would be a custom-built page-select feature that
'              doesn't require the user to spawn an additional window just to change paper sizes.
'              Note that in XP+, Windows no longer allows the Printer object to set arbitrary width and height values.
'              Programs must adhere to existing page sizes as specified by EnumForms, and if they *do* decide to add custom
'              forms, they must add them to the registry so all applications can access them.  It's a mess.  That said, here
'              are some resources that may help if I ever decide to implement paper size in the future:
'              http://support.microsoft.com/kb/282474/t
'              http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.powerpacks.printing.compatibility.vb6.printer.papersize.aspx
'              http://www.vbforums.com/showthread.php?t=451198
'
'Module for interfacing with the printer.  For a program that's not designed around printing, PhotoDemon's interface is
' surprisingly robust.  All settings are handled through the default VB printer object.  A key feature of this routine
' (PrintPictureToFitPage) is based off code first written by Waty Thierry (http://www.ppreview.net/).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'These usually exist in a common dialog object, but since we're manually handling all
' printer functionality we need to manually specify them
Public Enum PrinterOrientationConstants
    cdlLandscape = 2
    cdlPortrait = 1
End Enum

'Base DPI defaults to the current screen DPI (VB uses this as the default when sending an image to the printer via PaintPicture.)
' We use a variety of tricky math to adjust this, and in turn adjust the print quality of the image.
Private baseDPI As Double, desiredDPI As Double

'Changing the orientation box forces a refresh of the preview
Private Sub cbOrientation_Click()
    UpdatePrintPreview True
End Sub

Private Sub cbOrientation_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdatePrintPreview True
End Sub

'Allow the user to change the target printer
Private Sub cbPrinters_Click()

    Dim Prt As Printer
    
    For Each Prt In Printers
        If (Prt.DeviceName = cbPrinters) Then
            Set Printer = Prt
        End If
    Next Prt
    
    UpdatePaperSize
    UpdatePrintPreview True

End Sub

Private Sub cbPrinters_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim Prt As Printer
    For Each Prt In Printers
        If (Prt.DeviceName = cbPrinters) Then
            Set Printer = Prt
        End If
    Next Prt
    UpdatePaperSize
    UpdatePrintPreview

End Sub

'When various options are selected or de-selected, we need to update the preview to reflect it
Private Sub chkCenter_Click()
    UpdatePrintPreview True
End Sub

Private Sub chkCenter_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdatePrintPreview True
End Sub

Private Sub chkFit_Click()
    UpdatePrintPreview
    If chkFit.Value = vbUnchecked Then cmbDPI = baseDPI
End Sub

Private Sub chkFit_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdatePrintPreview
    If chkFit.Value = vbUnchecked Then cmbDPI = baseDPI
End Sub

Private Sub cmbDPI_Click()
    If EntryValid(cmbDPI, 1, 12000, False, False) Then
        desiredDPI = cmbDPI
        UpdatePrintPreview True
    End If
End Sub

Private Sub cmbDPI_KeyUp(KeyCode As Integer, Shift As Integer)
    If EntryValid(cmbDPI, 1, 12000, False, False) Then
        desiredDPI = cmbDPI
        UpdatePrintPreview True
    End If
End Sub

Private Sub cmbDPI_Scroll()
    If EntryValid(cmbDPI, 1, 12000, False, False) Then
        desiredDPI = cmbDPI
        UpdatePrintPreview True
    End If
End Sub

'LOAD form
Private Sub Form_Load()

    UpdatePaperSize
    
    'Though it's not really necessary, I'm only enabling print preview if FreeImage is Enabled (we use FreeImage to perform
    ' fast, high-quality rotations of images).  I anticipate that pretty much no one will ever use this printing option, so
    ' I don't mind ignoring a VB-only fallback for such a peripheral feature.
    If g_ImageFormats.FreeImageEnabled = True Then
        RebuildPreview
        lblWarning.Visible = False
    Else
        lblWarning.Caption = g_Language.TranslateMessage("Print previewing requires the FreeImage plugin, which could not be located on this computer. To enable previewing, please go to Edit -> Preferences and select ""check for missing plugins on program start.""  The next time you load PhotoDemon, it will offer to download this plugin for you.")
        lblWarning.Visible = True
    End If
    
    Dim x As Long
    
    'Load a list of printers into the combo box
    For x = 0 To Printers.Count - 1
        cbPrinters.AddItem Printers(x).DeviceName
    Next x

    'Pre-select the current printer
    cbPrinters = Printer.DeviceName

    'Set a default quality (high, since we're working with images)
    cbQuality.AddItem "draft", 0
    cbQuality.AddItem "low", 1
    cbQuality.AddItem "medium", 2
    cbQuality.AddItem "high", 3
    cbQuality.ListIndex = 3

    'Set image orientation based on its aspect ratio (as compared to an 8.5" x 11" sheet of paper)
    cbOrientation.AddItem "portrait", 0
    cbOrientation.AddItem "landscape", 1
    
    Dim imgAspect As Double, paperAspect As Double
    imgAspect = pdImages(g_CurrentImage).Width / pdImages(g_CurrentImage).Height
    paperAspect = 8.5 / 11
    
    If imgAspect < paperAspect Then
        cbOrientation.ListIndex = 0
    Else
        cbOrientation.ListIndex = 1
    End If
    
    UpdatePrintPreview

    'Temporarily copy the image into an image box
    picOut.Width = pdImages(g_CurrentImage).Width
    picOut.Height = pdImages(g_CurrentImage).Height
    picOut.ScaleMode = vbPixels
    
    Dim tmpComposite As pdDIB
    Set tmpComposite = New pdDIB
    pdImages(g_CurrentImage).getCompositedImage tmpComposite
    tmpComposite.renderToPictureBox picOut
    
    picOut.ScaleMode = vbTwips
    
    'Determine base DPI (should be screen DPI, but calculate it manually to be sure)
    Dim pic As StdPicture
    Set pic = New StdPicture
    Set pic = picOut.Picture
    
    Dim tPrnPicWidth As Double, tPrnPicHeight As Double
    tPrnPicWidth = Printer.scaleX(pic.Width, vbHiMetric, Printer.ScaleMode)
    tPrnPicHeight = Printer.scaleY(pic.Height, vbHiMetric, Printer.ScaleMode)
    Dim dpiX As Double, dpiY As Double
    dpiX = CSng(pdImages(g_CurrentImage).Width) / Printer.scaleX(tPrnPicWidth, Printer.ScaleMode, vbInches)
    dpiY = CSng(pdImages(g_CurrentImage).Height) / Printer.scaleY(tPrnPicHeight, Printer.ScaleMode, vbInches)
    baseDPI = Int((dpiX + dpiY) / 2 + 0.5)
    desiredDPI = baseDPI
    
    'Populate the DPI combo box with some suggested default values.  The user is free to provide custom values as well.
    cmbDPI.AddItem "72"
    cmbDPI.AddItem "144"
    cmbDPI.AddItem "300"
    cmbDPI.AddItem "600"
    cmbDPI.AddItem "1200"
    cmbDPI.AddItem "2400"
    cmbDPI.AddItem "3600"
    cmbDPI.AddItem "4000"
    cmbDPI = baseDPI
    
    'Apply translations and visual themes
    MakeFormPretty Me

    If g_UseFancyFonts Then txtCopies.Height = txtCopies.Height + 1
    
End Sub

'OK Button
Private Sub CmdOK_Click()

    On Error Resume Next
    
    Message "Sending image to printer..."
    
    'Before printing anything, check to make sure the textboxes have valid input
    If Not (NumberValid(txtCopies.Text) And RangeValid(Val(txtCopies.Text), 1, 1000)) Then
        AutoSelectText txtCopies
        Exit Sub
    End If

    'Set the number of copies
    Printer.Copies = Val(txtCopies.Text)
      
    'Assuming there have been no errors thus far (basically, assuming the user actually has a printer attatched)...
    If (Err = 0) Then
      
        'Set print quality
        Printer.PrintQuality = -(cbQuality.ListIndex + 1)
          
        'Assuming our quality option is valid (should be, but you never know)
        If (Err = 0) Then
          
            'Print the image
            If (PrintPictureToFitPage(Printer, picOut.Picture, cbOrientation.ListIndex + 1, CBool(chkCenter), CBool(chkFit)) = 0) Then
                PDMsgBox "%1 was unable to print the image.  Please make sure that the specified printer (%2) is powered-on and ready for printing.", vbExclamation + vbOKOnly + vbApplicationModal, "Printer Error", PROGRAMNAME, Printer.DeviceName
                Message "Print canceled."
            End If
              
        End If
    End If
        
    Printer.EndDoc

    Message "Image printed successfully. (Note: depending on your printer, additional print confirmation screens may appear.)"

    Unload Me

End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'This PrintPictureToFitPage function is based off code originally written by Waty Thierry.
' (It has been heavily modified for use within PhotoDemon, but you may download the original at http://www.freevbcode.com/ShowCode.asp?ID=194)
Public Function PrintPictureToFitPage(Prn As Printer, pic As StdPicture, ByVal iOrientation As PrinterOrientationConstants, Optional ByVal iCenter As Boolean = True, Optional ByVal iFit As Boolean = True) As Boolean

    Const vbHiMetric As Integer = 8

    Dim e As Long
    Dim PicRatio As Double
    Dim PrnWidth As Double, PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double, PrnPicHeight As Double

    Dim offsetX As Double, offsetY As Double

    On Error Resume Next

    'Set the printer orientation to match the orientation we were handed
    Prn.Orientation = iOrientation
    e = Err

    'Calculate an aspect ratio for the image
    PicRatio = pic.Width / pic.Height
    e = e Or Err

    'Calculate the printable dimensions (affected by page size)
    PrnWidth = Prn.scaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.scaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)

    'Calculate aspect ratio for the printed page
    PrnRatio = PrnWidth / PrnHeight
    e = e Or Err

    'If the user has not selected "fit image to page," calculate the printable size
    If (Not iFit) Then
        Dim dpiRatio As Double
        dpiRatio = baseDPI / desiredDPI
        PrnPicWidth = Prn.scaleX(pic.Width, vbHiMetric, Prn.ScaleMode) * dpiRatio
        PrnPicHeight = Prn.scaleY(pic.Height, vbHiMetric, Prn.ScaleMode) * dpiRatio
    
    'Otherwise, set the printable size and image size to the same dimensions
    Else
        If (PicRatio >= PrnRatio) Then
            PrnPicWidth = Prn.scaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.scaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
        Else
            PrnPicHeight = Prn.scaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.scaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
        End If
    End If

    'If the user has told us to center the image, calculate offsets
    If (iCenter) Then
        offsetX = (Prn.ScaleWidth - PrnPicWidth) \ 2
        offsetY = (Prn.ScaleHeight - PrnPicHeight) \ 2
    End If

    'Print the picture using VB's PaintPicture method
    Prn.PaintPicture pic, offsetX, offsetY, PrnPicWidth, PrnPicHeight
    e = e Or Err

    PrintPictureToFitPage = (e = 0)
    On Error GoTo 0

End Function

'Redraw the Print Preview box based on the specified image orientation
Private Sub UpdatePrintPreview(Optional forceDPI As Boolean = False)
    
    If chkFit.Value = vbUnchecked Then
        lblDPIWarning.Visible = False
        cmbDPI.Enabled = True
    Else
        lblDPIWarning.Visible = True
        cmbDPI.Enabled = False
    End If
    
    If g_ImageFormats.FreeImageEnabled = False Then Exit Sub
    
    'If the fit-to-page option is selected (which it is by default) this routine is very simple:
    If chkFit.Value = vbChecked Then
        If cbOrientation.ListIndex = 0 Then
            iSrc.Picture = picThumb.Picture
            iSrc.Refresh
            UpdateDPI CSng(pdImages(g_CurrentImage).Width) / Printer.scaleX(Printer.Width, Printer.ScaleMode, vbInches)
        Else
            iSrc.Picture = picThumbFinal.Picture
            iSrc.Refresh
            UpdateDPI CSng(pdImages(g_CurrentImage).Height) / Printer.scaleX(Printer.Width, Printer.ScaleMode, vbInches)
        End If
        
        Exit Sub
        
    End If
    
    'If the fit-to-page option is not selected, things get a bit hairier.
    'Note that this code is heavily derived from the PrintPictureToFitPage routine above, by Waty Thierry.
    Const vbHiMetric  As Integer = 8

    Dim PrnWidth      As Double
    Dim PrnHeight     As Double
    Dim PrnPicWidth   As Double
    Dim PrnPicHeight  As Double
    
    Dim offsetX       As Double
    Dim offsetY       As Double

    On Error Resume Next

    Printer.Orientation = cbOrientation.ListIndex + 1
    Printer.PrintQuality = -(cbQuality.ListIndex + 1)
    
    'Calculate the dimensions of the printable area in HiMetric
    PrnWidth = Printer.scaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHiMetric)
    PrnHeight = Printer.scaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHiMetric)

    Dim pic As StdPicture
    Set pic = picOut.Picture

    PrnPicWidth = Printer.scaleX(pic.Width, vbHiMetric, Printer.ScaleMode)
    PrnPicHeight = Printer.scaleY(pic.Height, vbHiMetric, Printer.ScaleMode)
    
    'Estimate DPI
    Dim dpiRatio As Double
    If forceDPI = False Then
        Dim dpiX As Double, dpiY As Double
        dpiX = CSng(pdImages(g_CurrentImage).Width) / Printer.scaleX(PrnPicWidth, Printer.ScaleMode, vbInches)
        dpiY = CSng(pdImages(g_CurrentImage).Height) / Printer.scaleY(PrnPicHeight, Printer.ScaleMode, vbInches)
        UpdateDPI ((dpiX + dpiY) / 2)
        dpiRatio = 1
    Else
        'Calculate a DPI ratio
        dpiRatio = baseDPI / desiredDPI
        PrnPicWidth = PrnPicWidth * dpiRatio
        PrnPicHeight = PrnPicHeight * dpiRatio
    End If
    
    If chkCenter.Value = vbChecked Then
        offsetX = (Printer.ScaleWidth - PrnPicWidth) \ 2
        offsetY = (Printer.ScaleHeight - PrnPicHeight) \ 2
    End If
    
    'Now, convert the printer-specific measurements to their corresponding measurements in the preview window
    If cbOrientation.ListIndex = 0 Then
        offsetX = (offsetX / Printer.ScaleWidth) * iSrc.ScaleWidth
        offsetY = (offsetY / Printer.ScaleHeight) * iSrc.ScaleHeight
        PrnPicWidth = (PrnPicWidth / Printer.ScaleWidth) * iSrc.ScaleWidth
        PrnPicHeight = (PrnPicHeight / Printer.ScaleHeight) * iSrc.ScaleHeight
    Else
        Dim tmpOX As Double, tmpOY As Double, tmpWidth As Double, tmpHeight As Double
        tmpOX = (offsetY / Printer.ScaleHeight) * iSrc.ScaleWidth
        tmpOY = (offsetX / Printer.ScaleWidth) * iSrc.ScaleHeight
        tmpWidth = (PrnPicHeight / Printer.ScaleHeight) * iSrc.ScaleWidth
        tmpHeight = (PrnPicWidth / Printer.ScaleWidth) * iSrc.ScaleHeight
        offsetX = tmpOX
        offsetY = tmpOY
        PrnPicWidth = tmpWidth
        PrnPicHeight = tmpHeight
    End If
    
    'TODO!  Rewrite this whole dialog.  It needs a ton of help.
    
    'Draw a new preview
    If cbOrientation.ListIndex = 0 Then
        DrawPreviewImage picThumb, , , True
        iSrc.Picture = LoadPicture("")
        SetStretchBltMode iSrc.hDC, STRETCHBLT_HALFTONE
        'StretchBlt iSrc.hDC, offsetX, offsetY, PrnPicWidth, PrnPicHeight, picThumb.hDC, pdImages(g_CurrentImage).getOldCompositedImage().previewX, pdImages(g_CurrentImage).getOldCompositedImage().previewY, pdImages(g_CurrentImage).getOldCompositedImage().previewWidth, pdImages(g_CurrentImage).getOldCompositedImage().previewHeight, vbSrcCopy
    Else
        DrawPreviewImage picThumb90, , , True
        iSrc.Picture = LoadPicture("")
        SetStretchBltMode iSrc.hDC, STRETCHBLT_HALFTONE
        'StretchBlt iSrc.hDC, offsetX, offsetY, PrnPicWidth, PrnPicHeight, picThumbFinal.hDC, pdImages(g_CurrentImage).getOldCompositedImage().previewY, pdImages(g_CurrentImage).getOldCompositedImage().previewX, pdImages(g_CurrentImage).getOldCompositedImage().previewHeight, pdImages(g_CurrentImage).getOldCompositedImage().previewWidth, vbSrcCopy
    End If
    
    iSrc.Picture = iSrc.Image
    iSrc.Refresh
      
End Sub

'This is called whenever the dimensions of the preview window change (for example, in response to a change in paper size)
Private Sub RebuildPreview(Optional forceDPI As Boolean = False)
    
    'FreeImage is used to rotate the image; if it's not installed, previewing is automatically disabled
    If g_ImageFormats.FreeImageEnabled Then
    
        'We're now going to create two temporary buffers; one contains the image resized to fit the "sheet of paper" preview
        ' on the left.  This is portrait mode.  The second buffer will contain the same thing, but rotated 90 degrees -
        ' e.g. landscape mode.  If the user clicks between those options, we can simply copy the buffers to the foreground
        ' picture box.
        
        'First is the easy one - Portrait Mode
        picThumb.Picture = LoadPicture("")
        picThumb.Width = iSrc.Width
        picThumb.Height = iSrc.Height
        DrawPreviewImage picThumb, , , True
        
        'Now we need to get the source image at the size expected post-rotation
        picThumb90.Picture = LoadPicture("")
        picThumbFinal.Picture = LoadPicture("")
        picThumb90.Width = iSrc.Height
        picThumb90.Height = iSrc.Width
        picThumbFinal.Width = iSrc.Width
        picThumbFinal.Height = iSrc.Height

        DrawPreviewImage picThumb90, , , True
        
        'Now comes the rotation itself.
        picThumbFinal.Picture = FreeImage_RotateIOP(picThumb90.Picture, 90)
        picThumbFinal.Refresh
        
        'Initiate a redraw of the preview according to the print settings currently specified by the user
        UpdatePrintPreview forceDPI
        
    End If
        
End Sub

'Display what size of paper we expect to output the image onto
Private Sub UpdatePaperSize()
    
    'Note: interacting with the printer is a messy, complex chunk of API hackery (see: http://support.microsoft.com/kb/282474/t)
    ' To that end, I've written this code to work with any paper size - all it requires is that the Printer.Width and Printer.Height
    ' values be set to whatever size the user desires.  Thus this code is abstracted away from the actual paper size selection
    ' process, and any new selection method will require zero changes here.
    Dim pWidth As Double, pHeight As Double
    pWidth = Printer.scaleX(Printer.Width, Printer.ScaleMode, vbInches)
    pHeight = Printer.scaleY(Printer.Height, Printer.ScaleMode, vbInches)
    Dim TxtWidth As String, TxtHeight As String
    TxtWidth = Format(pWidth, "#0.##")
    TxtHeight = Format(pHeight, "#0.##")
    If Right(TxtWidth, 1) = "." Then TxtWidth = Left$(TxtWidth, Len(TxtWidth) - 1)
    If Right(TxtHeight, 1) = "." Then TxtHeight = Left$(TxtHeight, Len(TxtHeight) - 1)
    
    lblPaperSize = g_Language.TranslateMessage("paper size") & ": " & TxtWidth & """ x  " & TxtHeight & """"
    
    'Now comes the tricky part - we need to resize the preview box to match the aspect ratio of the paper
    Dim aspectRatio As Double
    aspectRatio = pWidth / pHeight
    
    'The maximum available area for the preview is 220 pixels.  The margin on both top and left is 16 pixels.
    ' This means that the full workable area is 220 + 16 * 2, or 252 in either direction.
    ' We will rebuild the preview picture box according to the new aspect ratio and these size constraints.
    
    'If width is larger than height, assign width a size of 220 pixels and adjust height accordingly
    If aspectRatio >= 1 Then
        iSrc.Width = 220
        iSrc.Left = 16
        iSrc.Height = 220 * aspectRatio
        iSrc.Top = 16 + (260 - iSrc.Height) / 2
    Else
        iSrc.Height = 220
        iSrc.Top = 16
        iSrc.Width = 220 * aspectRatio
        iSrc.Left = 16 + (220 - iSrc.Width) / 2
    End If
    
    'Automatically initiate a rebuild of the preview picture boxes
    If chkFit.Value = vbChecked Then RebuildPreview Else RebuildPreview True
    
End Sub

'Display the assumed DPI that results from the current print settings
Private Sub UpdateDPI(ByVal eDPI As Double)
    cmbDPI = Int(eDPI + 0.5)
End Sub

Private Sub cmbDPI_Change()
    If EntryValid(cmbDPI, 1, 12000, False, False) Then
        desiredDPI = cmbDPI
        UpdatePrintPreview True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub txtCopies_GotFocus()
    AutoSelectText txtCopies
End Sub
