VERSION 5.00
Begin VB.Form FormPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print Image"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7575
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
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDPI 
      Height          =   315
      Left            =   4320
      MouseIcon       =   "VBP_FormPrint.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3675
      Width           =   1095
   End
   Begin VB.CommandButton cmdPaperSize 
      Caption         =   "Change Paper Size..."
      Height          =   375
      Left            =   3960
      MouseIcon       =   "VBP_FormPrint.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.PictureBox picThumbFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picThumb90 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkFit 
      Appearance      =   0  'Flat
      Caption         =   "Fit to page"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5280
      MouseIcon       =   "VBP_FormPrint.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1170
   End
   Begin VB.CheckBox chkCenter 
      Appearance      =   0  'Flat
      Caption         =   "Center on page"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5280
      MouseIcon       =   "VBP_FormPrint.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.ComboBox cbOrientation 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "VBP_FormPrint.frx":0548
      Left            =   4965
      List            =   "VBP_FormPrint.frx":0552
      MouseIcon       =   "VBP_FormPrint.frx":056B
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   2370
   End
   Begin VB.ComboBox cbQuality 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "VBP_FormPrint.frx":06BD
      Left            =   4575
      List            =   "VBP_FormPrint.frx":06CD
      MouseIcon       =   "VBP_FormPrint.frx":06EB
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1575
      Width           =   960
   End
   Begin VB.ComboBox cbPrinters 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3945
      MouseIcon       =   "VBP_FormPrint.frx":083D
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtCopies 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   1590
      Width           =   390
   End
   Begin VB.PictureBox iSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   240
      ScaleHeight     =   218
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   3300
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   $"VBP_FormPrint.frx":098F
         ForeColor       =   &H00C00000&
         Height          =   1935
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   435
      Left            =   4920
      MouseIcon       =   "VBP_FormPrint.frx":0AB2
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6240
      MouseIcon       =   "VBP_FormPrint.frx":0C04
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4680
      Width           =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   496
      X2              =   8
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   248
      X2              =   248
      Y1              =   8
      Y2              =   280
   End
   Begin VB.Label lblDPIWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: DPI is read-only when ""Fit to Page"" is selected."
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5520
      TabIndex        =   22
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label lblDPI 
      BackStyle       =   0  'Transparent
      Caption         =   "DPI:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblPaperSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size: "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3720
      Width           =   3300
   End
   Begin VB.Label lblLayoutOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Layout Options:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3960
      TabIndex        =   15
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label lblOrientation 
      BackStyle       =   0  'Transparent
      Caption         =   "Orientation"
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3945
      TabIndex        =   14
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblQuality 
      BackStyle       =   0  'Transparent
      Caption         =   "Quality"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3945
      TabIndex        =   13
      Top             =   1620
      Width           =   705
   End
   Begin VB.Image iOut 
      Height          =   615
      Left            =   360
      Top             =   2880
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:"
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblCopies 
      BackStyle       =   0  'Transparent
      Caption         =   "# of Copies:"
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   5985
      TabIndex        =   11
      Top             =   1620
      Width           =   945
   End
End
Attribute VB_Name = "FormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Printer Interface (including Print Preview)
'�2000-2012 Tanner Helland
'Created: 4/April/03
'Last updated: 26/June/12
'Last update: redesign from the ground up.  Print previewing via FreeImage. Manual and automatic DPI calculation support.
'             Paper size interface via PageSetupDlg.  All these calculations interact with each other, resulting in a true
'             "print preview" - WYSIWYG when it comes to this form.
'Still needs: Internal paper size support via EnumForms.  At present, PageSetupDlg is used to handle paper size, but it's
'             an inelegant solution at best.  One of its biggest problems is that it doesn't restrict page sizes to ones
'             supported by a given printer.  Lame.  The best solution would be a custom-built page-select feature that
'             doesn't require the user to spawn an additional window just to change paper sizes.
'             Note that in XP+, Windows no longer allows the Printer object to set arbitrary width and height values.
'             Programs must adhere to existing page sizes as specified by EnumForms, and if they *do* decide to add custom
'             forms, they must add them to the registry so all applications can access them.  It's a mess.  That said, here
'             are some resources that may help if I ever decide to implement paper size in the future:
'             http://support.microsoft.com/kb/282474/t
'             http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.powerpacks.printing.compatibility.vb6.printer.papersize.aspx
'             http://www.vbforums.com/showthread.php?t=451198
'
'Module for interfacing with the printer.  For a program that's not designed around printing, PhotoDemon's interface is
' surprisingly robust.  All settings are handled through the default VB printer object.  A key feature of this routine
' (PrintPictureToFitPage) was first written by Waty Thierry (http://www.ppreview.net/) - many thanks go out to him for
' his great work on printing in VB.
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
Dim baseDPI As Single, desiredDPI As Single

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

'Rather than implement our own paper size mechanic, I cheat and use the default Windows one.  It isn't pretty - but it gets the job done.
Private Sub cmdPaperSize_Click()
    Dim cdReturn As Boolean
    Dim cDialog As cCommonDialog
    Set cDialog = New cCommonDialog
    cdReturn = cDialog.VBPageSetupDlg2(Me.HWnd, True, True, False, False, , , , , , , , , Printer.PaperSize, cbOrientation.ListIndex + 1, -(cbQuality.ListIndex + 1), epsuinches, Printer)
    If cdReturn = True Then UpdatePaperSize
End Sub

'LOAD form
Private Sub Form_Load()

    UpdatePaperSize

    GetImageData
    
    'Though it's not really necessary, I'm only enabling print preview if FreeImage is Enabled (we use FreeImage to perform
    ' fast, high-quality rotations of images).  I anticipate that pretty much no one will ever use this printing option, so
    ' I don't mind ignoring a VB-only fallback for such a peripheral feature.
    If FreeImageEnabled = True Then
        RebuildPreview
    Else
        lblWarning.Visible = True
    End If
    
    'Load a list of printers into the combo box
    For x = 0 To Printers.Count - 1
        cbPrinters.AddItem Printers(x).DeviceName
    Next x

    'Pre-select the current printer
    cbPrinters = Printer.DeviceName

    'Set a default quality (high, since we're working with images)
    cbQuality.ListIndex = 3

    'Set image orientation based on its aspect ratio (as compared to an 8.5" x 11" sheet of paper)
    Dim imgAspect As Single, paperAspect As Single
    imgAspect = PicWidthL / PicHeightL
    paperAspect = 8.5 / 11
    
    If imgAspect < paperAspect Then
        cbOrientation.ListIndex = 0
    Else
        cbOrientation.ListIndex = 1
    End If
    
    UpdatePrintPreview

    'Temporarily copy the image into an image box
    iOut.Picture = FormMain.ActiveForm.BackBuffer.Picture
    iOut.Refresh
    
    'Determine base DPI (should be screen DPI, but calculate it manually to be sure)
    Dim pic As StdPicture
    Set pic = iOut
    
    Dim tPrnPicWidth As Single, tPrnPicHeight As Single
    tPrnPicWidth = Printer.ScaleX(pic.Width, vbHiMetric, Printer.ScaleMode)
    tPrnPicHeight = Printer.ScaleY(pic.Height, vbHiMetric, Printer.ScaleMode)
    Dim dpiX As Double, dpiY As Double
    dpiX = CSng(FormMain.ActiveForm.BackBuffer.ScaleWidth) / Printer.ScaleX(tPrnPicWidth, Printer.ScaleMode, vbInches)
    dpiY = CSng(FormMain.ActiveForm.BackBuffer.ScaleHeight) / Printer.ScaleY(tPrnPicHeight, Printer.ScaleMode, vbInches)
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

End Sub

'OK Button
Private Sub CmdOK_Click()

    On Error Resume Next
    
    Message "Sending image to printer..."
    
    'Before printing anything, check to make sure the textboxes have valid input
    If Not (NumberValid(txtCopies.Text) And RangeValid(val(txtCopies.Text), 1, 1000)) Then
        AutoSelectText txtCopies
        Exit Sub
    End If

    'Set the number of copies
    Printer.Copies = val(txtCopies.Text)
      
    'Assuming there have been no errors thus far (basically, assuming the user actually has a printer attatched)...
    If (Err = 0) Then
      
        'Set print quality
        Printer.PrintQuality = -(cbQuality.ListIndex + 1)
          
        'Assuming our quality option is valid (should be, but you never know)
        If (Err = 0) Then
          
            'Print the image
            If (PrintPictureToFitPage(Printer, iOut, cbOrientation.ListIndex + 1, CBool(chkCenter), CBool(chkFit)) = 0) Then
                MsgBox PROGRAMNAME & " was unable to print the image.  Please make sure that the specified printer (" & Printer.DeviceName & ") is powered-on and ready for printing.", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " Printer Error"
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

'Thanks to Waty Thierry for the original version of this function
' (Note that it has been heavily modified for use within PhotoDemon - you may download the original at http://www.freevbcode.com/ShowCode.asp?ID=194)
Public Function PrintPictureToFitPage(Prn As Printer, pic As StdPicture, ByVal iOrientation As PrinterOrientationConstants, Optional ByVal iCenter As Boolean = True, Optional ByVal iFit As Boolean = True) As Boolean

  ' #VBIDEUtils#***************************************************
  ' * Programmer Name  : Waty Thierry
  ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
  ' * E-Mail           : waty.thierry@usa.net
  ' * Date             : 13/Oct/98
  ' * Time             : 09:18
  ' * Module Name      : Capture_Module
  ' * Module Filename  : Capture.bas
  ' * Procedure Name   : PrintPictureToFitPage
  ' * Parameters       : Prn As Printer
  ' *                    Pic As Picture
  ' ***************************************************************
  ' * Comments         : Prints a Picture object as big as possible
  ' *
  ' ***************************************************************

    Const vbHiMetric  As Integer = 8

    Dim e As Long
    Dim PicRatio As Double
    Dim PrnWidth As Double, PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double, PrnPicHeight As Double

    Dim OffsetX As Double, OffsetY As Double

    On Error Resume Next

    'Set the printer orientation to match the orientation we were handed
    Prn.Orientation = iOrientation
    e = Err

    'Calculate an aspect ratio for the image
    PicRatio = pic.Width / pic.Height
    e = e Or Err

    'Calculate the printable dimensions (affected by page size)
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)

    'Calculate aspect ratio for the printed page
    PrnRatio = PrnWidth / PrnHeight
    e = e Or Err

    'If the user has not selected "fit image to page," calculate the printable size
    If (Not iFit) Then
        Dim dpiRatio As Double
        dpiRatio = baseDPI / desiredDPI
        PrnPicWidth = Prn.ScaleX(pic.Width, vbHiMetric, Prn.ScaleMode) * dpiRatio
        PrnPicHeight = Prn.ScaleY(pic.Height, vbHiMetric, Prn.ScaleMode) * dpiRatio
    
    'Otherwise, set the printable size and image size to the same dimensions
    Else
        If (PicRatio >= PrnRatio) Then
            PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
        Else
            PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
        End If
    End If

    'If the user has told us to center the image, calculate offsets
    If (iCenter) Then
        OffsetX = (Prn.ScaleWidth - PrnPicWidth) \ 2
        OffsetY = (Prn.ScaleHeight - PrnPicHeight) \ 2
    End If

    'Print the picture using VB's PaintPicture method
    Prn.PaintPicture pic, OffsetX, OffsetY, PrnPicWidth, PrnPicHeight
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
    
    If FreeImageEnabled = False Then Exit Sub
    
    'If the fit-to-page option is selected (which it is by default) this routine is very simple:
    If chkFit.Value = vbChecked Then
        If cbOrientation.ListIndex = 0 Then
            iSrc.Picture = picThumb.Picture
            iSrc.Refresh
            UpdateDPI CSng(FormMain.ActiveForm.BackBuffer.ScaleWidth) / Printer.ScaleX(Printer.Width, Printer.ScaleMode, vbInches)
        Else
            iSrc.Picture = picThumbFinal.Picture
            iSrc.Refresh
            UpdateDPI CSng(FormMain.ActiveForm.BackBuffer.ScaleHeight) / Printer.ScaleX(Printer.Width, Printer.ScaleMode, vbInches)
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
    
    Dim OffsetX       As Double
    Dim OffsetY       As Double

    On Error Resume Next

    Printer.Orientation = cbOrientation.ListIndex + 1
    Printer.PrintQuality = -(cbQuality.ListIndex + 1)
    
    'Calculate the dimensions of the printable area in HiMetric
    PrnWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHiMetric)
    PrnHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHiMetric)

    Dim pic As StdPicture
    Set pic = iOut

    PrnPicWidth = Printer.ScaleX(pic.Width, vbHiMetric, Printer.ScaleMode)
    PrnPicHeight = Printer.ScaleY(pic.Height, vbHiMetric, Printer.ScaleMode)
    
    'Estimate DPI
    Dim dpiRatio As Double
    If forceDPI = False Then
        Dim dpiX As Double, dpiY As Double
        dpiX = CSng(FormMain.ActiveForm.BackBuffer.ScaleWidth) / Printer.ScaleX(PrnPicWidth, Printer.ScaleMode, vbInches)
        dpiY = CSng(FormMain.ActiveForm.BackBuffer.ScaleHeight) / Printer.ScaleY(PrnPicHeight, Printer.ScaleMode, vbInches)
        UpdateDPI ((dpiX + dpiY) / 2)
        dpiRatio = 1
    Else
        'Calculate a DPI ratio
        dpiRatio = baseDPI / desiredDPI
        PrnPicWidth = PrnPicWidth * dpiRatio
        PrnPicHeight = PrnPicHeight * dpiRatio
    End If
    
    If chkCenter.Value = vbChecked Then
        OffsetX = (Printer.ScaleWidth - PrnPicWidth) \ 2
        OffsetY = (Printer.ScaleHeight - PrnPicHeight) \ 2
    End If
    
    'Now, convert the printer-specific measurements to their corresponding measurements in the preview window
    If cbOrientation.ListIndex = 0 Then
        OffsetX = (OffsetX / Printer.ScaleWidth) * iSrc.ScaleWidth
        OffsetY = (OffsetY / Printer.ScaleHeight) * iSrc.ScaleHeight
        PrnPicWidth = (PrnPicWidth / Printer.ScaleWidth) * iSrc.ScaleWidth
        PrnPicHeight = (PrnPicHeight / Printer.ScaleHeight) * iSrc.ScaleHeight
    Else
        Dim tmpOX As Double, tmpOY As Double, tmpWidth As Single, tmpHeight As Single
        tmpOX = (OffsetY / Printer.ScaleHeight) * iSrc.ScaleWidth
        tmpOY = (OffsetX / Printer.ScaleWidth) * iSrc.ScaleHeight
        tmpWidth = (PrnPicHeight / Printer.ScaleHeight) * iSrc.ScaleWidth
        tmpHeight = (PrnPicWidth / Printer.ScaleWidth) * iSrc.ScaleHeight
        OffsetX = tmpOX
        OffsetY = tmpOY
        PrnPicWidth = tmpWidth
        PrnPicHeight = tmpHeight
    End If
    
    
    'Draw a new preview
    If cbOrientation.ListIndex = 0 Then
        DrawPreviewImage picThumb
        iSrc.Picture = LoadPicture("")
        SetStretchBltMode iSrc.hdc, STRETCHBLT_HALFTONE
        StretchBlt iSrc.hdc, OffsetX, OffsetY, PrnPicWidth, PrnPicHeight, picThumb.hdc, PreviewX, PreviewY, PreviewWidth, PreviewHeight, vbSrcCopy
    Else
        DrawPreviewImage picThumb90
        iSrc.Picture = LoadPicture("")
        SetStretchBltMode iSrc.hdc, STRETCHBLT_HALFTONE
        StretchBlt iSrc.hdc, OffsetX, OffsetY, PrnPicWidth, PrnPicHeight, picThumbFinal.hdc, PreviewY, PreviewX, PreviewHeight, PreviewWidth, vbSrcCopy
    End If
    
    iSrc.Picture = iSrc.Image
    iSrc.Refresh
      
End Sub

'This is called whenever the dimensions of the preview window change (for example, in response to a change in paper size)
Private Sub RebuildPreview(Optional forceDPI As Boolean = False)
    
    'FreeImage is used to rotate the image; if it's not installed, previewing is automatically disabled
    If FreeImageEnabled = True Then
    
        'We're now going to create two temporary buffers; one contains the image resized to fit the "sheet of paper" preview
        ' on the left.  This is portrait mode.  The second buffer will contain the same thing, but rotated 90 degrees -
        ' e.g. landscape mode.  If the user clicks between those options, we can simply copy the buffers to the foreground
        ' picture box.
        
        'First is the easy one - Portrait Mode
        picThumb.Picture = LoadPicture("")
        picThumb.Width = iSrc.Width
        picThumb.Height = iSrc.Height
        DrawPreviewImage picThumb
        
        'Now we need to get the source image at the size expected post-rotation
        picThumb90.Picture = LoadPicture("")
        picThumbFinal.Picture = LoadPicture("")
        picThumb90.Width = iSrc.Height
        picThumb90.Height = iSrc.Width
        picThumbFinal.Width = iSrc.Width
        picThumbFinal.Height = iSrc.Height

        DrawPreviewImage picThumb90
        GetPreviewData picThumb90
        
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
    pWidth = Printer.ScaleX(Printer.Width, Printer.ScaleMode, vbInches)
    pHeight = Printer.ScaleY(Printer.Height, Printer.ScaleMode, vbInches)
    Dim txtWidth As String, txtHeight As String
    txtWidth = Format(pWidth, "#0.##")
    txtHeight = Format(pHeight, "#0.##")
    If Right(txtWidth, 1) = "." Then txtWidth = Left$(txtWidth, Len(txtWidth) - 1)
    If Right(txtHeight, 1) = "." Then txtHeight = Left$(txtHeight, Len(txtHeight) - 1)
    
    lblPaperSize = "Paper size: " & txtWidth & """ x  " & txtHeight & """"
    
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
        iSrc.Top = 16 + (220 - iSrc.Height) / 2
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
Private Sub UpdateDPI(ByVal eDPI As Single)
    cmbDPI = Int(eDPI + 0.5)
End Sub

Private Sub cmbDPI_Change()
    If EntryValid(cmbDPI, 1, 12000, False, False) Then
        desiredDPI = cmbDPI
        UpdatePrintPreview True
    End If
End Sub
