VERSION 5.00
Begin VB.Form FormPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8745
   ControlBox      =   0   'False
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
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldCopies 
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "copies"
      Min             =   1
      Max             =   100
      ScaleStyle      =   1
      Value           =   1
      NotchPosition   =   1
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   4545
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdDropDown cbOrientation 
      Height          =   735
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "orientation"
   End
   Begin PhotoDemon.pdDropDown cbQuality 
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "quality"
   End
   Begin PhotoDemon.pdDropDown cbPrinters 
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   4575
      _ExtentX        =   7858
      _ExtentY        =   1296
      Caption         =   "printer"
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   3300
      Left            =   240
      Top             =   360
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5821
   End
   Begin PhotoDemon.pdLabel lblPaperSize 
      Height          =   375
      Left            =   240
      Top             =   3840
      Width           =   3300
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "paper size"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Printer Interface (including Print Preview)
'Copyright 2000-2026 by Tanner Helland
'Created: 4/April/03
'Last updated: 13/April/22
'Last update: restructure this horrifying dialog into something slightly less horrifying... and remove a bunch of
'             VB picture boxes while I'm at it.
'
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
'Very, very simplified interface for printing the active image.  This dialog is primarily reserved for
' Windows XP systems because on Vista+ we can interface with the built-in Windows Print Wizard.
'
'This code is not a high point of PhotoDemon's design and I'm okay with that.  This pretty much exists just
' to cover a bare-minimum usage case.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Persistently cached previews in upright and rotated modes.  This is convenient because we can leave the
' on-screen "paper sized" preview the same dimensions, and simply switch between these two cached thumbnails.
Private m_previewPage As pdDIB, m_previewPage90 As pdDIB

'The final preview is rendered here, then "flipped" to the screen as necessary
Private m_finalPreview As pdDIB

'A composited copy of the current image is cached here
Private m_CompositeImage As pdDIB

'Changing the orientation box forces a refresh of the preview
Private Sub cbOrientation_Click()
    UpdatePrintPreview
End Sub

'Allow the user to change the target printer.  We do no validation at this point - errors will have to wait
' until the user actually attempts to print something.
Private Sub cbPrinters_Click()
    
    On Error GoTo PrinterNameProblem
    
    Dim srcPrinter As Printer
    
    For Each srcPrinter In Printers
        If Strings.StringsEqual(srcPrinter.DeviceName, cbPrinters.List(cbPrinters.ListIndex), True) Then
            Set Printer = srcPrinter
            Exit For
        End If
        
PrinterNameProblem:
    Next srcPrinter
    
    UpdatePaperSize
    UpdatePrintPreview
    
End Sub

Private Sub cmdBar_CancelClick()
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    
    On Error GoTo PrintingFailed
    
    Message "Sending image to printer..."
    
    'Set all printer properties before attempting to print anything.  It's entirely possible
    ' that this step will error out, which usually means that the selected printer is unreachable -
    ' if we encounter errors like that, we'll abandon printing entirely.  (Solving such errors is
    ' outside the purview of PhotoDemon!)
    Printer.Copies = sldCopies.Value
    Printer.PrintQuality = -(cbQuality.ListIndex + 1)
    Printer.Orientation = cbOrientation.ListIndex + 1
    
    'Attempt to print the image via standard GDI calls
    If (Not PrintPictureToFitPage(Printer)) Then GoTo PrintingFailed
    
    'Finalize the print and hope for the best
    Printer.EndDoc
    Message "Image printed successfully. (Note: depending on your printer, additional print confirmation screens may appear.)"
    
    Exit Sub
    
PrintingFailed:
    
    PDMsgBox "%1 was unable to print the image.  Please make sure that the specified printer (%2) is powered-on and ready for printing.", vbExclamation Or vbOKOnly, "Error", "PhotoDemon", Printer.DeviceName
    
    'Give the user a chance to try again
    cmdBar.DoNotUnloadForm
    
End Sub

Private Sub Form_Load()
    
    'You'll see more error-handling than usual in this dialog.  A lot can go wrong while printing,
    ' and we just have to roll with it as best we can
    On Error GoTo PrinterLoadError
    
    'Retrieve a copy of the image-to-print
    Dim tmpDIB As pdDIB
    PDImages.GetActiveImage.GetCompositedImage tmpDIB, True
    
    Set m_CompositeImage = New pdDIB
    m_CompositeImage.CreateBlank tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 32, vbWhite, 255
    m_CompositeImage.SetInitialAlphaPremultiplicationState True
    tmpDIB.AlphaBlendToDC m_CompositeImage.GetDIBDC
    Set tmpDIB = Nothing
    
    'Generate on-screen thumbnails for both portrait and landscape modes
    RebuildPreview
    
    'Load a list of printers to the UI
    Dim i As Long
    For i = 0 To Printers.Count - 1
        cbPrinters.AddItem Printers(i).DeviceName
    Next i
    
    'Pre-select the current system printer
    Dim printerIndex As Long
    printerIndex = cbPrinters.ListIndexByString(Printer.DeviceName, vbTextCompare)
    If (printerIndex < 0) Then printerIndex = 0
    cbPrinters.ListIndex = printerIndex
    UpdatePaperSize
    
    'Set a default quality (high, since we're working with images)
    cbQuality.AddItem "draft", 0
    cbQuality.AddItem "low", 1
    cbQuality.AddItem "medium", 2
    cbQuality.AddItem "high", 3
    cbQuality.ListIndex = 3
    
    'Set default image orientation based on its aspect ratio (as compared to an 8.5" x 11" sheet of paper)
    cbOrientation.AddItem "portrait", 0
    cbOrientation.AddItem "landscape", 1
    cbOrientation.ListIndex = 0
    
    Dim imgAspect As Double, paperAspect As Double
    imgAspect = PDImages.GetActiveImage.Width / PDImages.GetActiveImage.Height
    paperAspect = 8.5 / 11#
    
    If (imgAspect < paperAspect) Then
        cbOrientation.ListIndex = 0
    Else
        cbOrientation.ListIndex = 1
    End If
    
    UpdatePrintPreview
    
    'If something went wrong during initialization, still display the dialog because one or more printers
    ' may have still been initialized successfully
PrinterLoadError:
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

'Once upon a time, this PrintPictureToFitPage function was derived from code originally written by Waty Thierry.
' It no longer resembles that original version whatsoever (and now relies upon various PhotoDemon-specific functions),
' but you can still find Waty's original, general-purpose version here:
' https://web.archive.org/web/20171112231338/http://www.freevbcode.com/ShowCode.asp?ID=194
Private Function PrintPictureToFitPage(ByRef dstPrinter As Printer) As Boolean
    
    'Printing is potentially rife with errors.  I have no plans for dealing with such errors,
    ' so we're just gonna bail if anything goes wrong.
    PrintPictureToFitPage = False
    On Error GoTo PrintFailed
    
    'To coerce VB into printing (which triggers the StartDoc API equivalent), we need to send
    ' *something* to the printer - so cheat and print a " ".  This will e.g. raise a Save File dialog
    ' for the "Microsoft Print to PDF" printer.
    Printer.Print " "
    
    'If we're still here, printing will likely succeed.
    
    'Calculate an aspect ratio for the source image
    Dim picRatio As Double
    picRatio = m_CompositeImage.GetDIBWidth / m_CompositeImage.GetDIBHeight
    
    'Retrieve print page dimensions and convert to pixels
    Dim prnWidthPixels As Long, prnHeightPixels As Long
    prnWidthPixels = dstPrinter.scaleX(dstPrinter.ScaleWidth, dstPrinter.ScaleMode, vbPixels)
    prnHeightPixels = dstPrinter.scaleY(dstPrinter.ScaleHeight, dstPrinter.ScaleMode, vbPixels)
    If (prnWidthPixels <= 0) Or (prnHeightPixels <= 0) Then Exit Function
    
    'Calculate aspect ratio for the printed page
    Dim printerRatio As Double
    printerRatio = prnWidthPixels / prnHeightPixels
    
    'Fit the aspect ratios to each other
    Dim dstWidth As Long, dstHeight As Long
    If (picRatio >= printerRatio) Then
        dstWidth = prnWidthPixels
        dstHeight = dstWidth / picRatio
    Else
        dstHeight = prnHeightPixels
        dstWidth = dstHeight * picRatio
    End If
    
    'Calculate offsets between the fitted size and the page size
    Dim offsetX As Long, offsetY As Long
    offsetX = (prnWidthPixels - dstWidth) \ 2
    offsetY = (prnHeightPixels - dstHeight) \ 2
    
    'StretchBlt that sucker into place!
    GDI.StretchBltWrapper dstPrinter.hDC, offsetX, offsetY, dstWidth, dstHeight, m_CompositeImage.GetDIBDC, 0, 0, m_CompositeImage.GetDIBWidth, m_CompositeImage.GetDIBHeight, vbSrcCopy
    
    PrintPictureToFitPage = True
    Exit Function
    
PrintFailed:
    PrintPictureToFitPage = False
    
End Function

'Redraw the Print Preview box based on the specified image orientation
Private Sub UpdatePrintPreview()
    
    'If the fit-to-page option is selected (which it is by default) this routine is very simple:
    If (cbOrientation.ListIndex = 0) Then
        Set m_finalPreview = m_previewPage
    Else
        Set m_finalPreview = m_previewPage90
    End If
    
    picPreview.RequestRedraw True
    
End Sub

'This is called whenever the dimensions of the preview window change
' (for example, in response to a change in paper size)
Private Sub RebuildPreview()
    
    If (m_CompositeImage Is Nothing) Then Exit Sub
    
    'We're now going to create two temporary buffers; one contains the image resized to fit the
    ' "sheet of paper" preview on the left.  The second buffer will contain the same thing,
    ' but rotated 90 degrees - e.g. landscape mode.  If the user clicks between those options,
    ' we can simply repaint these prepared buffers to the foreground picture box.
    Set m_previewPage = New pdDIB
    m_previewPage.CreateBlank picPreview.GetWidth, picPreview.GetHeight, 32, vbWhite, 255
    m_previewPage.SetInitialAlphaPremultiplicationState True
    
    'Calculate portrait mode dimensions, then paint the image into place
    Dim newWidth As Long, newHeight As Long
    PDMath.ConvertAspectRatio m_CompositeImage.GetDIBWidth, m_CompositeImage.GetDIBHeight, picPreview.GetWidth - 2, picPreview.GetHeight - 2, newWidth, newHeight, True
    
    Dim xOffset As Long, yOffset As Long
    xOffset = (m_previewPage.GetDIBWidth - newWidth) \ 2
    yOffset = (m_previewPage.GetDIBHeight - newHeight) \ 2
    GDI_Plus.GDIPlus_StretchBlt m_previewPage, xOffset, yOffset, newWidth, newHeight, m_CompositeImage, 0, 0, m_CompositeImage.GetDIBWidth, m_CompositeImage.GetDIBHeight
    
    'We'll need a temporary placeholder for the rotated image, at the rotated dimensions
    Dim tmpImage As pdDIB
    Set tmpImage = New pdDIB
    tmpImage.CreateBlank picPreview.GetHeight, picPreview.GetWidth, 32, vbWhite, 255
    tmpImage.SetInitialAlphaPremultiplicationState True
    
    'Repeat the previous steps, then rotate the final result into the module-level DIB
    PDMath.ConvertAspectRatio m_CompositeImage.GetDIBWidth, m_CompositeImage.GetDIBHeight, tmpImage.GetDIBWidth - 2, tmpImage.GetDIBHeight - 2, newWidth, newHeight, True
    xOffset = (tmpImage.GetDIBWidth - newWidth) \ 2
    yOffset = (tmpImage.GetDIBHeight - newHeight) \ 2
    GDI_Plus.GDIPlus_StretchBlt tmpImage, xOffset, yOffset, newWidth, newHeight, m_CompositeImage, 0, 0, m_CompositeImage.GetDIBWidth, m_CompositeImage.GetDIBHeight
    GDI_Plus.GDIPlusRotateFlipDIB tmpImage, m_previewPage90, GP_RF_90FlipNone
    m_previewPage90.SetInitialAlphaPremultiplicationState True
    
    'We now want to paint a neutral border around both preview images
    Dim cPenBorder As pd2DPen
    Set cPenBorder = New pd2DPen
    If (Not g_Themer Is Nothing) Then cPenBorder.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
    cPenBorder.SetPenWidth 1!
    cPenBorder.SetPenLineJoin P2_LJ_Miter
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_previewPage, False
    PD2D.DrawRectangleI cSurface, cPenBorder, 0, 0, m_previewPage.GetDIBWidth - 1, m_previewPage.GetDIBHeight - 1
    Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_previewPage90, False
    PD2D.DrawRectangleI cSurface, cPenBorder, 0, 0, m_previewPage90.GetDIBWidth - 1, m_previewPage90.GetDIBHeight - 1
    
    'Initiate a redraw of the preview according to the print settings currently specified by the user
    UpdatePrintPreview
        
End Sub

'Display what size of paper we expect to output the image onto
Private Sub UpdatePaperSize()
    
    On Error GoTo PaperSizeError
    
    Dim scaleActionWorked As Boolean
    
    'Note: interacting with the printer is a messy, complex chunk of API hackery (see: http://support.microsoft.com/kb/282474/t)
    ' To that end, I've written this code to work with any paper size - all it requires is that the Printer.Width and Printer.Height
    ' values be set to whatever size the user desires.  Thus this code is abstracted away from the actual paper size selection
    ' process, and any new selection method will require zero changes here.
    Dim pWidth As Double, pHeight As Double
    pWidth = Me.scaleX(Printer.Width, Printer.ScaleMode, vbInches)
    pHeight = Me.scaleY(Printer.Height, Printer.ScaleMode, vbInches)
    scaleActionWorked = True
    
    Dim sWidth As String, sHeight As String
    sWidth = Format$(pWidth, "0.0#")
    sHeight = Format$(pHeight, "0.0#")
    lblPaperSize.Caption = g_Language.TranslateMessage("paper size") & ": " & sWidth & """ x  " & sHeight & """"
    
    'Now comes the tricky part - we need to resize the preview box to match the aspect ratio of the paper
    Dim aspectRatio As Double
    If (pHeight <> 0#) Then aspectRatio = pWidth / pHeight Else aspectRatio = 1#
    
    'The maximum available area for the preview is 220 pixels.  The margin on both top and left is 16 pixels.
    ' This means that the full workable area is 220 + 16 * 2, or 252 in either direction.
    ' We will rebuild the preview picture box according to the new aspect ratio and these size constraints.
    
    'If width is larger than height, assign width a size of 220 pixels and adjust height accordingly
    If (aspectRatio >= 1#) Then
        picPreview.SetWidth 220
        picPreview.SetLeft 16
        picPreview.SetHeight 220 * aspectRatio
        picPreview.SetTop 16 + (lblPaperSize.GetHeight - picPreview.GetHeight) \ 2
    Else
        picPreview.SetHeight 220
        picPreview.SetTop lblPaperSize.GetTop - 16 - 220
        picPreview.SetWidth 220 * aspectRatio
        picPreview.SetLeft 16 + (220 - picPreview.GetWidth) \ 2
    End If
    
    'Automatically initiate a rebuild of the preview picture boxes
    RebuildPreview
    
PaperSizeError:

    If (Not scaleActionWorked) Then lblPaperSize.Caption = vbNullString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_finalPreview Is Nothing) Then m_finalPreview.AlphaBlendToDC targetDC
End Sub
