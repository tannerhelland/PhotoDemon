VERSION 5.00
Begin VB.Form FormPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print"
   ClientHeight    =   5745
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
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldCopies 
      Height          =   735
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
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
      TabIndex        =   8
      Top             =   5010
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdDropDown cbOrientation 
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "orientation"
   End
   Begin PhotoDemon.pdDropDown cbQuality 
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "quality"
   End
   Begin PhotoDemon.pdDropDown cbPrinters 
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   4575
      _ExtentX        =   7858
      _ExtentY        =   1296
      Caption         =   "printer"
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
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3300
   End
   Begin PhotoDemon.pdLabel lblPaperSize 
      Height          =   375
      Left            =   240
      Top             =   4440
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
'Copyright 2000-2021 by Tanner Helland
'Created: 4/April/03
'Last updated: 22/September/17
'Last update: emergency fixes to restore printing support on XP.  Some features were stripped to enable compatibility
'             with problematic recent Windows patches.  Because this is such a fringe case, better fixes are not planned
'             until post-7.0's release.
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
'Module for interfacing with the printer.  All settings are handled through the default VB printer object.
' A key feature of this routine (PrintPictureToFitPage) is based off code first written by Waty Thierry
' (http://www.ppreview.net/).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'These usually exist in a common dialog object, but since we're manually handling all
' printer functionality we need to manually specify them
Private Enum PrinterOrientationConstants
    cdlLandscape = 2
    cdlPortrait = 1
End Enum

'Changing the orientation box forces a refresh of the preview
Private Sub cbOrientation_Click()
    UpdatePrintPreview
End Sub

'Allow the user to change the target printer
Private Sub cbPrinters_Click()

    On Error GoTo PrinterNameProblem

    Dim Prt As Printer
    
    For Each Prt In Printers
        If Strings.StringsEqual(Prt.DeviceName, cbPrinters.List(cbPrinters.ListIndex), True) Then
            Set Printer = Prt
            Exit For
        End If
        
PrinterNameProblem:
    Next Prt
    
    UpdatePaperSize
    UpdatePrintPreview

End Sub

Private Sub cmdBar_CancelClick()
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    On Error GoTo PrintingFailed
    
    Message "Sending image to printer..."
    
    'Set the number of copies
    Printer.Copies = sldCopies.Value
      
    'Assuming there have been no errors thus far (basically, assuming the user actually has a printer attatched)...
    If (Err = 0) Then
      
        'Set print quality
        Printer.PrintQuality = -(cbQuality.ListIndex + 1)
          
        'Assuming our quality option is valid (should be, but you never know)
        If (Err = 0) Then
          
            'Print the image
            If (Not PrintPictureToFitPage(Printer, picOut.Picture, cbOrientation.ListIndex + 1)) Then GoTo PrintingFailed
            
        End If
    End If
    
    Printer.EndDoc
    
    Message "Image printed successfully. (Note: depending on your printer, additional print confirmation screens may appear.)"
    
    Exit Sub
    
PrintingFailed:

    PDMsgBox "%1 was unable to print the image.  Please make sure that the specified printer (%2) is powered-on and ready for printing.", vbExclamation Or vbOKOnly, "Error", "PhotoDemon", Printer.DeviceName
    cmdBar.DoNotUnloadForm
    
End Sub

'LOAD form
Private Sub Form_Load()

    On Error GoTo PrinterLoadError
    
    RebuildPreview
    
    Dim x As Long
    
    'Load a list of printers into the combo box
    For x = 0 To Printers.Count - 1
        cbPrinters.AddItem Printers(x).DeviceName
    Next x

    'Pre-select the current printer
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

    'Set image orientation based on its aspect ratio (as compared to an 8.5" x 11" sheet of paper)
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

    'Temporarily copy the image into an image box
    picOut.Width = PDImages.GetActiveImage.Width
    picOut.Height = PDImages.GetActiveImage.Height
    picOut.ScaleMode = vbPixels
    
    Dim tmpComposite As pdDIB
    Set tmpComposite = New pdDIB
    PDImages.GetActiveImage.GetCompositedImage tmpComposite
    tmpComposite.RenderToPictureBox picOut, , , True
    
    picOut.ScaleMode = vbTwips
    
PrinterLoadError:
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

'This PrintPictureToFitPage function is based off code originally written by Waty Thierry.
' (It has been heavily modified for use within PhotoDemon, but you may download the original at http://www.freevbcode.com/ShowCode.asp?ID=194)
Private Function PrintPictureToFitPage(Prn As Printer, pic As StdPicture, ByVal iOrientation As PrinterOrientationConstants) As Boolean

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
    
    'Otherwise, set the printable size and image size to the same dimensions
    If (PicRatio >= PrnRatio) Then
        PrnPicWidth = Prn.scaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.scaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
        PrnPicHeight = Prn.scaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.scaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    
    'If the user has told us to center the image, calculate offsets
    offsetX = (Prn.ScaleWidth - PrnPicWidth) \ 2
    offsetY = (Prn.ScaleHeight - PrnPicHeight) \ 2
    
    'Print the picture using VB's PaintPicture method
    Prn.PaintPicture pic, offsetX, offsetY, PrnPicWidth, PrnPicHeight
    e = e Or Err

    PrintPictureToFitPage = (e = 0)
    On Error GoTo 0

End Function

'Redraw the Print Preview box based on the specified image orientation
Private Sub UpdatePrintPreview()
    
    On Error GoTo PrintPreviewError
    
    'If the fit-to-page option is selected (which it is by default) this routine is very simple:
    If (cbOrientation.ListIndex = 0) Then
        iSrc.Picture = picThumb.Picture
        iSrc.Refresh
    Else
        iSrc.Picture = picThumbFinal.Picture
        iSrc.Refresh
    End If
    
PrintPreviewError:
      
End Sub

'This is called whenever the dimensions of the preview window change (for example, in response to a change in paper size)
Private Sub RebuildPreview()

    'We're now going to create two temporary buffers; one contains the image resized to fit the "sheet of paper" preview
    ' on the left.  This is portrait mode.  The second buffer will contain the same thing, but rotated 90 degrees -
    ' e.g. landscape mode.  If the user clicks between those options, we can simply copy the buffers to the foreground
    ' picture box.
    
    'First is the easy one - Portrait Mode
    picThumb.Picture = LoadPicture(vbNullString)
    picThumb.Width = iSrc.Width
    picThumb.Height = iSrc.Height
    DrawPreviewImage picThumb, True
    
    'Now we need to get the source image at the size expected post-rotation
    picThumb90.Picture = LoadPicture(vbNullString)
    picThumbFinal.Picture = LoadPicture(vbNullString)
    picThumb90.Width = iSrc.Height
    picThumb90.Height = iSrc.Width
    picThumbFinal.Width = iSrc.Width
    picThumbFinal.Height = iSrc.Height

    DrawPreviewImage picThumb90, True
    
    'Now comes the rotation itself.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromDC picThumb90.hDC, 0, 0, picThumb90.ScaleWidth, picThumb90.ScaleHeight, 32, True
    
    Dim dstDIB As pdDIB
    Set dstDIB = New pdDIB
    GDI_Plus.GDIPlusRotateFlipDIB tmpDIB, dstDIB, GP_RF_90FlipNone
    dstDIB.RenderToPictureBox picThumbFinal, False, False, True
    
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
    If (aspectRatio >= 1) Then
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
    RebuildPreview
    
PaperSizeError:

    If (Not scaleActionWorked) Then lblPaperSize.Caption = vbNullString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Used to draw the main image onto a preview picture box
Private Sub DrawPreviewImage(ByRef dstPicture As PictureBox, Optional forceWhiteBackground As Boolean = False)
    
    Dim tmpDIB As pdDIB
    
    'Start by calculating the aspect ratio of both the current image and the previewing picture box
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = dstPicture.ScaleWidth
    dstHeight = dstPicture.ScaleHeight
    
    Dim srcWidth As Double, srcHeight As Double
    
    'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
    Dim srcDIB As pdDIB
    PDImages.GetActiveImage.GetCompositedImage srcDIB
    srcWidth = srcDIB.GetDIBWidth
    srcHeight = srcDIB.GetDIBHeight
            
    'Now, use that aspect ratio to determine a proper size for our temporary DIB
    Dim newWidth As Long, newHeight As Long
    ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
    
    'Normally this will draw a preview of PDImages.GetActiveImage.containingForm's relevant image.  However, another picture source can be specified.
    If srcDIB.GetDIBColorDepth = 32 Then
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB srcDIB, newWidth, newHeight
        If forceWhiteBackground Then tmpDIB.CompositeBackgroundColor 255, 255, 255
        tmpDIB.RenderToPictureBox dstPicture
    Else
        srcDIB.RenderToPictureBox dstPicture
    End If
    
End Sub
