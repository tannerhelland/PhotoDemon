VERSION 5.00
Begin VB.Form dialog_ImportPDF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13500
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
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   900
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldPreview 
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1508
      Caption         =   "preview"
   End
   Begin PhotoDemon.pdButtonStrip btsPanel 
      Height          =   1095
      Left            =   4800
      TabIndex        =   4
      Top             =   1320
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1931
      Caption         =   "import settings"
      FontSize        =   12
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   5535
      Left            =   240
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9763
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6735
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   1
      Left            =   4800
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Caption         =   "file information"
      FontSize        =   12
   End
   Begin PhotoDemon.pdLabel lblOriginal 
      Height          =   345
      Index           =   0
      Left            =   4920
      Top             =   510
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   609
      Caption         =   ""
   End
   Begin PhotoDemon.pdLabel lblOriginal 
      Height          =   375
      Index           =   1
      Left            =   4920
      Top             =   840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      Caption         =   ""
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   4095
      Index           =   2
      Left            =   4800
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      Begin PhotoDemon.pdColorSelector clsBackground 
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   3360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1085
      End
      Begin PhotoDemon.pdButtonStrip btsTransparency 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1720
         Caption         =   "background"
      End
      Begin PhotoDemon.pdButtonStrip btsAntialiasing 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1720
         Caption         =   "antialiasing"
      End
      Begin PhotoDemon.pdButtonStrip btsAnnotations 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1720
         Caption         =   "annotations"
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   4095
      Index           =   0
      Left            =   4800
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      Begin PhotoDemon.pdResize rszUI 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         DefaultToRealWorldUnits=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   2
         Left            =   120
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
         Caption         =   "document size"
         FontSize        =   12
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   4095
      Index           =   1
      Left            =   4800
      Top             =   2520
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      Begin PhotoDemon.pdCheckBox chkPageOptions 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         Caption         =   "reverse page order"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdTextBox txtPageRange 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButtonStrip btsPages 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1720
         Caption         =   "pages to import"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   3
         Left            =   120
         Top             =   1680
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   661
         Caption         =   "other options"
         FontSize        =   12
      End
   End
End
Attribute VB_Name = "dialog_ImportPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PDF Import Dialog
'Copyright 2024-2026 by Tanner Helland
'Created: 29/February/24
'Last updated: 13/July/24
'Last update: fix crash on setting focus to invalid entries in the "page range" edit box
'             (needed if the user enteres an invalid page range then switches to a different panel)
'
'Like Photoshop and GIMP (and probably others), PhotoDemon allows users to set their own PDF resolution,
' image size, pages-to-import, and various other PDF-specific features at import-time.  This dialog
' provides a UI for any custom settings.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Reference to the source PDF object handling the import; it handles rendering previews, retrieving original
' PDF properties, etc.
Private m_PDF As pdPDF

'Base image size IN POINTS.  (We must convert this to pixels based on the user's desired import resolution.)
Private m_baseImageWidthInPts As Single, m_baseImageHeightInPts As Single

'Base image size IN PIXELS, ASSUMING 96 DPI.  (Calculated at run-time, and may be modified by the user -
' but we need a baseline value in case the dialog's "reset to defaults" button is pressed.)
Private Const DEFAULT_DPI As Long = 96
Private m_baseImageWidthInPx As Long, m_baseImageHeightInPx As Long

'This dialog automatically redraws the preview window as necessary.  To suspend this behavior
' (while prepping the dialog for first show, for example) set this to TRUE.
Private m_SuspendRedraws As Boolean

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetDialogParamString() As String
    
    'Build a param string out of the various PDF settings
    Dim finalWidthPx As Single, finalHeightPx As Single, finalDPI As Single
    finalWidthPx = rszUI.ResizeWidthInPixels()
    finalHeightPx = rszUI.ResizeHeightInPixels()
    finalDPI = rszUI.ResizeDPIAsPPI
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Dimensions
    cParams.AddParam "final-width-px", finalWidthPx, True
    cParams.AddParam "final-height-px", finalHeightPx, True
    cParams.AddParam "final-dpi", finalDPI, True
    
    'Pages to import
    If (btsPages.ListIndex = 0) Then
        cParams.AddParam "import-pages", "all", True, True
    ElseIf (btsPages.ListIndex = 1) Then
        cParams.AddParam "import-pages", "custom", True, True
        cParams.AddParam "page-list", Trim$(Str$(sldPreview.Value)), True
    Else
    
        cParams.AddParam "import-pages", "custom", True, True
        
        'Ensure the page range can be parsed and that at least one valid page is listed
        Dim listOfPages As pdStack
        Set listOfPages = New pdStack
        If TextSupport.ConvertPageRangeToStack(txtPageRange.Text, listOfPages) Then
            
            Dim listOfPagesAsText As pdString
            Set listOfPagesAsText = New pdString
            
            Dim i As Long, idxLast As Long
            idxLast = listOfPages.GetNumOfInts - 1
            For i = 0 To listOfPages.GetNumOfInts - 1
                listOfPagesAsText.Append listOfPages.GetInt(i)
                If (i < idxLast) Then listOfPagesAsText.Append ","
            Next i
            
            cParams.AddParam "page-list", listOfPagesAsText.ToString(), True
            
        End If
        
    End If
    
    'Other page settings
    cParams.AddParam "reverse-pages", chkPageOptions(0).Value, True
    
    'Rendering settings
    cParams.AddParam "background-solid", (btsTransparency.ListIndex = 0)
    cParams.AddParam "background-color", clsBackground.Color
    cParams.AddParam "antialiasing", btsAntialiasing.ListIndex
    cParams.AddParam "annotations", (btsAnnotations.ListIndex = 0)
    
    GetDialogParamString = cParams.GetParamString()
    
End Function

Private Sub btsAnnotations_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsAntialiasing_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsPages_Click(ByVal buttonIndex As Long)
    UpdatePagesUI
End Sub

'When the user selects "custom page range", provide a text box where they can specify the desired range
Private Sub UpdatePagesUI()
    txtPageRange.Visible = (btsPages.ListIndex = 2)
End Sub

Private Sub UpdateTransparencyUI()
    clsBackground.Visible = (Me.btsTransparency.ListIndex = 0)
End Sub

Private Sub btsPanel_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub btsTransparency_Click(ByVal buttonIndex As Long)
    UpdateTransparencyUI
    UpdatePreview
End Sub

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub cmdBar_AddCustomPresetData()
    
    'Save the resize bar's DPI as a custom value; we want to restore this (and ONLY this) on future loads
    cmdBar.AddPresetData "doc-dpi", rszUI.ResizeDPIAsPPI
    
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_ExtraValidations()
    
    'If the user wants to specify a custom page range, ensure their text passes validation
    If (btsPages.ListIndex = 2) Then
        
        Dim listOfPages As pdStack
        If (Not TextSupport.ConvertPageRangeToStack(txtPageRange.Text, listOfPages)) Then
            PDMsgBox "The range ""%1"" is not valid.", vbExclamation Or vbOKOnly, "Error", txtPageRange.Text
            cmdBar.ValidationFailed
            txtPageRange.SetFocus
            Exit Sub
        
        'The range is valid *textually speaking*, but it may include pages outside this PDF's range.
        ' Check that next.
        Else
            
            If (Not m_PDF Is Nothing) Then
                
                'Sort the list of pages from least-to-most
                listOfPages.SortStackByValue True
                
                Dim idxFirst As Long, idxLast As Long
                idxFirst = listOfPages.GetInt(0)
                idxLast = listOfPages.GetInt(listOfPages.GetNumOfInts - 1)
                
                'The warning text is copied from PD's standard range warning in the TextSupport module;
                ' it's not *ideal*, but it cuts down on unique text for my volunteer localizers.
                If (Not TextSupport.EntryValid(idxFirst, 1, m_PDF.GetPageCount(), False, False)) Then
                    PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation Or vbOKOnly, "Invalid entry", idxFirst, 1, m_PDF.GetPageCount()
                    cmdBar.ValidationFailed
                    
                    'Ensure the page range text box is visible
                    If (Me.btsPanel.ListIndex <> 1) Then Me.btsPanel.ListIndex = 1
                    If (Me.btsPages.ListIndex <> 2) Then Me.btsPages.ListIndex = 2
                    
                    txtPageRange.SetFocus
                    Exit Sub
                    
                ElseIf (Not TextSupport.EntryValid(idxLast, 1, m_PDF.GetPageCount(), False, False)) Then
                    PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation Or vbOKOnly, "Invalid entry", idxLast, 1, m_PDF.GetPageCount()
                    cmdBar.ValidationFailed
                    
                    'Ensure the page range text box is visible
                    If (Me.btsPanel.ListIndex <> 1) Then Me.btsPanel.ListIndex = 1
                    If (Me.btsPages.ListIndex <> 2) Then Me.btsPages.ListIndex = 2
                    
                    txtPageRange.SetFocus
                    Exit Sub
                    
                End If
                
            End If
            
        End If
    
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_ReadCustomPresetData()
    
    'Retrieve the user's previous DPI (if any) - we want to restore this value but *not* the last-used dimensions
    Dim prevDPI As Single, strPrevDPI As String
    strPrevDPI = cmdBar.RetrievePresetData("doc-dpi")
    If (LenB(strPrevDPI) = 0) Then prevDPI = DEFAULT_DPI Else prevDPI = CDbl(strPrevDPI)
    
    'Use the retrieved DPI to calculate a new "default" image size
    Dim baseImageWidthInPx As Long, baseImageHeightInPx As Long
    baseImageWidthInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageWidthInPts, prevDPI))
    baseImageHeightInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageHeightInPts, prevDPI))
    rszUI.SetInitialDimensions baseImageWidthInPx, baseImageHeightInPx, prevDPI
    rszUI.AspectRatioLock = True
    
    'Do *not* remember the user's last page setting
    sldPreview.Value = 1
    
    'Do *not* remember the user's last custom page range
    If (Not m_PDF Is Nothing) Then
        If (m_PDF.GetPageCount > 1) Then
            txtPageRange.Text = "1-" & m_PDF.GetPageCount
        Else
            txtPageRange.Text = "1"
        End If
    End If
    
End Sub

Private Sub cmdBar_ResetClick()
    
    sldPreview.Value = 1
    rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
    rszUI.AspectRatioLock = True
    chkPageOptions(0).Value = False
    clsBackground.Color = RGB(255, 255, 255)
    btsAntialiasing.ListIndex = 0
    btsAnnotations.ListIndex = 1
    
    'Attempt to set default page listing for "custom pages" to "all pages"
    If (Not m_PDF Is Nothing) Then
        If (m_PDF.GetPageCount > 1) Then
            txtPageRange.Text = "1-" & m_PDF.GetPageCount
        Else
            txtPageRange.Text = "1"
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    'TODO: investigate why these two lines were used on other dialogs?
    'rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
    'rszUI.AspectRatioLock = True
End Sub

Private Sub Form_Load()
    
    'Do not allow previews until the form is fully loaded
    m_SuspendRedraws = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByRef srcPDF As pdPDF)
    
    'Do not provide any page previews until the load process is complete
    m_SuspendRedraws = True
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure a proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify import options... "
    
    'All PDF info comes from the source PDF object, which has already loaded the target PDF
    Set m_PDF = srcPDF
    
    'Populate the preview dropdown with all available pages
    cmdBar.RequestPresetNoLoad sldPreview
    If (Not m_PDF Is Nothing) Then
        If (m_PDF.GetPageCount > 0) Then
            sldPreview.Min = 1
            sldPreview.Max = m_PDF.GetPageCount
            sldPreview.Value = 1
        Else
            sldPreview.Min = 0
            sldPreview.Value = 0
            sldPreview.Max = 0
        End If
    End If
    
    'This dialog has too many import options, so it's split into separate panels
    btsPanel.AddItem "size", 0
    btsPanel.AddItem "pages", 1
    btsPanel.AddItem "rendering", 2
    btsPanel.ListIndex = 0
    UpdatePanelVisibility
    
    'Allow the user to import first page / all pages / custom pages
    Dim textAllPages As String
    textAllPages = g_Language.TranslateMessage("all pages")
    If (Not m_PDF Is Nothing) Then textAllPages = textAllPages & " " & g_Language.TranslateMessage("(%1 total)", m_PDF.GetPageCount())
    btsPages.AddItem textAllPages, 0
    btsPages.AddItem "current page only", 1
    btsPages.AddItem "custom range", 2
    btsPages.ListIndex = 0
    UpdatePagesUI
    
    'Background in PDFs typically assume white, but the user can toggle this
    btsTransparency.AddItem "solid color", 0
    btsTransparency.AddItem "transparent", 1
    btsTransparency.ListIndex = 0
    
    'Other rendering options are available
    btsAntialiasing.AddItem "optimize for screens", 0
    btsAntialiasing.AddItem "optimize for printing", 1
    btsAntialiasing.AddItem "off", 2
    btsAntialiasing.ListIndex = 0
    
    btsAnnotations.AddItem "on", 0
    btsAnnotations.AddItem "off", 1
    btsAnnotations.ListIndex = 1
    
    'PDFs supply their size in points.  We need to convert this to pixels to set a default size.
    
    'Retrieve the dimensions of the first page IN POINTS
    If (Not m_PDF Is Nothing) Then
        
        'Ensure a page is loaded
        m_PDF.LoadPage 0
        
        m_baseImageWidthInPts = m_PDF.GetPageWidthInPoints()
        m_baseImageHeightInPts = m_PDF.GetPageHeightInPoints()
        
        'Use the pts dimensions to calculate a page size IN PIXELS.
        m_baseImageWidthInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageWidthInPts, DEFAULT_DPI))
        m_baseImageHeightInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageHeightInPts, DEFAULT_DPI))
        
        'Use the pixel measurements to initialize the resize box.  Note that this uses a default DPI value,
        ' but we can override this in a later step (after the command bar has initialized and retrieved the
        ' user's last-used values).
        rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
        rszUI.AspectRatioLock = True
        
        'Because a source PDF object was supplied, we can display basic page count and dimension information.
        lblOriginal(0).Caption = g_Language.TranslateMessage("%1 pages", m_PDF.GetPageCount())
        
        'If page size is uniform, display it alongsize page count.
        If m_PDF.IsPageSizeUniform() Then
            lblOriginal(1).Caption = g_Language.TranslateMessage("page size:  %1", GetPageSizeAsString(0))
        Else
            lblOriginal(1).Caption = g_Language.TranslateMessage("page size varies (first page is %1)", GetPageSizeAsString(0))
        End If
        
    'If no source PDF was supplied, default to screen size and blank out the "original PDF settings" box
    Else
        rszUI.SetInitialDimensions g_Displays.GetDesktopWidth, g_Displays.GetDesktopHeight
        rszUI.AspectRatioLock = False
        lblOriginal(0).Caption = g_Language.TranslateMessage("unknown")
        lblOriginal(1).Caption = vbNullString
    End If
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "PDF")
    
    'Allow previews
    m_SuspendRedraws = False
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'Return a given page's dimensions as a nicely formatted string (e.g.'8.5" x 11"')
Private Function GetPageSizeAsString(Optional ByVal idxPage As Long = 0) As String
    
    If (Not m_PDF Is Nothing) Then
    
        'Retrieve page dimensions from the PDF object
        Dim pWidth As Single, pHeight As Single
        pWidth = m_PDF.GetPageWidthInPoints_ByIndex(idxPage)
        pHeight = m_PDF.GetPageHeightInPoints_ByIndex(idxPage)
        
        'Convert that dimension to inches and cm
        Dim inchWidth As Single, inchHeight As Single
        inchWidth = pWidth / 72!
        inchHeight = pHeight / 72!
        
        Dim cmWidth As Single, cmHeight As Single
        cmWidth = Units.GetCMFromInches(inchWidth)
        cmHeight = Units.GetCMFromInches(inchHeight)
        
        'Format by current OS locale
        If Units.LocaleUsesMetric() Then
            GetPageSizeAsString = g_Language.TranslateMessage("%1 cm x %2 cm", Format$(cmWidth, "0.0#"), Format$(cmHeight, "0.0#"))
        Else
            GetPageSizeAsString = g_Language.TranslateMessage("%1"" x %2""", Format$(inchWidth, "0.0#"), Format$(inchHeight, "0.0#"))
        End If
        
    End If
    
End Function

'Update the preview window with a preview of the current PDF page
Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    If m_SuspendRedraws Then Exit Sub
    
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundDC targetDC
    cSurface.SetSurfaceAntialiasing P2_AA_None
    
    'Fill the background with a neutral color
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
    PD2D.FillRectangleI cSurface, cBrush, 0, 0, ctlWidth, ctlHeight
    
    'Ensure we have a preview object to use
    If (Not m_PDF Is Nothing) Then
        If m_PDF.HasPDF() Then
            
            'TODO: bounding box will need to be accounted for here... eventually
            
            'Select the currently active page
            If (sldPreview.Value < 1) Or (sldPreview.Value > m_PDF.GetPageCount) Then
                m_PDF.LoadPage 0
            Else
                m_PDF.LoadPage sldPreview.Value - 1
            End If
            
            'Prep a temporary DIB the size of the preview picture box, *but with aspect ratio preserved*
            ' against the source DIB's dimensions.
            Dim newWidth As Long, newHeight As Long
            PDMath.ConvertAspectRatio m_baseImageWidthInPx, m_baseImageHeightInPx, picPreview.GetWidth - 2, picPreview.GetHeight - 2, newWidth, newHeight
            
            'TODO: adjust background color painting per user settings
            Dim previewDIB As pdDIB
            Set previewDIB = New pdDIB
            If (btsTransparency.ListIndex = 0) Then
                previewDIB.CreateBlank newWidth, newHeight, 32, clsBackground.Color, 255
            Else
                previewDIB.CreateBlank newWidth, newHeight, 32, 0, 0
            End If
            previewDIB.SetInitialAlphaPremultiplicationState True
            
            'Convert UI settings to underlying library rendering flags
            
            'Antialias for displays
            Dim renderFlags As PDFium_RenderOptions: renderFlags = 0
            
            If (btsAntialiasing.ListIndex = 0) Then
                renderFlags = FPDF_LCD_TEXT
                
            'Antialias for printing
            ElseIf (btsAntialiasing.ListIndex = 1) Then
                renderFlags = FPDF_PRINTING
                
            'No antialiasing
            Else
                renderFlags = FPDF_RENDER_NO_SMOOTHIMAGE Or FPDF_RENDER_NO_SMOOTHPATH Or FPDF_RENDER_NO_SMOOTHTEXT
            End If
            
            If (btsAnnotations.ListIndex = 0) Then renderFlags = renderFlags Or FPDF_ANNOT
            
            'Ask the PDF object for a preview
            m_PDF.RenderCurrentPageToPDDib previewDIB, 0, 0, newWidth, newHeight, FPDF_Normal, renderFlags
            If (btsTransparency.ListIndex = 1) Then previewDIB.SetAlphaPremultiplication True, True
            
            'We now need to figure out positioning of the DIB in the target window (and we may need a checkerboard
            ' background behind it, too)
            Dim dstX As Long, dstY As Long
            dstX = (ctlWidth - previewDIB.GetDIBWidth) \ 2
            dstY = (ctlHeight - previewDIB.GetDIBHeight) \ 2
            
            PD2D.FillRectangleI cSurface, g_CheckerboardBrush, dstX, dstY, previewDIB.GetDIBWidth, previewDIB.GetDIBHeight
            previewDIB.AlphaBlendToDC targetDC, 255, dstX, dstY
            
        End If
        
    Else
        picPreview.PaintText "preview not available", 12, False, True
    End If

    'Render a border around the control too
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayDark)
    PD2D.DrawRectangleI cSurface, cPen, 0, 0, ctlWidth - 1, ctlHeight - 1

End Sub

Private Sub UpdatePanelVisibility()
    
    Dim i As Long
    For i = pnlOptions.lBound To pnlOptions.UBound
        pnlOptions(i).Visible = (i = btsPanel.ListIndex)
    Next i
    
End Sub

Private Sub UpdatePreview()
    picPreview.RequestRedraw True
End Sub

Private Sub sldPreview_Change()
    UpdatePreview
End Sub
