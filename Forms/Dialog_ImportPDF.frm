VERSION 5.00
Begin VB.Form dialog_ImportPDF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13350
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
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   4800
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "image size"
      FontSize        =   12
   End
   Begin PhotoDemon.pdDropDown cboPreview 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdResize rszUI 
      Height          =   2895
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      DefaultToRealWorldUnits=   -1  'True
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   4335
      Left            =   120
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7646
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5415
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
End
Attribute VB_Name = "dialog_ImportPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PDF Import Dialog
'Copyright 2024-2024 by Tanner Helland
'Created: 29/February/24
'Last updated: 29/February/24
'Last update: use the SVG import dialog as the basis for a similar(ish) PDF import dialog
'
'Like Photoshop and GIMP (and probably others), PhotoDemon allows users to set their own PDF resolution,
' image size, and various other PDF-specific features at import-time.  This dialog provides a UI for
' those settings.
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
Private m_suspendRedraws As Boolean

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
    cParams.AddParam "final-width-px", finalWidthPx, True
    cParams.AddParam "final-height-px", finalHeightPx, True
    cParams.AddParam "final-dpi", finalDPI, True
    
    GetDialogParamString = cParams.GetParamString()
    
End Function

Private Sub cboPreview_Click()
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
    cboPreview.ListIndex = 0
    
End Sub

Private Sub cmdBar_ResetClick()
    cboPreview.ListIndex = 0
    rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
    rszUI.AspectRatioLock = True
End Sub

Private Sub Form_Activate()
    'TODO: investigate why these two lines were used on other dialogs?
    'rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
    'rszUI.AspectRatioLock = True
End Sub

Private Sub Form_Load()
    
    'Do not allow previews until the form is fully loaded
    m_suspendRedraws = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByRef srcPDF As pdPDF)
    
    'Do not provide any page previews until the load process is complete
    m_suspendRedraws = True
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure a proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify import options... "
    
    'All PDF info comes from the source PDF object, which has already loaded the target PDF
    Set m_PDF = srcPDF
    
    'Populate the preview dropdown with all available pages
    cmdBar.RequestPresetNoLoad cboPreview
    If (m_PDF.GetPageCount > 0) Then
        Dim i As Long
        For i = 0 To m_PDF.GetPageCount - 1
            cboPreview.AddItem g_Language.TranslateMessage("Page %1", i + 1), i
        Next i
    End If
    cboPreview.ListIndex = 0
    
    'PDFs supply their size in points.  We need to convert this to pixels to set a default size.
    
    'Retrieve the dimensions of the first page IN POINTS
    m_baseImageWidthInPts = m_PDF.GetPageWidthInPoints()
    m_baseImageHeightInPts = m_PDF.GetPageHeightInPoints()
    
    'Use this to calculate a page size IN PIXELS.
    m_baseImageWidthInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageWidthInPts, DEFAULT_DPI))
    m_baseImageHeightInPx = Int(Units.ConvertOtherUnitToPixels(mu_Points, m_baseImageHeightInPts, DEFAULT_DPI))
    
    'Use the pixel measurements to initialize the resize box.  Note that this uses a default DPI value,
    ' but we can override this in a later step (after the command bar has initialized and retrieved the
    ' user's last-used values).
    rszUI.SetInitialDimensions m_baseImageWidthInPx, m_baseImageHeightInPx, DEFAULT_DPI
    rszUI.AspectRatioLock = True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "PDF")
    
    'Allow previews
    m_suspendRedraws = False
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'Update the preview window with a preview of the current PDF page
Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    If m_suspendRedraws Then Exit Sub
    
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
            If (cboPreview.ListIndex < 0) Or (cboPreview.ListIndex >= m_PDF.GetPageCount) Then
                m_PDF.LoadPage 0
            Else
                m_PDF.LoadPage cboPreview.ListIndex
            End If
            
            'Prep a temporary DIB the size of the preview picture box, *but with aspect ratio preserved*
            ' against the source DIB's dimensions.
            Dim newWidth As Long, newHeight As Long
            PDMath.ConvertAspectRatio m_baseImageWidthInPx, m_baseImageHeightInPx, picPreview.GetWidth - 2, picPreview.GetHeight - 2, newWidth, newHeight
            
            'TODO: adjust background color painting per user settings
            Dim previewDIB As pdDIB
            Set previewDIB = New pdDIB
            previewDIB.CreateBlank newWidth, newHeight, 32, RGB(255, 255, 255), 255
            previewDIB.SetInitialAlphaPremultiplicationState True
            
            'Ask the PDF object for a preview
            m_PDF.RenderCurrentPageToPDDib previewDIB, 0, 0, newWidth, newHeight
            
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

Private Sub UpdatePreview()
    picPreview.RequestRedraw True
End Sub
