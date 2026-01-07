VERSION 5.00
Begin VB.Form dialog_ImportSVG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5385
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
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdResize rszUI 
      Height          =   2895
      Left            =   4800
      TabIndex        =   1
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
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
      Top             =   4635
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
End
Attribute VB_Name = "dialog_ImportSVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'SVG Import Dialog (also works on Windows metafiles, e.g. EMF/WMF)
'Copyright 2022-2026 by Tanner Helland
'Created: 02/March/22
'Last updated: 09/December/22
'Last update: expand the dialog to work with Windows metafiles too (EMF/WMF)
'
'PhotoDemon offers to losslessly resize vector images at load-time.  This dialog provides the UI
' for this feature.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Handle to a resvg tree (for rendering SVG previews).  When non-zero, please set m_hGdipImage to 0.
Private m_hResvgTree As Long

'Handle to a GDI+ image handle (for rendering EMF/WMF previews).  When non-zero, please set m_hResvgTree to 0.
Private m_hGdipImage As Long

'SVG default width/height
Private m_DefaultWidth As Long, m_DefaultHeight As Long

'User-specified custom width/height, if any
Private m_UserWidth As Long, m_UserHeight As Long, m_UserDPI As Long

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetUserWidth() As Long
    GetUserWidth = m_UserWidth
End Function

Public Function GetUserHeight() As Long
    GetUserHeight = m_UserHeight
End Function

Public Function GetUserDPI() As Long
    GetUserDPI = m_UserDPI
End Function

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    
    'Cache resize values?
    m_UserWidth = rszUI.ResizeWidthInPixels
    m_UserHeight = rszUI.ResizeHeightInPixels
    m_UserDPI = rszUI.ResizeDPIAsPPI
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_ResetClick()
    rszUI.SetInitialDimensions m_DefaultWidth, m_DefaultHeight
    rszUI.AspectRatioLock = True
End Sub

Private Sub Form_Activate()
    rszUI.SetInitialDimensions m_DefaultWidth, m_DefaultHeight
    rszUI.AspectRatioLock = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal hResvgTree As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal hGdipImage As Long = 0&)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure a proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify import options... "
    
    m_DefaultWidth = srcWidth
    m_DefaultHeight = srcHeight
    rszUI.SetInitialDimensions m_DefaultWidth, m_DefaultHeight
    rszUI.AspectRatioLock = True
    
    m_hResvgTree = hResvgTree
    m_hGdipImage = hGdipImage
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    If (m_hResvgTree <> 0) Then
        Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "SVG")
    ElseIf (m_hGdipImage <> 0) Then
        Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "EMF")
    End If
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundDC targetDC
    cSurface.SetSurfaceAntialiasing P2_AA_None
    
    'Fill the background with a neutral color
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
    PD2D.FillRectangleI cSurface, cBrush, 0, 0, ctlWidth, ctlHeight
    
    'Prep either an SVG preview or an EMF/WMF preview
    If (m_hResvgTree <> 0) Or (m_hGdipImage <> 0) Then
        
        'Prep a temporary DIB the size of the preview picture box, *but with aspect ratio preserved*
        ' against the source DIB's dimensions.
        Dim newWidth As Long, newHeight As Long
        PDMath.ConvertAspectRatio m_DefaultWidth, m_DefaultHeight, picPreview.GetWidth - 2, picPreview.GetHeight - 2, newWidth, newHeight
        
        Dim previewDIB As pdDIB
        Set previewDIB = New pdDIB
        previewDIB.CreateBlank newWidth, newHeight, 32, 0, 0
        previewDIB.SetInitialAlphaPremultiplicationState True
        
        'Ask the appropriate renderer for a preview
        If (m_hResvgTree <> 0) Then
            Plugin_resvg.RenderToArbitraryDIB m_hResvgTree, previewDIB
        ElseIf (m_hGdipImage <> 0) Then
            GDI_Plus.PaintMetafileToArbitraryDIB previewDIB, m_hGdipImage
        End If
        
        'We now need to figure out positioning of the SVG in the target window (and we'll need a checkerboard
        ' background behind it, too)
        Dim dstX As Long, dstY As Long
        dstX = (ctlWidth - previewDIB.GetDIBWidth) \ 2
        dstY = (ctlHeight - previewDIB.GetDIBHeight) \ 2
        
        'GDI+/GDI intermixing does not always behave as expected.  Paint a checkerboard background,
        ' then free the GDI+ surface
        PD2D.FillRectangleI cSurface, g_CheckerboardBrush, dstX, dstY, previewDIB.GetDIBWidth, previewDIB.GetDIBHeight
        previewDIB.AlphaBlendToDC targetDC, 255, dstX, dstY
    
    Else
        picPreview.PaintText "preview not available", 12, False, True
    End If
    
    'Render a border around the control too
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayDark)
    PD2D.DrawRectangleI cSurface, cPen, 0, 0, ctlWidth - 1, ctlHeight - 1

End Sub
