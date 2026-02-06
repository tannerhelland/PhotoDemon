VERSION 5.00
Begin VB.UserControl pdFxPreviewCtl 
   AccessKeys      =   "T"
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ClipControls    =   0   'False
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
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "pdFxPreview.ctx":0000
   Begin PhotoDemon.pdPreview pdPreviewBox 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
   End
   Begin PhotoDemon.pdButtonStrip btsZoom 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      FontSize        =   9
   End
   Begin PhotoDemon.pdButtonStrip btsState 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      FontSize        =   9
   End
End
Attribute VB_Name = "pdFxPreviewCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Effect Preview custom control
'Copyright 2013-2026 by Tanner Helland
'Created: 10/January/13
'Last updated: 13/October/20
'Last update: rework resize code to allow for efficient run-time resizing of previews
'
'For the first decade of its life, PhotoDemon relied on simple picture boxes for rendering its effect previews.
' This worked well enough when there were only a handful of tools available, but as the complexity of the program
' - and its various effects and tools - has grown, it has become more and more painful to update the preview
' system, because any changes have to be mirrored across a huge number of forms.
'
'Thus, this control was born.  It is now used on every single effect form in place of a regular picture box.  This
' allows me to add preview-related features just once - to the base control - and have every tool automatically
' reap the benefits.
'
'The control is capable of storing a copy of the original image and any filter-modified versions of the image.
' The user can toggle between these by using the command link below the main picture box, or by pressing Alt+T.
' This replaces the side-by-side "before and after" of past versions.
'
'A few other extra features have been implemented, which can be enabled on a tool-by-tool basis.  Specifically:
' 1) The user can toggle between "fit image" and "100% zoom + click-drag-to-scroll" modes.  Note that 100% zoom
'    is not appropriate for some tools (i.e. perspective transformations and other algorithms that only operate
'    on the full image area).
' 2) Click-to-select color functionality.  This is helpful for tools that rely on color information within the
'    image for their operation, e.g. green screen.
' 3) Click-to-select-coordinate functionality.  This is helpful for giving the user an easy way to select a
'    location on the image as, say, a center point for a filter (e.g. vignetting works great with this).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Preview boxes can now let the user switch between "full image" and "100% zoom" states; we have to let the caller know about
' these events, because a new effect preview must be generated
Public Event ViewportChanged()
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Some preview boxes will let the user click to set a new centerpoint for a filter or effect.
Public Event PointSelected(xRatio As Double, yRatio As Double)

'Some preview boxes allow the user to click and select a color from the source image
Public Event ColorSelected()

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDFXPREVIEW_COLOR_LIST
    [_First] = 0
    PDFX_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_FxPreviewCtl
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'At design-time, use this property to determine whether the user is allowed to select colors directly from the
' preview window (helpful for tools like green screen, etc).
Public Property Get AllowColorSelection() As Boolean
    AllowColorSelection = pdPreviewBox.AllowColorSelection
End Property

Public Property Let AllowColorSelection(ByVal isAllowed As Boolean)
    pdPreviewBox.AllowColorSelection = isAllowed
    PropertyChanged "AllowColorSelection"
End Property

'At design-time, use this property to determine whether the user is allowed to select new center points for a filter
' or effect by clicking the preview window.
Public Property Get AllowPointSelection() As Boolean
    AllowPointSelection = pdPreviewBox.AllowPointSelection
End Property

Public Property Let AllowPointSelection(ByVal isAllowed As Boolean)
    pdPreviewBox.AllowPointSelection = isAllowed
    PropertyChanged "AllowPointSelection"
End Property

'At design-time, use this property to prevent the user from changing the preview area between zoom/pan and fit mode.
Public Property Get AllowZoomPan() As Boolean
    AllowZoomPan = pdPreviewBox.AllowZoomPan
End Property

Public Property Let AllowZoomPan(ByVal isAllowed As Boolean)
    pdPreviewBox.AllowZoomPan = isAllowed
    PropertyChanged "DisableZoomPan"
    If (PDMain.IsProgramRunning()) Then UpdateControlLayout
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    If PDMain.IsProgramRunning() Then RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Function GetUniqueID() As Double
    GetUniqueID = pdPreviewBox.GetUniqueID
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'OffsetX/Y are used when the preview is in 1:1 mode, and the user is allowed to scroll around the underlying image
Public Property Get GetOffsetX() As Long
    GetOffsetX = pdPreviewBox.GetOffsetX
End Property

Public Property Get GetOffsetY() As Long
    GetOffsetY = pdPreviewBox.GetOffsetY
End Property

'External functions may need to access the color selected by the preview control
Public Property Get SelectedColor() As Long
    SelectedColor = pdPreviewBox.SelectedColor
End Property

Public Property Get ViewportFitFullImage() As Boolean
    ViewportFitFullImage = (btsZoom.ListIndex = 1)
End Property

Private Sub pdPreviewBox_ColorSelected()
    RaiseEvent ColorSelected
End Sub

Private Sub pdPreviewBox_PointSelected(xRatio As Double, yRatio As Double)
    RaiseEvent PointSelected(xRatio, yRatio)
End Sub

Private Sub pdPreviewBox_ViewportChanged()
    If ucSupport.AmIVisible Then RaiseEvent ViewportChanged
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_DIALOG_RESIZE_FINISHED) Then pdPreviewBox.RequestImmediateRefresh
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

'Use this to supply the preview with a copy of the original image's data.  The preview object can use this to display
' the original image when the user clicks the "show original image" link.
Public Sub SetOriginalImage(ByRef srcDIB As pdDIB)
    pdPreviewBox.SetOriginalImage srcDIB
End Sub

'Use this to supply the object with a copy of the processed image's data.  The preview object can use this to display
' the processed image again if the user clicks the "show original image" link, then clicks it again.
Public Sub SetFXImage(ByRef srcDIB As pdDIB, Optional ByVal colorManagementAlreadyHandled As Boolean = False)
    pdPreviewBox.SetFXImage srcDIB, colorManagementAlreadyHandled
End Sub

'Has this preview control had an original version of the image set?
Public Function HasOriginalImage() As Boolean
    HasOriginalImage = pdPreviewBox.HasOriginalImage
End Function

'Return dimensions of the preview picture box
Public Function GetPreviewWidth() As Long
    GetPreviewWidth = pdPreviewBox.GetPreviewWidth
End Function

Public Function GetPreviewHeight() As Long
    GetPreviewHeight = pdPreviewBox.GetPreviewHeight
End Function

Public Sub NotifyNonStandardSource(ByVal srcWidth As Long, ByVal srcHeight As Long)
    pdPreviewBox.NotifyNonStandardSource srcWidth, srcHeight
End Sub

Private Sub btsState_Click(ByVal buttonIndex As Long)
    pdPreviewBox.ShowOriginalInstead = (buttonIndex = 0)
End Sub

'When zoom state changes, we must raise a viewport change event so the effect can be redrawn.
Private Sub btsZoom_Click(ByVal buttonIndex As Long)
    pdPreviewBox.ViewportFitFullImage = (buttonIndex = 1)
End Sub

'When the control's access key is pressed (alt+t) , toggle the original/current image
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    btsState.ListIndex = 1 - btsState.ListIndex
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    ucSupport.SubclassCustomMessage WM_PD_DIALOG_RESIZE_FINISHED, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDFXPREVIEW_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDFXPreview", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Prep the various buttonstrips
    btsState.AddItem "before", 0
    btsState.AddItem "after", 1
    btsState.ListIndex = 1
    
    btsZoom.AddItem "1:1", 0
    btsZoom.AddItem "fit", 1
    btsZoom.ListIndex = 1
            
End Sub

'Initialize our effect preview control
Private Sub UserControl_InitProperties()
    
    'By default, the control *allows* the user to zoom/pan the transformation
    AllowZoomPan = True
    
    'By default, the control does *not* allow the user to select coordinate points or colors by clicking the preview area
    AllowColorSelection = False
    AllowPointSelection = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        AllowColorSelection = .ReadProperty("ColorSelection", False)
        AllowPointSelection = .ReadProperty("PointSelection", False)
        AllowZoomPan = Not .ReadProperty("DisableZoomPan", False)
    End With
End Sub

'Redraw the user control after it has been resized
Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then UpdateControlLayout
End Sub

'After a resize or paint request, update the layout of our control
Private Sub UpdateControlLayout()
    
    If PDMain.IsProgramRunning() Then
        
        'The primary object in this control is the preview picture box.  Everything else is positioned relative to it.
        Dim newPicWidth As Long, newPicHeight As Long
        newPicWidth = ucSupport.GetControlWidth
        newPicHeight = ucSupport.GetControlHeight - (btsState.GetHeight + Interface.FixDPI(4))
        pdPreviewBox.SetPositionAndSize 0, 0, newPicWidth, newPicHeight
        
        'If zoom/pan is not allowed, hide that button entirely
        btsZoom.Visible = Me.AllowZoomPan
        
        'Adjust the button strips to appear just below the preview window
        Dim newButtonTop As Long, newButtonWidth As Long
        newButtonTop = ucSupport.GetControlHeight - btsState.GetHeight
        
        'If zoom/pan is still visible, split the horizontal difference between that button strip, and the before/after strip.
        If Me.AllowZoomPan Then
            newButtonWidth = (newPicWidth \ 2) - Interface.FixDPI(4)
            btsZoom.SetPositionAndSize newPicWidth - newButtonWidth, newButtonTop, newButtonWidth, btsState.GetHeight
            
        'If zoom/pan is NOT visible, let the before/after button have the entire horizontal space
        Else
            newButtonWidth = newPicWidth
        End If
        
        'Move the before/after toggle into place
        btsState.SetPositionAndSize 0, newButtonTop, newButtonWidth, btsState.GetHeight
    
    End If
                
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ColorSelection", AllowColorSelection, False
        .WriteProperty "DisableZoomPan", Not AllowZoomPan, False
        .WriteProperty "PointSelection", AllowPointSelection, False
    End With
End Sub

'This control currently handles border rendering around the preview area, so it *does* maintain a backbuffer that
' may need to be redrawn under certain circumstances.
Private Sub RedrawBackBuffer()
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDFX_Background))
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDFX_Background, "Background", IDE_WHITE
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        pdPreviewBox.UpdateAgainstCurrentTheme
        btsState.UpdateAgainstCurrentTheme
        btsZoom.UpdateAgainstCurrentTheme
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub
