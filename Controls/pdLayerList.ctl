VERSION 5.00
Begin VB.UserControl pdLayerList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdLayerList.ctx":0000
   Begin PhotoDemon.pdScrollBar vScroll 
      Height          =   1575
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2778
   End
   Begin PhotoDemon.pdLayerListInner lbView 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2778
   End
End
Attribute VB_Name = "pdLayerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layer List control
'Copyright 2016-2026 by Tanner Helland
'Created: 18/September/15
'Last updated: 16/August/20
'Last update: set scroll bar SmallChange property based on 1/2 height of items in the layer list box;
'             this makes for much faster scrolling when click-holding scroll up/down
'
'Unicode-compatible layer list box replacement.  Refer to the pdLayerListInner sub-control for additional
' details; it handles most the heavy lifting for this control.  (This control instance's only job is
' synchronizing the listview and scrollbar, as necessary.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control raises much fewer events than a standard ListBox, by design
Public Event Click()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Because this control supports captions, the main interaction area (list + scrollbar) may be shifted slightly downward.
' The usable space of both objects is defined by this rect.
Private m_InteractiveRect As RectF

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDLAYERLIST_COLOR_LIST
    [_First] = 0
    PDLL_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_LayerList
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    lbView.Enabled = newValue
    vScroll.Enabled = newValue
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'To support high-DPI settings properly, we expose some specialized move+size functions
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

'External functions can request a redraw of the layer box by calling this function.  (This is necessary
' whenever layers are added, deleted, re-ordered, etc.)
Public Sub RequestRedraw(Optional ByVal refreshThumbnailCache As Boolean = True, Optional ByVal layerID As Long = -1)
    lbView.RequestRedraw refreshThumbnailCache, layerID
End Sub

'Layer list-specific functions and subs.  Most of these simply relay the request to the embedded
' lbView object, and it will raise redraw requests as relevant.
Private Sub lbView_ScrollMaxChanged(ByVal newMax As Long)
    
    Dim scrollVisible As Boolean
    scrollVisible = (newMax <> 0)
    If (vScroll.Visible <> scrollVisible) Then vScroll.Visible = scrollVisible
    
    If (newMax >= 0) Then vScroll.Max = newMax
    vScroll.LargeChange = lbView.GetListItemHeight()
    vScroll.SmallChange = lbView.GetListItemHeight() / 2
    vScroll.Value = lbView.ScrollValue
    
    UpdateControlLayout
    
End Sub

Private Sub lbView_ScrollValueChanged(ByVal newValue As Long)
    If vScroll.Visible Then vScroll.Value = newValue
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub VScroll_Scroll(ByVal eventIsCritical As Boolean)
    If (lbView.ScrollValue <> vScroll.Value) Then lbView.ScrollValue = vScroll.Value
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLAYERLIST_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDListBox", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    With m_InteractiveRect
        .Left = 0
        .Top = 0
        .Width = bWidth - .Left
        .Height = bHeight - .Top
    End With
    
    'If the scrollbar is visible, we'll calculate its left-most position first.
    Dim lbRightPosition As Long, initScrollVisibility As Boolean
    
    initScrollVisibility = vScroll.Visible
    If (lbView.ScrollMax > 0) Then
        lbRightPosition = (m_InteractiveRect.Width - vScroll.GetWidth)
    Else
        lbRightPosition = m_InteractiveRect.Left + m_InteractiveRect.Width
    End If
    
    'If we wanted to move the listbox into position, we could do it with this line:
    'lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, lbRightPosition - m_InteractiveRect.Left, m_InteractiveRect.Height
    
    '...however, we still need to figure out if a scroll bar is required.  If it is, we'll need to shrink the layer list
    ' horizontally to leave room for the scrollbar.  Rather than risk moving the listbox twice, let's figure out the
    ' scroll bar situation, then render everything all at once.
    Dim scrollShouldBeVisible As Boolean
    scrollShouldBeVisible = lbView.IsScrollbarRequiredForHeight(m_InteractiveRect.Height)
    
    vScroll.Visible = scrollShouldBeVisible
    If scrollShouldBeVisible Then
        
        vScroll.SetPositionAndSize lbRightPosition, m_InteractiveRect.Top + 1, vScroll.GetWidth, m_InteractiveRect.Height - 2
        lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, lbRightPosition - m_InteractiveRect.Left, m_InteractiveRect.Height
        
        'As a failsafe, synchronize scroll bar values if the scrollbar is visible
        vScroll.Max = lbView.ScrollMax
        vScroll.Value = lbView.ScrollValue
        
    Else
        lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, m_InteractiveRect.Width, m_InteractiveRect.Height
    End If
        
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDLL_Background, "Background", IDE_WHITE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        lbView.UpdateAgainstCurrentTheme
        vScroll.UpdateAgainstCurrentTheme
    End If
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
