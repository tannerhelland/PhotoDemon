VERSION 5.00
Begin VB.UserControl pdListBoxOD 
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
   ToolboxBitmap   =   "pdListBoxOD.ctx":0000
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
   Begin PhotoDemon.pdListBoxViewOD lbView 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2778
   End
End
Attribute VB_Name = "pdListBoxOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Owner-Drawn List Box control
'Copyright 2016-2026 by Tanner Helland
'Created: 26/March/16
'Last updated: 27/March/16
'Last update: continued migrating code from the default list box to this instance
'
'Unicode-compatible owner-drawn list box replacement.  Refer to the pdListSupport class and pdListViewOD sub-control
' for additional details; they handle most the heavy lifting for this control.  (This control instance's only job is
' synchronizing the owner-drawn listview and the scrollbar, as necessary.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control raises much fewer events than a standard ListBox, by design
Public Event Click()

'Drag/drop events are raised (these are just relays, identical to standard VB drag/drop events).
' Note that these are *only* raised by the child pdListBoxView object, and we simply relay them.
Public Event CustomDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event CustomDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

'Note that drawing events *must* be responded to!  If you don't handle them, your listbox won't display anything.
Public Event DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, ByRef itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

'If you want to handle something like custom tooltips, a MouseOver event helps.  (These are ultimately raised
' by the underlying pdListBoxViewOD control.)
Public Event MouseLeave()
Public Event MouseOver(ByVal itemIndex As Long, ByRef itemTextEn As String)

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
Private Enum PDLISTBOX_COLOR_LIST
    [_First] = 0
    PDLB_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ListBoxOD
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

Public Property Get BorderlessMode() As Boolean
    BorderlessMode = lbView.BorderlessMode
End Property

Public Property Let BorderlessMode(ByVal newMode As Boolean)
    lbView.BorderlessMode = newMode
End Property

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText()
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

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

'Instead of using a fontsize to determine rendering metrics, owner-drawn list boxes require the owner to know the desired list item
' size in advance.  Do not change this value after adding items to the listbox, as it forces expensive rendering recalculations.
Public Property Get ListItemHeight() As Long
    ListItemHeight = lbView.ListItemHeight
End Property

Public Property Let ListItemHeight(ByVal newSize As Long)
    lbView.ListItemHeight = newSize
    PropertyChanged "ListItemHeight"
End Property

Public Property Get FontSizeCaption() As Single
    FontSizeCaption = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSizeCaption(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSizeCaption"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Sub CloneExternalListSupport(ByRef srcListSupport As pdListSupport, Optional ByVal desiredListIndexTop As Long = 0, Optional ByVal newListSupportMode As PD_ListSupportMode = PDLM_LB_Inside_DD)
    lbView.CloneExternalListSupport srcListSupport, desiredListIndexTop, newListSupportMode
End Sub

'Helper functions to allow our parent to read/write scroll values
Public Function GetScrollValue() As Long
    If vScroll.Visible Then GetScrollValue = vScroll.Value Else GetScrollValue = 0
End Function

Public Sub SetScrollValue(ByVal newValue As Long)
    If vScroll.Visible Then
        If (newValue < vScroll.Min) Then newValue = vScroll.Min
        If (newValue > vScroll.Max) Then newValue = vScroll.Max
        vScroll.Value = newValue
    End If
End Sub

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

Public Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    lbView.NotifyKeyDown Shift, vkCode, markEventHandled
End Sub

'Listbox-specific functions and subs.  Most of these simply relay the request to the listSupport object, and it will
' raise redraw requests as relevant.
Public Sub AddItem(Optional ByVal srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1)
    lbView.AddItem srcItemText, itemIndex, hasTrailingSeparator, itemHeight
End Sub

Public Sub Clear()
    lbView.Clear
End Sub

Public Function List(ByVal itemIndex As Long, Optional ByVal returnTranslatedText As Boolean = False) As String
    List = lbView.List(itemIndex, returnTranslatedText)
End Function

Public Function ListCount() As Long
    ListCount = lbView.ListCount
End Function

Public Property Get ListIndex() As Long
    ListIndex = lbView.ListIndex
End Property

Public Property Let ListIndex(ByVal newIndex As Long)
    lbView.ListIndex = newIndex
End Property

Public Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) As Long
    ListIndexByString = lbView.ListIndexByString(srcString, compareMode)
End Function

Public Sub RemoveItem(ByVal itemIndex As Long)
    lbView.RemoveItem itemIndex
End Sub

'In response to things like MouseOver events, the caller can request different cursors.
' (By default, list items are always treated as clickable - so they get a hand cursor.)
Public Sub RequestCursor(Optional ByVal sysCursorID As SystemCursorConstant = IDC_HAND)
    lbView.RequestCursor sysCursorID
End Sub

'The caller can suspend automatic redraws caused by things like adding an item to the list box.  Just make sure to enable redraws
' once you're ready, or you'll never get rendering requests!
Public Sub SetAutomaticRedraws(ByVal newState As Boolean, Optional ByVal raiseRedrawImmediately As Boolean = False)
    lbView.SetAutomaticRedraws newState, raiseRedrawImmediately
End Sub

Private Sub lbView_Click()
    RaiseEvent Click
End Sub

Private Sub lbView_CustomDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent CustomDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub lbView_CustomDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent CustomDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub lbView_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    RaiseEvent DrawListEntry(bufferDC, itemIndex, itemTextEn, itemIsSelected, itemIsHovered, ptrToRectF)
End Sub

Private Sub lbView_MouseLeave()
    RaiseEvent MouseLeave
End Sub

Private Sub lbView_MouseOver(ByVal itemIndex As Long, itemTextEn As String)
    RaiseEvent MouseOver(itemIndex, itemTextEn)
End Sub

Private Sub lbView_ScrollMaxChanged(ByVal newMax As Long)
    
    If (vScroll.Visible <> lbView.ShouldScrollBarBeVisible) Then vScroll.Visible = lbView.ShouldScrollBarBeVisible
    If (newMax >= 0) Then vScroll.Max = newMax
    vScroll.LargeChange = lbView.GetDefaultItemHeight
    vScroll.Value = lbView.ScrollValue
    
    UpdateControlLayout
    
End Sub

Private Sub lbView_ScrollValueChanged(ByVal newValue As Long)
    If vScroll.Visible Then vScroll.Value = newValue
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

Private Sub VScroll_Scroll(ByVal eventIsCritical As Boolean)
    If (lbView.ScrollValue <> vScroll.Value) Then lbView.ScrollValue = vScroll.Value
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    ucSupport.RequestCaptionSupport False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLISTBOX_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDListBox", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    BorderlessMode = False
    Caption = vbNullString
    FontSizeCaption = 12
    ListItemHeight = 36
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BorderlessMode = .ReadProperty("BorderlessMode", False)
        Caption = .ReadProperty("Caption", vbNullString)
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12)
        ListItemHeight = .ReadProperty("ListItemHeight", 36)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BorderlessMode", lbView.BorderlessMode, False
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSizeCaption", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "ListItemHeight", lbView.ListItemHeight, 36
    End With
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'The list area is placed relative to the caption
        With m_InteractiveRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = bWidth - .Left
            .Height = bHeight - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_InteractiveRect
            .Left = 0
            .Top = 0
            .Width = bWidth - .Left
            .Height = bHeight - .Top
        End With
        
    End If
    
    'If the scrollbar is visible, we'll calculate its left-most position first.
    Dim lbRightPosition As Long, initScrollVisibility As Boolean
    
    initScrollVisibility = vScroll.Visible
    If lbView.ShouldScrollBarBeVisible Then
        lbRightPosition = (m_InteractiveRect.Width - vScroll.GetWidth)
    Else
        lbRightPosition = m_InteractiveRect.Left + m_InteractiveRect.Width
    End If
    
    'Move the listbox into position
    lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, lbRightPosition - m_InteractiveRect.Left, m_InteractiveRect.Height
    
    'Because the listbox is in a new position, it may or may not still need a scrollbar
    Dim scrollShouldBeVisible As Boolean
    scrollShouldBeVisible = lbView.ShouldScrollBarBeVisible
    
    vScroll.Visible = scrollShouldBeVisible
    If scrollShouldBeVisible Then
        vScroll.SetPositionAndSize lbRightPosition, m_InteractiveRect.Top + 1, vScroll.GetWidth, m_InteractiveRect.Height - 2
        lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, lbRightPosition - m_InteractiveRect.Left, m_InteractiveRect.Height
    Else
        lbView.SetPositionAndSize m_InteractiveRect.Left, m_InteractiveRect.Top, m_InteractiveRect.Width, m_InteractiveRect.Height
    End If
    
    'As a failsafe, synchronize scroll bar values if the scrollbar is visible
    If scrollShouldBeVisible Then
        vScroll.Max = lbView.ScrollMax
        vScroll.Value = lbView.ScrollValue
    End If
                
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    If (bufferDC = 0) Then Exit Sub
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDLB_Background, "Background", IDE_WHITE
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
    lbView.AssignTooltip newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
