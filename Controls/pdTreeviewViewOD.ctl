VERSION 5.00
Begin VB.UserControl pdTreeviewViewOD 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdTreeviewViewOD.ctx":0000
End
Attribute VB_Name = "pdTreeviewViewOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Owner-Drawn Treeview View control (e.g. the list part of a treeview, *not* including the scroll bar)
'Copyright 2024-2026 by Tanner Helland
'Created: 05/September/24
'Last updated: 23/September/24
'Last update: wrap up keyboard support
'
'The list portion of a pdTreeviewOD (owner-drawn) object, with all drawing functionality provided as events
' that the parent control *must* respond to.  The list view manages all list data, and if no scroll bar is
' required, it's a fully functional listbox object.  If a scroll bar is required, however, you need to use
' the parent "pdTreeviewOD" control, which contains additional UI work for synchronizing against a scroll bar.
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
Public Event DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, ByRef itemID As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToItemRectF As Long, ByVal ptrToCaptionRectF As Long, ByVal ptrToControlRectF As Long)

'If you want to handle something like custom tooltips, a MouseOver event helps
Public Event MouseLeave()
Public Event MouseOver(ByVal itemIndex As Long, ByRef itemTextEn As String)

'It also relays some events from the list box management class
Public Event ScrollMaxChanged(ByVal newMax As Long)
Public Event ScrollValueChanged(ByVal newValue As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Height (in pixels) of a given list entry.  For best results, set this before adding any items to the list,
' and do not change the value once set.  Note that the caller must handle their own DPI adjustments and padding
' (if any), as this class performs no drawing of its own.
Private m_ListItemHeight As Long

'The rectangle where the list is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_ListRect As RectF

'List box support class.  Handles data storage and coordinate math for rendering.
Private WithEvents listSupport As pdTreeSupport
Attribute listSupport.VB_VarHelpID = -1

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Most owner-drawn listboxes use the same general visual behavior as other listboxes (e.g. glowing outline on hover,
' control outline on focus, etc).  Some may choose to suspend this behavior in favor of a custom solution, however.
Private m_BorderlessMode As Boolean

'The last cursor requested by our owner.  During MouseMove events, we will default to
' this cursor code (typically the "hand" cursor).
Private m_LastCursor As SystemCursorConstant

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDLISTBOXOD_COLOR_LIST
    [_First] = 0
    PDLB_Background = 0
    PDLB_Border = 1
    PDLB_SelectedItemFill = 2
    PDLB_SelectedItemBorder = 3
    PDLB_UnselectedItemFill = 4
    PDLB_UnselectedItemBorder = 5
    PDLB_SeparatorLine = 6
    [_Last] = 6
    [_Count] = 7
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_TreeviewViewOD
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

Public Property Get BorderlessMode() As Boolean
    BorderlessMode = m_BorderlessMode
End Property

Public Property Let BorderlessMode(ByVal newMode As Boolean)
    If (newMode <> m_BorderlessMode) Then
        m_BorderlessMode = newMode
        UpdateControlLayout
    End If
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'Instead of using a fontsize to determine rendering metrics, owner-drawn list boxes require the owner to know
' the desired list item size (in pixels) in advance.  Do not change this value after adding items to the listbox,
' as it forces expensive rendering recalculations.
Public Property Get ListItemHeight() As Long
    ListItemHeight = m_ListItemHeight
End Property

Public Property Let ListItemHeight(ByVal newSize As Long)
    m_ListItemHeight = newSize
    listSupport.DefaultItemHeight = m_ListItemHeight
    PropertyChanged "ListItemHeight"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
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

Private Sub listSupport_Click()
    RaiseEvent Click
End Sub

'When the list manager detects that an action requires the list to be redrawn (like adding a new item),
' it raises this event.  Whether or not we respond depends on the user control's visibility.
Private Sub listSupport_RedrawNeeded()
    If ucSupport.AmIVisible Then RedrawBackBuffer True
End Sub

Private Sub listSupport_ScrollMaxChanged()
    RaiseEvent ScrollMaxChanged(listSupport.ScrollMax)
End Sub

Private Sub listSupport_ScrollValueChanged()
    RaiseEvent ScrollValueChanged(Me.ScrollValue)
End Sub

Public Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    listSupport.NotifyKeyDown Shift, vkCode, markEventHandled
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseClick Button, Shift, x, y
    UpdateMousePosition
End Sub

Private Sub ucSupport_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseDoubleClick Button, Shift, x, y
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    listSupport.NotifyKeyDown Shift, vkCode, markEventHandled
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    listSupport.NotifyKeyUp Shift, vkCode, markEventHandled
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    listSupport.NotifyMouseDown Button, Shift, x, y
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseEnter Button, Shift, x, y
    UpdateMousePosition
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseLeave Button, Shift, x, y
    UpdateMousePosition
    RaiseEvent MouseLeave
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    listSupport.NotifyMouseMove Button, Shift, x, y
    UpdateMousePosition
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    listSupport.NotifyMouseUp Button, Shift, x, y, clickEventAlsoFiring
End Sub

Private Sub UpdateMousePosition()
    If (listSupport.ListIndexHovered >= 0) Then
        ucSupport.RequestCursor m_LastCursor
        RaiseEvent MouseOver(listSupport.ListIndexHovered, listSupport.List(listSupport.ListIndexHovered))
    Else
        ucSupport.RequestCursor IDC_DEFAULT
    End If
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    listSupport.NotifyMouseWheelVertical Button, Shift, x, y, scrollAmount
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If newVisibility Then listSupport.SetAutomaticRedraws True, True
End Sub

Public Property Get HasFocus() As Boolean
    HasFocus = ucSupport.DoIHaveFocus()
End Property

'Listbox-specific functions and subs.  Most of these simply relay the request to the listSupport object, and it will
' raise redraw requests as relevant.
Public Sub AddItem(ByRef srcItemID As String, ByRef srcItemText As String, Optional ByRef parentID As String = vbNullString, Optional ByVal initialCollapsedState As Boolean = False)
    listSupport.AddItem srcItemID, srcItemText, parentID, initialCollapsedState
End Sub

Public Sub Clear()
    listSupport.Clear
End Sub

Public Function GetDefaultItemHeight() As Long
    GetDefaultItemHeight = listSupport.DefaultItemHeight
End Function

'If you need detailed access to underlying tree data, you can access the tree support object here
Public Function AccessUnderlyingTreeSupport() As pdTreeSupport
    Set AccessUnderlyingTreeSupport = listSupport
End Function

Public Function List(ByVal itemIndex As Long) As String
    List = listSupport.List(itemIndex)
End Function

Public Function ListCount() As Long
    ListCount = listSupport.ListCount
End Function

Public Property Get ListIndex() As Long
    ListIndex = listSupport.ListIndex
End Property

Public Property Let ListIndex(ByVal newIndex As Long)
    listSupport.ListIndex = newIndex
End Property

Public Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) As Long
    ListIndexByString = listSupport.ListIndexByString(srcString, compareMode)
End Function

Public Function ListIndexByPosition(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal checkXAsWell As Boolean = True) As Long
    ListIndexByPosition = listSupport.ListIndexByPosition(srcX, srcY, checkXAsWell)
End Function

'In response to things like MouseOver events, the caller can request different cursors.
' (By default, list items are always treated as clickable - so they get a hand cursor.)
Public Sub RequestCursor(Optional ByVal sysCursorID As SystemCursorConstant = IDC_HAND)
    ucSupport.RequestCursor sysCursorID
    m_LastCursor = sysCursorID
End Sub

'The caller can suspend automatic redraws caused by things like adding an item to the list box.  Just make sure to enable redraws
' once you're ready, or you'll never get rendering requests!
Public Sub SetAutomaticRedraws(ByVal newState As Boolean, Optional ByVal raiseRedrawImmediately As Boolean = False)
    listSupport.SetAutomaticRedraws newState, raiseRedrawImmediately
End Sub

Public Function ShouldScrollBarBeVisible() As Boolean
    ShouldScrollBarBeVisible = (Me.ScrollMax > 0)
End Function

Public Function ScrollMax() As Long
    ScrollMax = listSupport.ScrollMax
End Function

Public Property Get ScrollValue() As Long
    ScrollValue = listSupport.ScrollValue()
End Property

Public Property Let ScrollValue(ByRef newValue As Long)
    listSupport.ScrollValue = newValue
End Property

Public Sub RequestListRedraw()
    RedrawBackBuffer True
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    
    'Request any keys related to tree navigation
    ucSupport.SpecifyRequiredKeys VK_DOWN, VK_UP, VK_RIGHT, VK_LEFT, VK_PAGEDOWN, VK_PAGEUP, VK_HOME, VK_END, VK_RETURN, VK_SPACE, VK_OEM_PLUS, VK_OEM_MINUS, VK_ADD, VK_SUBTRACT, VK_MULTIPLY
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLISTBOXOD_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDListBoxView", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    'Initialize a helper list class; it manages the actual list data, and a bunch of rendering and layout decisions
    Set listSupport = New pdTreeSupport
    listSupport.SetAutomaticRedraws False
    
    'Set any other typical defaults
    m_LastCursor = IDC_HAND
    
End Sub

Private Sub UserControl_InitProperties()
    BorderlessMode = False
    Enabled = True
    ListItemHeight = 36
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent CustomDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent CustomDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BorderlessMode = .ReadProperty("BorderlessMode", False)
        Enabled = .ReadProperty("Enabled", True)
        ListItemHeight = .ReadProperty("ListItemHeight", 10)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BorderlessMode", m_BorderlessMode, False
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "ListItemHeight", Me.ListItemHeight, 10
    End With
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Determine the position of the list rect.  While we don't necessarily use this at present, I include it in case we
    ' ever want something like chunky borders in the future
    With m_ListRect
        .Left = 1
        .Top = 1
        .Width = (bWidth - 2) - .Left
        .Height = (bHeight - 2) - .Top
    End With
    
    'Notify the list manager of our new size
    listSupport.NotifyParentRectF m_ListRect
            
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer(Optional ByVal forciblyRedrawScreen As Boolean = False)
    
    Dim enabledState As Boolean
    enabledState = Me.Enabled
    
    Dim finalBackColor As Long
    finalBackColor = m_Colors.RetrieveColor(PDLB_Background, enabledState)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, finalBackColor)
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If PDMain.IsProgramRunning() Then
    
        'Cache colors in advance, so we can simply reuse them in the inner loop
        Dim itemColorSelectedBorder As Long, itemColorSelectedFill As Long
        Dim itemColorSelectedBorderHover As Long, itemColorSelectedFillHover As Long
        Dim itemColorUnselectedBorder As Long, itemColorUnselectedFill As Long
        Dim itemColorUnselectedBorderHover As Long, itemColorUnselectedFillHover As Long
        Dim arrowColor As Long
        
        itemColorUnselectedBorder = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, False)
        itemColorUnselectedBorderHover = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, True)
        itemColorUnselectedFill = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, False)
        itemColorUnselectedFillHover = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, True)
        itemColorSelectedBorder = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, False)
        itemColorSelectedBorderHover = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, True)
        itemColorSelectedFill = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, False)
        itemColorSelectedFillHover = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, True)
        
        'Start by retrieving basic rendering metrics from the support object
        Dim firstItemIndex As Long, lastItemIndex As Long, listIsEmpty As Boolean
        listSupport.GetRenderingLimits firstItemIndex, lastItemIndex, listIsEmpty
        
        'If the list either 1) has keyboard focus, or 2) is actively being hovered by the mouse, we render
        ' it differently, using PD's standard hover behavior (accent colors and chunky border)
        Dim listHasFocus As Boolean
        listHasFocus = ucSupport.DoIHaveFocus Or listSupport.IsMouseInsideTreeView
        
        'pd2D is used for all rendering
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC bufferDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        Dim cPen As pd2DPen: Set cPen = New pd2DPen
        Dim cBrush As pd2DBrush: Set cBrush = New pd2DBrush
        
        If (Not listIsEmpty) Then
            
            Dim curListIndex As Long, curColor As Long
            curListIndex = listSupport.ListIndex
            
            Dim itemIsSelected As Boolean, itemIsHovered As Boolean
            Dim tmpListItem As PD_TreeItem
            Dim scrollOffsetX As Long, scrollOffsetY As Long
            Dim srcCaptionRect As RectF, srcControlRect As RectF, srcItemRect As RectF
            
            Dim i As Long
            For i = firstItemIndex To lastItemIndex
            
                'For each list item, we follow a pretty standard formula: retrieve the item's data...
                If listSupport.GetRenderingItem(i, tmpListItem, scrollOffsetX, scrollOffsetY) Then
                    
                    'Make a local copy of the source item's rectangles
                    srcItemRect = tmpListItem.ItemRect
                    srcCaptionRect = tmpListItem.captionRect
                    srcControlRect = tmpListItem.controlRect
                    
                    'Offset all rectangles by the retrieved scroll values
                    srcItemRect.Top = srcItemRect.Top - scrollOffsetY
                    srcCaptionRect.Top = srcCaptionRect.Top - scrollOffsetY
                    srcControlRect.Top = srcControlRect.Top - scrollOffsetY
                    
                    'Add necessary offsets for the chunky line we draw around the entire control if
                    ' the list has focus.
                    srcItemRect.Left = srcItemRect.Left + 1!
                    srcItemRect.Width = srcItemRect.Width - 2!
                    srcCaptionRect.Width = srcCaptionRect.Width - 2!
                    
                    itemIsSelected = (i = curListIndex)
                    itemIsHovered = (i = listSupport.ListIndexHovered)
                    
                    '...then render its background fill...
                    If itemIsSelected Then
                        If itemIsHovered Then curColor = itemColorSelectedFillHover Else curColor = itemColorSelectedFill
                    Else
                        If itemIsHovered Then curColor = itemColorUnselectedFillHover Else curColor = itemColorUnselectedFill
                    End If
                    
                    cBrush.SetBrushColor curColor
                    PD2D.FillRectangleF_FromRectF cSurface, cBrush, srcItemRect
                    
                    'Next, let's draw the expand/contract caret
                    If tmpListItem.hasChildren Then
                    
                        If itemIsSelected Then
                            arrowColor = g_Themer.GetGenericUIColor(UI_TextClickableSelected)
                        Else
                            arrowColor = g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
                        End If
                        
                        Dim arrowPt1 As PointFloat, arrowPt2 As PointFloat, arrowPt3 As PointFloat
                        Dim arrowHeight As Single: arrowHeight = tmpListItem.ItemRect.Height / 4
                        
                        'Corresponding group is closed, so arrow points right
                        If tmpListItem.isCollapsed Then
                        
                            arrowPt1.x = srcControlRect.Left + (srcControlRect.Height / 2) + 3 - (arrowHeight / 2)
                            arrowPt1.y = srcControlRect.Top + arrowHeight * 1.3 - 1.5
                            
                            arrowPt3.x = arrowPt1.x
                            arrowPt3.y = (srcControlRect.Top + srcControlRect.Height) - arrowHeight * 1.3 + 1.5
                        
                            arrowPt2.x = arrowPt1.x + (arrowHeight / 2) + 0.5
                            arrowPt2.y = arrowPt1.y + (arrowPt3.y - arrowPt1.y) / 2
                        
                        'Corresponding group is open, so arrow points down
                        Else
                        
                            arrowPt1.x = srcControlRect.Left + arrowHeight * 1.3 - 0.5
                            arrowPt1.y = srcControlRect.Top + (srcControlRect.Height / 2) + 3 - (arrowHeight / 2)
                            
                            arrowPt3.x = (srcControlRect.Left + srcControlRect.Width) - arrowHeight * 1.3 - 0.5
                            arrowPt3.y = arrowPt1.y
                            
                            arrowPt2.x = arrowPt1.x + (arrowPt3.x - arrowPt1.x) / 2
                            arrowPt2.y = arrowPt1.y + (arrowHeight / 2) + 0.5
                            
                        End If
                        
                        'Draw the drop-down caret
                        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
                        Drawing2D.QuickCreateSolidPen cPen, 2!, arrowColor, 100!, P2_LJ_Round, P2_LC_Round
                        PD2D.DrawLineF_FromPtF cSurface, cPen, arrowPt1, arrowPt2
                        PD2D.DrawLineF_FromPtF cSurface, cPen, arrowPt2, arrowPt3
                        cSurface.SetSurfaceAntialiasing P2_AA_None
                        
                    '/end item has children
                    End If
        
                    '...then interject an event, so our parent can draw the remainder of this object
                    RaiseEvent DrawListEntry(bufferDC, i, tmpListItem.textEn, itemIsSelected, itemIsHovered, VarPtr(srcItemRect), VarPtr(srcCaptionRect), VarPtr(srcControlRect))
                    
                    '...then paint its hover-specific border over the top...
                    If itemIsSelected Then
                        If itemIsHovered Then curColor = itemColorSelectedBorderHover Else curColor = itemColorSelectedBorder
                    Else
                        If itemIsHovered Then curColor = itemColorUnselectedBorderHover Else curColor = itemColorUnselectedBorder
                    End If
                    
                    ' (As of the 7.0 release, the border is only drawn if the current item is selected.  This is a deliberate decision
                    '  to improve aesthetics on the Metadata dialog, among others.  This may be revisited in the future.
                    '  Note also that the caller can manually request borderless rendering via the matching property.)
                    If ((itemIsHovered Or itemIsSelected) And (Not m_BorderlessMode)) Then
                        cPen.SetPenWidth 1
                        cPen.SetPenLineJoin P2_LJ_Miter
                        cPen.SetPenColor curColor
                        PD2D.DrawRectangleF_FromRectF cSurface, cPen, srcItemRect
                    End If
                
                '/End "item is collapsed"
                End If
                
                '...and finally, render a separator line, if any
                'If itemHasSeparator Then
                '    lineY = tmpRect.Top + tmpHeightWithoutSeparator + (tmpHeight - tmpHeightWithoutSeparator) / 2
                '    cPen.SetPenColor separatorColor
                '    PD2D.DrawLineF cSurface, cPen, m_ListRect.Left + FixDPI(12), lineY, m_ListRect.Left + m_ListRect.Width - FixDPI(12), lineY
                'End If
                
            Next i
            
        End If
        
        'Last of all, we render the listbox border.  Note that we actually draw *two* borders.  The actual border,
        ' which is slightly inset from the list box boundaries, then a second border - pure background, erasing any item
        ' rendering that may have fallen outside the clipping area.
        If (Not m_BorderlessMode) Then
        
            Dim borderWidth As Single, borderColor As Long
            If listHasFocus Then borderWidth = 3! Else borderWidth = 1!
            borderColor = m_Colors.RetrieveColor(PDLB_Border, enabledState, listHasFocus)
            
            cPen.SetPenWidth borderWidth
            cPen.SetPenColor borderColor
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ListRect
            
            If (Not listHasFocus) Then
                cPen.SetPenWidth 1
                cPen.SetPenColor finalBackColor
                PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, 0, 0, bWidth - 1, bHeight - 1
            End If
            
        End If
        
        Set cSurface = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint forciblyRedrawScreen
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDLB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDLB_Border, "Border", IDE_GRAY
        .LoadThemeColor PDLB_SelectedItemFill, "SelectedItemFill", IDE_BLUE
        .LoadThemeColor PDLB_SelectedItemBorder, "SelectedItemBorder", IDE_BLUE
        .LoadThemeColor PDLB_UnselectedItemFill, "UnselectedItemFill", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemBorder, "UnselectedItemBorder", IDE_WHITE
        .LoadThemeColor PDLB_SeparatorLine, "SeparatorLine", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD avoids design-time tooltips (localizing these is hard).
' Instead, apply tooltips at run-time with this function.
' (IMPORTANT NOTE: translations of passed text are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
