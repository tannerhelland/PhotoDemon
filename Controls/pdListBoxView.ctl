VERSION 5.00
Begin VB.UserControl pdListBoxView 
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
   ToolboxBitmap   =   "pdListBoxView.ctx":0000
End
Attribute VB_Name = "pdListBoxView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon List Box View control (e.g. the list part of a list box, not including the scroll bar)
'Copyright 2015-2026 by Tanner Helland
'Created: 22/December/15
'Last updated: 18/February/16
'Last update: continued work on initial build
'
'The list portion of a pdListBox object.  The list view manages all the list data, and if no scroll bar is required,
' it is basically a fully functional listbox object.  If a scroll bar is required, however, you need to use the
' parent "ListBox" control.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control raises much fewer events than a standard ListBox, by design
Public Event Click()

'It also relays some events from the list box management class
Public Event ScrollMaxChanged(ByVal newMax As Long)
Public Event ScrollValueChanged(ByVal newValue As Long)

'Drag/drop events are raised (these are just relays, identical to standard VB drag/drop events)
Public Event CustomDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event CustomDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Font size of the listview.  This controls all rendering metrics, so try not to change it at run-time.
Private m_FontSize As Single

'Padding around individual list items.  This value is added to the default font metrics to arrive at a default
' per-item size.
Private Const LIST_PADDING_HORIZONTAL As Single = 4!
Private Const LIST_PADDING_VERTICAL As Single = 2!

'The rectangle where the list is actually rendered
Private m_ListRect As RectF

'Callers can set "file" display mode; when active, this will truncate displayed text
' to the available window space using PathCompactPathW
Private m_FileDisplayMode As Boolean
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function PathCompactPathW Lib "shlwapi" (ByVal hDC As Long, ByVal ptrToMaxPathString As Long, ByVal maxSizeInPixels As Long) As Long

'List box support class.  Handles data storage and coordinate math for rendering.
Private WithEvents listSupport As pdListSupport
Attribute listSupport.VB_VarHelpID = -1

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDLISTBOX_COLOR_LIST
    [_First] = 0
    PDLB_Background = 0
    PDLB_Border = 1
    PDLB_SelectedItemFill = 2
    PDLB_SelectedItemBorder = 3
    PDLB_SelectedItemText = 4
    PDLB_UnselectedItemFill = 5
    PDLB_UnselectedItemBorder = 6
    PDLB_UnselectedItemText = 7
    PDLB_SeparatorLine = 8
    [_Last] = 8
    [_Count] = 9
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ListBoxView
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    If ucSupport.AmIVisible() Then RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'Font settings other than size are not supported.  If you want specialized per-item rendering, use an owner-drawn list box
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    m_FontSize = newSize
    listSupport.DefaultItemHeight = Fonts.GetDefaultStringHeight(m_FontSize) + LIST_PADDING_VERTICAL * 2
    PropertyChanged "FontSize"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Sub CloneExternalListSupport(ByRef srcListSupport As pdListSupport, Optional ByVal desiredListIndexTop As Long = 0, Optional ByVal newListSupportMode As PD_ListSupportMode = PDLM_LB_Inside_DD)
    listSupport.CloneExternalListSupport srcListSupport, desiredListIndexTop, newListSupportMode
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

Private Sub listSupport_Click()
    RaiseEvent Click
End Sub

'When the list manager detects that an action requires the list to be redrawn (like adding a new item), it will raise
' this event.  Whether or not we respond depends on whether the user control is currently visible.
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
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    listSupport.NotifyMouseMove Button, Shift, x, y
    UpdateMousePosition
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    listSupport.NotifyMouseUp Button, Shift, x, y, clickEventAlsoFiring
End Sub

Private Sub UpdateMousePosition()
    If (listSupport.ListIndexHovered >= 0) Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
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

'Listbox-specific functions and subs.  Most of these simply relay the request to the listSupport object, and it will
' raise redraw requests as relevant.
Public Sub AddItem(Optional ByVal srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1)
    listSupport.AddItem srcItemText, itemIndex, hasTrailingSeparator, itemHeight
End Sub

Public Sub Clear()
    listSupport.Clear
End Sub

Public Function GetDefaultItemHeight() As Long
    GetDefaultItemHeight = listSupport.DefaultItemHeight
End Function

Public Function List(ByVal itemIndex As Long, Optional ByVal returnTranslatedText As Boolean = False) As String
    List = listSupport.List(itemIndex, returnTranslatedText)
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

Public Sub RemoveItem(ByVal itemIndex As Long)
    listSupport.RemoveItem itemIndex
End Sub

'The caller can suspend automatic redraws caused by things like adding an item to the list box.  Just make sure to enable redraws
' once you're ready, or you'll never get rendering requests!
Public Sub SetAutomaticRedraws(ByVal newState As Boolean, Optional ByVal raiseRedrawImmediately As Boolean = False)
    listSupport.SetAutomaticRedraws newState, raiseRedrawImmediately
End Sub

Public Sub SetDisplayMode_Files(ByVal newState As Boolean)
    m_FileDisplayMode = newState
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

Public Function UpdateItem(ByVal itemIndex As Long, ByVal newItemText As String) As Boolean
    UpdateItem = listSupport.UpdateItem(itemIndex, newItemText)
End Function

Public Sub RequestListRedraw()
    RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_DOWN, VK_UP, VK_PAGEDOWN, VK_PAGEUP, VK_HOME, VK_END, VK_RETURN, VK_SPACE
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLISTBOX_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDListBoxView", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Initialize a helper list class; it manages the actual list data, and a bunch of rendering and layout decisions
    Set listSupport = New pdListSupport
    listSupport.SetAutomaticRedraws False
    listSupport.ListSupportMode = PDLM_ListBox
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    FontSize = 10
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent CustomDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent CustomDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", Me.FontSize, 10
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
        Dim fontColorSelected As Long, fontColorSelectedHover As Long
        Dim fontColorUnselected As Long, fontColorUnselectedHover As Long
        Dim separatorColor As Long
        
        itemColorUnselectedBorder = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, False)
        itemColorUnselectedBorderHover = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, True)
        itemColorUnselectedFill = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, False)
        itemColorUnselectedFillHover = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, True)
        itemColorSelectedBorder = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, False)
        itemColorSelectedBorderHover = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, True)
        itemColorSelectedFill = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, False)
        itemColorSelectedFillHover = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, True)
            
        fontColorSelected = m_Colors.RetrieveColor(PDLB_SelectedItemText, enabledState, False, False)
        fontColorSelectedHover = m_Colors.RetrieveColor(PDLB_SelectedItemText, enabledState, False, True)
        fontColorUnselected = m_Colors.RetrieveColor(PDLB_UnselectedItemText, enabledState, False, False)
        fontColorUnselectedHover = m_Colors.RetrieveColor(PDLB_UnselectedItemText, enabledState, False, True)
        
        separatorColor = m_Colors.RetrieveColor(PDLB_SeparatorLine, enabledState, False, False)
        
        'Start by retrieving basic rendering metrics from the support object
        Dim firstItemIndex As Long, lastItemIndex As Long, listIsEmpty As Boolean
        listSupport.GetRenderingLimits firstItemIndex, lastItemIndex, listIsEmpty
        
        'If the list either 1) has keyboard focus, or 2) is actively being hovered by the mouse, we render
        ' it differently, using PD's standard hover behavior (accent colors and chunky border)
        Dim listHasFocus As Boolean
        listHasFocus = ucSupport.DoIHaveFocus Or listSupport.IsMouseInsideListBox
        
        'pd2D is used for painting
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        Drawing2D.QuickCreateSolidBrush cBrush
        Drawing2D.QuickCreateSolidPen cPen
        
        If (Not listIsEmpty) Then
            
            Dim curListIndex As Long, curColor As Long
            curListIndex = listSupport.ListIndex
            
            Dim itemIsSelected As Boolean, itemIsHovered As Boolean, itemHasSeparator As Boolean
            
            'This control doesn't maintain its own fonts; instead, it borrows it from the public PD UI font cache, as necessary
            Dim tmpFont As pdFont, textPadding As Single, strTruncate As String, targetText As String
            Set tmpFont = Fonts.GetMatchingUIFont(m_FontSize)
            tmpFont.AttachToDC bufferDC
            tmpFont.SetTextAlignment vbLeftJustify
                
            textPadding = LIST_PADDING_HORIZONTAL
            If listHasFocus Then textPadding = textPadding - 1
            
            Dim tmpTop As Long, tmpHeight As Long, tmpHeightWithoutSeparator As Long
            Dim lineY As Single
            Dim tmpListItem As PD_ListItem, tmpRect As RectF
            
            'Left and Width are the same for all list entries
            If listHasFocus Then
                tmpRect.Left = m_ListRect.Left + 2
                tmpRect.Width = m_ListRect.Width - 4
            Else
                tmpRect.Left = m_ListRect.Left + 1
                tmpRect.Width = m_ListRect.Width - 2
            End If
            
            Dim i As Long
            For i = firstItemIndex To lastItemIndex
                
                'For each list item, we follow a pretty standard formula: retrieve the item's data...
                listSupport.GetRenderingItem i, tmpListItem, tmpTop, tmpHeight, tmpHeightWithoutSeparator
                tmpRect.Top = tmpTop
                itemHasSeparator = tmpListItem.isSeparator
                If itemHasSeparator Then tmpRect.Height = tmpHeightWithoutSeparator - 1 Else tmpRect.Height = tmpHeight - 1
                
                itemIsSelected = (i = curListIndex)
                itemIsHovered = (i = listSupport.ListIndexHovered)
                
                '...then render its fill...
                If itemIsSelected Then
                    If itemIsHovered Then cBrush.SetBrushColor itemColorSelectedFillHover Else cBrush.SetBrushColor itemColorSelectedFill
                Else
                    If itemIsHovered Then cBrush.SetBrushColor itemColorUnselectedFillHover Else cBrush.SetBrushColor itemColorUnselectedFill
                End If
                
                PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRect
                
                '...followed by its border...
                If itemIsSelected Then
                    If itemIsHovered Then cPen.SetPenColor itemColorSelectedBorderHover Else cPen.SetPenColor itemColorSelectedBorder
                Else
                    If itemIsHovered Then cPen.SetPenColor itemColorUnselectedBorderHover Else cPen.SetPenColor itemColorUnselectedBorder
                End If
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, tmpRect
                
                '...and finally, its caption
                If itemIsSelected Then
                    If itemIsHovered Then curColor = fontColorSelectedHover Else curColor = fontColorSelected
                Else
                    If itemIsHovered Then curColor = fontColorUnselectedHover Else curColor = fontColorUnselected
                End If
                
                tmpFont.SetFontColor curColor
                
                If m_FileDisplayMode Then
                    strTruncate = String$(MAX_PATH_LEN, 0)
                    CopyMemoryStrict StrPtr(strTruncate), StrPtr(tmpListItem.textTranslated), PDMath.Min2Int(MAX_PATH_LEN * 2, LenB(tmpListItem.textTranslated))
                    If (PathCompactPathW(bufferDC, StrPtr(strTruncate), tmpRect.Width - textPadding * 2) <> 0) Then
                        targetText = Left$(strTruncate, lstrlenW(StrPtr(strTruncate)))
                    Else
                        targetText = tmpListItem.textTranslated
                    End If
                Else
                    targetText = tmpListItem.textTranslated
                End If
                
                tmpFont.FastRenderTextWithClipping tmpRect.Left + textPadding, tmpRect.Top + LIST_PADDING_VERTICAL, tmpRect.Width - LIST_PADDING_HORIZONTAL, tmpRect.Height - LIST_PADDING_VERTICAL, targetText, True, True, False
                
                'Separators are drawn separately, external to the other items
                If itemHasSeparator Then
                    lineY = tmpRect.Top + tmpHeightWithoutSeparator + (tmpHeight - tmpHeightWithoutSeparator) * 0.5
                    cPen.SetPenColor separatorColor
                    PD2D.DrawLineF cSurface, cPen, m_ListRect.Left + Interface.FixDPI(12), lineY, m_ListRect.Left + m_ListRect.Width - Interface.FixDPI(12), lineY
                End If
                
            Next i
            
            tmpFont.ReleaseFromDC
            Set tmpFont = Nothing
        
        End If
        
        'Last of all, we render the listbox border.  Note that we actually draw *two* borders.  The actual border,
        ' which is slightly inset from the list box boundaries, then a second border - pure white, erasing any item
        ' rendering that may have fallen outside the clipping area.
        Dim borderWidth As Single, borderColor As Long
        If listHasFocus Then borderWidth = 3! Else borderWidth = 1!
        borderColor = m_Colors.RetrieveColor(PDLB_Border, enabledState, listHasFocus)
        cPen.SetPenWidth borderWidth
        cPen.SetPenColor borderColor
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ListRect
        
        If (Not listHasFocus) Then
            cPen.SetPenColor finalBackColor
            cPen.SetPenWidth 1!
            PD2D.DrawRectangleI cSurface, cPen, 0, 0, bWidth - 1, bHeight - 1
        End If
        
        Set cPen = Nothing: Set cBrush = Nothing: Set cSurface = Nothing
        
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
        .LoadThemeColor PDLB_SelectedItemText, "SelectedItemText", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemFill, "UnselectedItemFill", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemBorder, "UnselectedItemBorder", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemText, "UnselectedItemText", IDE_BLACK
        .LoadThemeColor PDLB_SeparatorLine, "SeparatorLine", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        listSupport.UpdateAgainstCurrentTheme
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
