VERSION 5.00
Begin VB.UserControl pdListBoxView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "pdListBoxView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon List Box View control (e.g. the list part of a list box, not including the scroll bar)
'Copyright 2015-2016 by Tanner Helland
'Created: 22/December/15
'Last updated: 18/February/16
'Last update: continued work on initial build
'
'The list portion of a pdListBox object.  The list view manages all the list data, and if no scroll bar is required,
' it is basically a fully functional listbox object.  If a scroll bar is required, however, you need to use the
' parent "ListBox" control.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control raises much fewer events than a standard ListBox, by design
Public Event Click()

'It also relays some events from the list box management class
Public Event ScrollMaxChanged(ByVal newMax As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Font size of the listview.  This controls all rendering metrics, so try not to change it at run-time.
Private m_FontSize As Single

'Padding around individual list items.  This value is added to the default font metrics to arrive at a default
' per-item size.
Private Const LIST_PADDING_HORIZONTAL As Single = 4#
Private Const LIST_PADDING_VERTICAL As Single = 2#

'The rectangle where the list is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_ListRect As RECTF

'List box support class.  Handles data storage and coordinate math for rendering.
Private WithEvents listSupport As pdListSupport
Attribute listSupport.VB_VarHelpID = -1

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
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

'Font settings other than size are not supported.  If you want specialized per-item rendering, use an owner-drawn list box
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    m_FontSize = newSize
    listSupport.DefaultItemHeight = Font_Management.GetDefaultStringHeight(m_FontSize) + LIST_PADDING_VERTICAL * 2
    PropertyChanged "FontSize"
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

'When the list manager detects that an action requires the list to be redrawn (like adding a new item), it will raise
' this event.  Whether or not we respond depends on whether the user control is currently visible.
Private Sub listSupport_RedrawNeeded()
    If ucSupport.AmIVisible Then RedrawBackBuffer
End Sub

Private Sub listSupport_ScrollMaxChanged()
    RaiseEvent ScrollMaxChanged(listSupport.ScrollMax)
End Sub

Public Function GetScrollMax() As Long
    GetScrollMax = listSupport.ScrollMax
End Function

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseClick Button, Shift, x, y
    UpdateMousePosition x, y
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseEnter Button, Shift, x, y
    UpdateMousePosition x, y
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseLeave Button, Shift, x, y
    UpdateMousePosition -100, -100
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    listSupport.NotifyMouseMove Button, Shift, x, y
    UpdateMousePosition x, y
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    If listSupport.ListIndexHovered >= 0 Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If newVisibility Then listSupport.SetAutomaticRedraws True, True
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

'Listbox-specific functions and subs.  Most of these simply relay the request to the listSupport object, and it will
' raise redraw requests as relevant.
Public Sub AddItem(Optional ByVal srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1)
    listSupport.AddItem srcItemText, itemIndex, hasTrailingSeparator, itemHeight
End Sub

Public Sub Clear()
    listSupport.Clear
End Sub

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

Public Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare, Optional ByVal ReadValueOnly As Boolean = False) As Long
    ListIndexByString = listSupport.ListIndexByString(srcString, compareMode, ReadValueOnly)
End Function

Public Function ListIndexByPosition(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal checkXAsWell As Boolean = True) As Long
    ListIndexByPosition = listSupport.ListIndexByPosition(srcX, srcY, checkXAsWell)
End Function

Public Sub RemoveItem(ByVal itemIndex As Long)
    listSupport.RemoveItem itemIndex
End Sub

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
    RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    ucSupport.RequestExtraFunctionality True, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLISTBOX_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDListBoxView", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    'Initialize a helper list class; it manages the actual list data, and a bunch of rendering and layout decisions
    Set listSupport = New pdListSupport
    listSupport.SetAutomaticRedraws False
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    FontSize = 10
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not g_IsProgramRunning Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
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
Private Sub RedrawBackBuffer()
    
    Dim enabledState As Boolean
    enabledState = Me.Enabled
    
    Dim BackgroundColor As Long
    BackgroundColor = m_Colors.RetrieveColor(PDLB_Background, enabledState)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(True, BackgroundColor)
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Cache colors in advance, so we can simply reuse them in the inner loop
    Dim itemColorSelectedBorder As Long, itemColorSelectedFill As Long
    Dim itemColorSelectedBorderHover As Long, itemColorSelectedFillHover As Long
    Dim itemColorUnselectedBorder As Long, itemColorUnselectedFill As Long
    Dim itemColorUnselectedBorderHover As Long, itemColorUnselectedFillHover As Long
    Dim fontColorSelected As Long, fontColorSelectedHover As Long
    Dim fontColorUnselected As Long, fontColorUnselectedHover As Long
    
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
    
    If g_IsProgramRunning Then
        
        'Start by retrieving basic rendering metrics from the support object
        Dim firstItemIndex As Long, lastItemIndex As Long, listIsEmpty As Boolean
        listSupport.GetRenderingLimits firstItemIndex, lastItemIndex, listIsEmpty
        
        'If the list either 1) has keyboard focus, or 2) is actively being hovered by the mouse, we render
        ' it differently, using PD's standard hover behavior (accent colors and chunky border)
        Dim listHasFocus As Boolean
        listHasFocus = ucSupport.DoIHaveFocus Or listSupport.IsMouseInsideListBox
        
        If Not listIsEmpty Then
            
            Dim curListIndex As Long, curColor As Long, itemIsSelected As Boolean, itemIsHovered As Boolean
            curListIndex = listSupport.ListIndex
            
            'This control doesn't maintain its own fonts; instead, it borrows it from the public PD UI font cache, as necessary
            Dim tmpFont As pdFont, textPadding As Single
            Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSize)
            textPadding = LIST_PADDING_HORIZONTAL
            If listHasFocus Then textPadding = textPadding - 1
            
            Dim tmpListItem As PD_LISTITEM, tmpTop As Long, tmpHeight As Long, tmpRect As RECTF
            
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
                listSupport.GetRenderingItem i, tmpListItem, tmpTop, tmpHeight
                tmpRect.Top = tmpTop
                tmpRect.Height = tmpHeight - 1
                
                itemIsSelected = CBool(i = curListIndex)
                itemIsHovered = CBool(i = listSupport.ListIndexHovered)
                
                '...then render its fill...
                If itemIsSelected Then
                    If itemIsHovered Then curColor = itemColorSelectedFillHover Else curColor = itemColorSelectedFill
                Else
                    If itemIsHovered Then curColor = itemColorUnselectedFillHover Else curColor = itemColorUnselectedFill
                End If
                
                GDI_Plus.GDIPlusFillRectFToDC bufferDC, tmpRect, curColor
                
                '...followed by its border...
                If itemIsSelected Then
                    If itemIsHovered Then curColor = itemColorSelectedBorderHover Else curColor = itemColorSelectedBorder
                Else
                    If itemIsHovered Then curColor = itemColorUnselectedBorderHover Else curColor = itemColorUnselectedBorder
                End If
                GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, tmpRect, curColor, , , , LineJoinMiter
                
                '...and finally, its caption
                If itemIsSelected Then
                    If itemIsHovered Then curColor = fontColorSelectedHover Else curColor = fontColorSelected
                Else
                    If itemIsHovered Then curColor = fontColorUnselectedHover Else curColor = fontColorUnselected
                End If
                
                tmpFont.SetFontColor curColor
                tmpFont.AttachToDC bufferDC
                tmpFont.SetTextAlignment vbLeftJustify
                tmpFont.FastRenderTextWithClipping tmpRect.Left + textPadding, tmpRect.Top + LIST_PADDING_VERTICAL, tmpRect.Width - LIST_PADDING_HORIZONTAL, tmpRect.Height - LIST_PADDING_VERTICAL, tmpListItem.textTranslated, True, True
                tmpFont.ReleaseFromDC
                
            Next i
            
            Set tmpFont = Nothing
        
        End If
        
        'Last of all, we render the listbox border.  Note that we actually draw *two* borders.  The actual border,
        ' which is slightly inset from the list box boundaries, then a second border - pure white, erasing any item
        ' rendering that may have fallen outside the clipping area.
        Dim borderWidth As Single, borderColor As Long
        If listHasFocus Then borderWidth = 3# Else borderWidth = 1#
        borderColor = m_Colors.RetrieveColor(PDLB_Border, enabledState, listHasFocus)
        
        GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, m_ListRect, borderColor, , borderWidth, , LineJoinMiter
        
        If Not listHasFocus Then
            GDI_Plus.GDIPlusDrawRectOutlineToDC bufferDC, 0, 0, bWidth - 1, bHeight - 1, BackgroundColor, , , , LineJoinMiter
        End If
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint True
    
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
Public Sub UpdateAgainstCurrentTheme()
    UpdateColorList
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
