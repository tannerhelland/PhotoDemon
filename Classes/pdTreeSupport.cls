VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTreeSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Treeview support class
'Copyright 2024-2024 by Tanner Helland
'Created: 03/September/24
'Last updated: 03/September/24
'Last update: start initial build (split off from the list box support class)
'
'This class was originally created to enable a tree-view like control in PD's custom hotkey dialog.
'
'For a good understanding of how this class works, I recommend looking through the pdTreeView usercontrol.
' That UC uses this class to implement a generic, text-only treeview, and it does so with very little custom code
' (basically just enough to bubble up certain input events to this class, then rendering item text when this
' class raises a RedrawNeeded() event.)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control raises a few helper events, mostly related to rendering.
' If our owner does something that requires a redraw, we'll raise a "RedrawNeeded" event,
' which the owner can then respond to at their leisure.  (Or ignore, if e.g. they're not visible.)
'
'As far as rendering is concerned, please note that some events are always synonymous with a redraw
' (e.g. Click() always means the .ListIndex has changed, which in turn means a redraw is required).
' You *do not* have to render content from within the Click() event - you will always get a
' corresponding RedrawNeeded() event.
Public Event RedrawNeeded()
Public Event ScrollMaxChanged()
Public Event ScrollValueChanged()
Public Event Click()

'To simplify rendering, a custom struct tracks each tree item.
Private m_Items() As PD_TreeItem

'Current tree item count.  May or may not correspond to the actual *allocation size* of m_Items, FYI.
Private Const INITIAL_LIST_SIZE As Long = 16&
Private m_NumOfItems As Long

'Default height (in pixels) of a tree item.
' This is controlled by the owner, and we cannot do anything useful until its value is set.
Private m_DefaultHeight As Long

'Total height of the entire treeview, as it would appear without scrolling.
' This is used to calculate scrollbar values.
' (It must also be recalculated on collapse/expand events.)
Private m_TotalHeight As Long

'This class tracks .ListIndex .  It needs to know .ListIndex to auto-suggest things like
' scroll bar position after a keypress.
Private m_ListIndex As Long

'Scroll bar values.  This class doesn't do any actual rendering, but it tracks things like
' scroll value to make life easier on the owner (and because we have the data required for
' all those calculations anyway).
Private m_ScrollMax As Long, m_ScrollValue As Long
Private m_ListIndexHover As Long, m_MouseInsideList As Boolean
Private m_ListRectF As RectF
Private m_InitMouseX As Single, m_InitMouseY As Single, m_InitScroll As Single
Private m_LMBDown As Boolean

'Typically, adding an item to the treeview requires us to redraw.  This is a waste of time if
' the caller needs to add a bunch of items.  Instead of raising redraws automatically,
' the caller can choose to suspend redraws prior to adding items, then enable redraws after
' all items have been added.  (Please do this - it makes a big difference.)
Private m_RedrawAutomatically As Boolean

'Whenever a property changes that affects the on-screen appearance of the tree (e.g. adding an item,
' scrolling the list, expanding/collapsing something), we'll immediately cache the first and last
' elements that need to be drawn on-screen.  Then, when it comes time to render the list, we don't
' have to regenerate that info from scratch.
Private m_FirstRenderIndex As Long, m_LastRenderIndex As Long

'If the list is *not* in automatic redraw mode (e.g. redraw notifications are raised on every list change),
' we won't calculate rendering metrics as we go.  Instead, we'll just mark rendering metrics as dirty,
' and recalculate them when the owner finally requests rendering data.
Private m_RenderDataCorrect As Boolean

'When the tree is initialized, note the current language.  If the language changes on subsequent theme updates,
' we'll re-translate everything to match.
Private m_LanguageAtLastCheck As String

Private Sub Class_Initialize()
    
    m_DefaultHeight = 0
    m_RedrawAutomatically = True
    m_MouseInsideList = False
    m_ListIndexHover = -1
    
    Me.Clear
    
End Sub

'Add an item to the tree.  Unlike listviews, most params are mandatory (to enable hierarchical ordering).
Friend Sub AddItem(ByRef srcItemID As String, ByRef srcItemText As String, Optional ByRef parentID As String = vbNullString, Optional ByVal initialCollapsedState As Boolean = False, Optional ByVal itemShouldBeTranslated As Boolean = True)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Make sure there's room in the array for this item.
    If (m_NumOfItems > UBound(m_Items)) Then ReDim Preserve m_Items(0 To m_NumOfItems * 2 - 1) As PD_TreeItem
        
    'Unlike listviews, items must currently be added sequentially (for now; this may change in the future)
    Dim itemIndex As Long
    itemIndex = m_NumOfItems
    
    'Insert the given item
    With m_Items(itemIndex)
        .textEn = srcItemText
        
        'Add a translated copy of the string as well; this will be the string actually rendered onto the screen.
        If (Not g_Language Is Nothing) Then
            If g_Language.TranslationActive Then
                If itemShouldBeTranslated Then
                    .textTranslated = g_Language.TranslateMessage(srcItemText)
                Else
                    .textTranslated = srcItemText
                End If
            Else
                .textTranslated = srcItemText
            End If
        Else
            .textTranslated = srcItemText
        End If
        
        'For now, item height is fixed
        .itemHeight = m_DefaultHeight
        
        'Increase the net height of the entire list
        m_TotalHeight = m_TotalHeight + .itemHeight
        
        'Positioning values for a treeview work differently than a listview - I precalculate all item positions
        ' as if all list items were expanded.  This simplifies rendering because we can simply calculate a running
        ' negative offset representing collapsed items, and subtracting that from an item's .itemTop value gives
        ' us a precise rendering position.
        If (itemIndex > 0) Then
            .itemTop = m_Items(itemIndex - 1).itemTop + m_Items(itemIndex - 1).itemHeight
        Else
            .itemTop = 0
        End If
        
    End With
    
    'If this is the first item, note the current translation language.
    ' (If this changes, we need to re-translate the list.)
    If (m_NumOfItems = 0) Then
        If (Not g_Language Is Nothing) Then m_LanguageAtLastCheck = g_Language.GetCurrentLanguage
    End If
    
    'Increment the number of list entries
    m_NumOfItems = m_NumOfItems + 1
    
    'Redraw as necessary
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
End Sub

'Reset the current treeview.  An optional starting list size can be passed;
' if it is not passed, it will default to INITIAL_LIST_SIZE.
Friend Sub Clear(Optional ByVal newListSize As Long = INITIAL_LIST_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Failsafe bounds check
    If (newListSize <= 0) Then newListSize = INITIAL_LIST_SIZE
    
    'Reset the array
    ReDim m_Items(0 To newListSize - 1) As PD_TreeItem
    
    'Reset some obvious things (that don't require special handling)
    m_ListIndex = -1
    m_NumOfItems = 0
    m_TotalHeight = 0
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
    Exit Sub
    
FailsafeReset:
    If (newListSize <= 0) Then newListSize = INITIAL_LIST_SIZE
    ReDim m_Items(0 To newListSize - 1) As PD_TreeItem
    
End Sub

'Need to render a specific tree item?  Call this to retrieve a full copy of a given item's data,
' plus rendering-specific information like the item's literal position (in pixels) in the treeview.
Friend Function GetDirectTreeItem(ByVal srcListIndex As Long) As PD_TreeItem
    GetDirectTreeItem = m_Items(srcListIndex)
End Function

Friend Function GetHeightOfAllTreeItems() As Long
    GetHeightOfAllTreeItems = m_TotalHeight
End Function

Friend Function GetPointerToInternalListStruct() As Long
    GetPointerToInternalListStruct = VarPtr(m_Items(0))
End Function

Friend Function GetSizeOfInternalListStruct() As Long
    GetSizeOfInternalListStruct = UBound(m_Items)
End Function

'Font size controls the default height of each list item.  When the font size changes, we need to
' recalculate internal size metrics, so it's advisable to set this UP FRONT before doing anything else.
Friend Property Get DefaultItemHeight() As Single
    DefaultItemHeight = m_DefaultHeight
End Property

Friend Property Let DefaultItemHeight(ByVal newHeight As Single)
    If (m_DefaultHeight <> newHeight) Then
        m_DefaultHeight = newHeight
        
        'If a non-standard size mode is in use, we technically need to calculate new positioning metrics
        ' for *all* tree items.  This is stupid, and I'd prefer not to support it - so instead, just set
        ' the damn font size correctly *before* you add items!
        
    End If
End Property

Friend Function IsMouseInsideTreeView() As Boolean
    IsMouseInsideTreeView = m_MouseInsideList
End Function

'Retrieve a specified tree item - TODO pending how I actually use this class
Friend Function List(ByVal itemIndex As Long, Optional ByVal returnTranslatedText As Boolean = False) As String
    
    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        If returnTranslatedText Then
            List = m_Items(itemIndex).textTranslated
        Else
            List = m_Items(itemIndex).textEn
        End If
    Else
        List = vbNullString
    End If
    
End Function

Friend Function ListCount() As Long
    ListCount = m_NumOfItems
End Function

Friend Property Get ListIndex() As Long
    ListIndex = m_ListIndex
End Property

'TODO: see if this remains relevant
Friend Property Let ListIndex(ByVal newIndex As Long)
    
    If (newIndex < m_NumOfItems) And PDMain.IsProgramRunning() Then
        
        m_ListIndex = newIndex
        If (newIndex >= 0) Then
        
            RaiseEvent Click
        
            'Changing the list index may require us to shift the scrollbar value, so that the newly selected item fits on-screen.
            If MakeSureListIndexFitsOnscreen Then
                RaiseEvent ScrollValueChanged
                If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
            Else
                If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
            End If
            
        Else
            If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
        End If
        
    End If
    
End Property

'TODO: remove as necessary
'As a convenience, this function lets the caller retrieve a given ListIndex by associated string contents.
' IMPORTANT NOTE: this function doesn't *change* the ListIndex; it only returns a hypothetical one matching the input string.
Friend Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) As Long
    
    ListIndexByString = -1
    
    If (m_NumOfItems > 0) Then
        
        Dim newIndex As Long
        newIndex = -1
        
        Dim i As Long
        For i = 0 To m_NumOfItems - 1
            If Strings.StringsEqual(srcString, m_Items(i).textEn, (compareMode = vbTextCompare)) Then
                newIndex = i
                Exit For
            End If
        Next i
        
        If (newIndex >= 0) Then ListIndexByString = newIndex
        
    End If
    
End Function

'TODO: remove as necessary
'As a convenience, this function lets the caller retrieve a given ListIndex by mouse position within the container.
' IMPORTANT NOTE: this function doesn't *change* the ListIndex; it only returns a hypothetical one from position (x, y).
Friend Function ListIndexByPosition(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal checkXAsWell As Boolean = True) As Long
    
    ListIndexByPosition = -1
    
    'First, do a spot-check on srcX.  If it lies outside the list region, skip this whole step.
    If checkXAsWell Then
        If (srcX < m_ListRectF.Left) Or (srcX > m_ListRectF.Left + m_ListRectF.Width) Then
            ListIndexByPosition = -1
            Exit Function
        End If
    End If
    
    'Convert the y-position to an absolute value
    srcY = (srcY - m_ListRectF.Top) + m_ScrollValue
    
    'TODO: handle collapsed items, obviously
    
    'On a fixed-size list, this calculation can be automated.
    ListIndexByPosition = srcY \ m_DefaultHeight
    If (ListIndexByPosition >= m_NumOfItems) Then ListIndexByPosition = -1
        
    
End Function

'This property exists purely for improving UI feedback.  It may or may not be identical to the default .ListIndex value.
Friend Property Get ListIndexHovered() As Long
    ListIndexHovered = m_ListIndexHover
End Property

'While this class doesn't actually render anything (that's left to the parent), it has enough information
' to manage a lot of the backend rendering details automatically.  This makes it much easier to construct
' custom treeviews, as things like hit detection can be handled here.
Friend Sub NotifyParentRectF(ByRef srcListRectF As RectF)
    If (m_ListRectF.Width <> srcListRectF.Width) Or (m_ListRectF.Height <> srcListRectF.Height) Or (m_ListRectF.Top <> srcListRectF.Top) Or (m_ListRectF.Left <> srcListRectF.Left) Then
        m_ListRectF = srcListRectF
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    End If
End Sub

'Key events primarily affect the current .ListIndex property.  Unlike listviews, keys can also expand/contract
' items which changes all rendering data, so a lot of this code triggers full-list redraws.
Friend Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    'Keys can set a new listindex.  We'll calculate a theoretical new ListIndex, then apply a universal
    ' bounds-check at the end.
    Dim newListIndex As Long
    newListIndex = m_ListIndex
    
    Select Case vkCode
        
        Case VK_DOWN
            newListIndex = newListIndex + 1
            markEventHandled = True
            
        Case VK_UP
            newListIndex = newListIndex - 1
            markEventHandled = True
            
        Case VK_PAGEDOWN
            'TODO: the math for this obviously changes depending on expand/contract state of neighboring items.
            ' We can still calculate how *many* items we need to "jump", but then we'll need to manually iterate
            ' that many *expanded* items.
            newListIndex = newListIndex + (m_LastRenderIndex - m_FirstRenderIndex)
            markEventHandled = True
            
        Case VK_PAGEUP
            newListIndex = newListIndex - (m_LastRenderIndex - m_FirstRenderIndex)
            markEventHandled = True
            
        Case VK_HOME
            newListIndex = 0
            markEventHandled = True
            
        Case VK_END
            newListIndex = m_NumOfItems - 1
            markEventHandled = True
        
        'TODO: hotkeys for expand/contract - check how default treeviews work?
        
    End Select
    
    If (m_NumOfItems = 0) Then
        newListIndex = -1
    Else
        If (newListIndex < 0) Then newListIndex = 0
        If (newListIndex > m_NumOfItems - 1) Then newListIndex = m_NumOfItems - 1
    End If
    
    Me.ListIndex = newListIndex
    
End Sub

'Not used at present?
Friend Sub NotifyKeyUp(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

End Sub

'Mouse events control a whole bunch of possible things: hover state, .ListIndex, scroll.  As such, their handling is
' fairly involved, despite this class not doing any actual UI rendering.
'
'TODO: does caller or this class handle expand/contract button hit detection?
Friend Sub NotifyMouseClick(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If (Button And pdLeftButton) <> 0 Then
        Dim tmpListIndex As Long
        tmpListIndex = Me.ListIndexByPosition(x, y)
        If tmpListIndex >= 0 Then Me.ListIndex = tmpListIndex
    End If
End Sub

Friend Sub NotifyMouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateHoveredIndex x, y, True
End Sub

Friend Sub NotifyMouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateHoveredIndex -100, -100, True
End Sub

Friend Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_InitMouseX = x
    m_InitMouseY = y
    m_InitScroll = m_ScrollValue
    If ((Button And pdLeftButton) <> 0) Then m_LMBDown = True
End Sub

Friend Sub NotifyMouseMove(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If m_LMBDown Then
        ScrollListByDrag x, y
    Else
        UpdateHoveredIndex x, y
    End If
End Sub

Friend Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean)
    If ((Button And pdLeftButton) <> 0) And m_LMBDown Then
        m_LMBDown = False
        If (Not clickEventAlsoFiring) Then ScrollListByDrag x, y
    End If
End Sub

Friend Sub NotifyMouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    If (m_ScrollMax > 0) And (scrollAmount <> 0) Then
    
        Dim newScrollValue As Long: newScrollValue = m_ScrollValue
        
        If (scrollAmount > 0) Then
            newScrollValue = newScrollValue - Me.DefaultItemHeight
        Else
            newScrollValue = newScrollValue + Me.DefaultItemHeight
        End If
        
        If (newScrollValue < 0) Then newScrollValue = 0
        If (newScrollValue > m_ScrollMax) Then newScrollValue = m_ScrollMax
        Me.ScrollValue = newScrollValue
        
    End If
    
End Sub

'TODO: make this function collapse/expand-state friendly
'When a new ListIndex value is set, PD makes sure that item appears on-screen.
' This function will return TRUE if the current scroll value was changed to bring the item on-screen.
Private Function MakeSureListIndexFitsOnscreen() As Boolean

    MakeSureListIndexFitsOnscreen = False
    
    'If the item's current top and bottom values fit with the listview's container area, we don't need to change anything.
    If (m_ListIndex >= 0) And (m_ListIndex < m_NumOfItems) Then
        
        Dim liTop As Single, liBottom As Single
        liTop = m_Items(m_ListIndex).itemTop
        liBottom = liTop + m_Items(m_ListIndex).itemHeight
        
        Dim liContainerTop As Single, liContainerBottom As Single
        liContainerTop = m_ScrollValue
        liContainerBottom = liContainerTop + m_ListRectF.Height
        
        'If either the top or bottom of the item rect lies outside the container rect, calculate new scroll values
        If (liTop < liContainerTop) Or (liBottom > liContainerBottom) Then
        
            MakeSureListIndexFitsOnscreen = True
            
            'If the item lies *above* the viewport rect, scroll down to compensate.
            If (liTop < liContainerTop) Then
                m_ScrollValue = liTop
                
            'If the item lies *below* the viewport rect, scroll up to compensate.
            Else
                m_ScrollValue = (liTop - (m_ListRectF.Height - m_Items(m_ListIndex).itemHeight)) - 1
            End If
            
            If (m_ScrollValue > m_ScrollMax) Then m_ScrollValue = m_ScrollMax
            If (m_ScrollValue < 0) Then m_ScrollValue = 0
        
        End If
        
    End If

End Function

'All PD lists aim to support scroll-by-drag behavior
Private Sub ScrollListByDrag(ByVal newX As Single, ByVal newY As Single)
    If (m_ScrollMax > 0) Then
        Dim tmpScrollValue As Long
        tmpScrollValue = m_InitScroll + (m_InitMouseY - newY)
        If (tmpScrollValue < 0) Then tmpScrollValue = 0
        If (tmpScrollValue > m_ScrollMax) Then tmpScrollValue = m_ScrollMax
        Me.ScrollValue = tmpScrollValue
    End If
End Sub

'When the mouse enters, leaves, or moves inside the underlying control, we will update the current hover index
' (provided our parent relays those events to us, obviously).
Private Sub UpdateHoveredIndex(ByVal x As Long, ByVal y As Long, Optional ByVal ensureRedrawEvent As Boolean = False)
    
    Dim oldMouseInsideList As Boolean
    oldMouseInsideList = m_MouseInsideList
    m_MouseInsideList = PDMath.IsPointInRectF(x, y, m_ListRectF)
    
    Dim tmpListIndex As Long
    If m_MouseInsideList Then tmpListIndex = Me.ListIndexByPosition(x, y) Else tmpListIndex = -1
    
    If (tmpListIndex <> m_ListIndexHover) Or (oldMouseInsideList <> m_MouseInsideList) Or ensureRedrawEvent Then
        m_ListIndexHover = tmpListIndex
        If m_RedrawAutomatically Or ensureRedrawEvent Then
            RaiseEvent RedrawNeeded
        Else
            m_RenderDataCorrect = False
        End If
    End If
    
End Sub

'Remove an item.  (TODO: re-implement this?  Do I even need it?)
Friend Sub RemoveItem(ByVal itemIndex As Long)
    
    'First, make sure the requested index is valid
    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        
        'Remove this item's size from the net height tracker
        Dim missingItemHeight As Long
        missingItemHeight = m_Items(itemIndex).itemHeight
        m_TotalHeight = m_TotalHeight - missingItemHeight
        
        'If this item is not being removed from the *end* of the list, shift everything past it downward.
        Dim i As Long
        If itemIndex < (m_NumOfItems - 1) Then
            For i = itemIndex To m_NumOfItems - 2
                m_Items(i) = m_Items(i + 1)
                m_Items(i).itemTop = m_Items(i).itemTop - missingItemHeight
            Next i
        End If
        
        'Reduce the total list size
        m_NumOfItems = m_NumOfItems - 1
        
        'If the removal affected the current ListIndex, update it to match
        If (itemIndex <= m_ListIndex) Then
            m_ListIndex = m_ListIndex - 1
            If (m_ListIndex < 0) Then m_ListIndex = 0
            RaiseEvent Click
        End If
        
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
        
    End If
    
End Sub

'Need to render the tree?  Call this first to get rendering limits.
Friend Sub GetRenderingLimits(ByRef firstRenderIndex As Long, ByRef lastRenderIndex As Long, ByRef listIsEmpty As Boolean)
    firstRenderIndex = m_FirstRenderIndex
    lastRenderIndex = m_LastRenderIndex
    listIsEmpty = (m_NumOfItems = 0)
End Sub

'Need to render a specific item?  Call this to retrieve a full copy of a given item's data,
' plus rendering-specific information like the item's literal position in the current view.
Friend Sub GetRenderingItem(ByVal srcListIndex As Long, ByRef dstListItem As PD_TreeItem, ByRef dstItemTop As Long, ByRef dstItemHeight As Long, ByRef dstItemHeightWithoutSeparator As Long)
    
    dstListItem = m_Items(srcListIndex)
    
    'TODO: handle collapse/expand state
    dstItemTop = (srcListIndex * m_DefaultHeight) - m_ScrollValue
    dstItemHeight = m_DefaultHeight
    dstItemHeightWithoutSeparator = m_DefaultHeight
    
End Sub

'While this class doesn't do any actual rendering, it does calculate all relevant scroll bar and positioning values.
' This makes life easier on the caller.

'Obviously, these values may not be correct if the class has not been notified of the list box size and/or not all of
' its contents have been loaded yet.
Friend Property Get ScrollMax() As Long
    If m_RenderDataCorrect Then
        ScrollMax = m_ScrollMax
    Else
        If m_RedrawAutomatically Then
            CalculateRenderMetrics
        Else
            m_ScrollMax = m_TotalHeight - m_ListRectF.Height
            If (m_ScrollMax < 0) Then m_ScrollMax = 0
        End If
    End If
End Property

Friend Property Get ScrollValue() As Long
    ScrollValue = m_ScrollValue
End Property

'When assigning a new scroll value, you should probably double-check the passed newValue.
' This class will automatically reset the value to an appropriate range if it's too small or too large.
Friend Property Let ScrollValue(ByRef newValue As Long)
    
    'Range-check the incoming value
    If (newValue < 0) Then newValue = 0
    If (newValue > m_ScrollMax) Then newValue = m_ScrollMax
    m_ScrollValue = newValue
    
    'Changing the scroll value changes the on-screen position of list elements, so we need to recalculate rendering data.
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    RaiseEvent ScrollValueChanged
    
End Property

Private Sub CalculateNewScrollMax()
    
    Dim newScrollMax As Long
    newScrollMax = m_TotalHeight - m_ListRectF.Height
    If (newScrollMax < 0) Then newScrollMax = 0
    
    If (newScrollMax <> m_ScrollMax) Then
        m_ScrollMax = newScrollMax
        RaiseEvent ScrollMaxChanged
    End If
    
End Sub

'The caller can suspend automatic redraws caused by things like adding an item to the list box.
' Just make sure to enable redraws once you're ready, or you'll never get rendering requests!
Friend Sub SetAutomaticRedraws(ByVal newState As Boolean, Optional ByVal raiseRedrawImmediately As Boolean = False)
    
    m_RedrawAutomatically = newState
    
    If raiseRedrawImmediately Then
        If m_RenderDataCorrect Then
            RaiseEvent RedrawNeeded
        Else
            Me.CalculateRenderMetrics
        End If
    End If
    
End Sub

'Call this sub to request a full redraw of the list.  This sub doesn't actually perform any drawing;
' instead, it raises a series of RenderListItem() events, which the caller can then handle on their own terms.
Friend Sub CalculateRenderMetrics()
    
    'Prior to requesting a redraw, determine the indices of the first and last items our owner needs to draw.
    ' We'll cache these, so we don't have to calculate them again (until something changes, at least).
    Dim i As Long
    
    If (m_ListRectF.Height <= 0) Or (m_DefaultHeight <= 0) Then Exit Sub
    
    'TODO: handle expand/contract state
    
    'Calculate the first overlapping item that overlaps the viewable area (and cache it)
    i = m_ScrollValue \ m_DefaultHeight
    m_FirstRenderIndex = i
    
    'Search for the first item that doesn't overlap the existing container area
    Do
        
        'Move to the next entry in the list
        i = i + 1
        
        'If we're past the number of items in the list, exit immediately
        If (i >= m_NumOfItems) Then Exit Do
        
    Loop While (i * m_DefaultHeight) < (m_ScrollValue + m_ListRectF.Height)
    
    m_LastRenderIndex = i - 1
    
    'Note that our rendering data is up-to-date.  As long as this stays TRUE, we don't have to recalculate rendering data.
    m_RenderDataCorrect = True
    
    'Whenever list contents or container sizes change, we can cache a new scroll bar maximum value.
    CalculateNewScrollMax
    
    If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
    
End Sub

Friend Sub UpdateAgainstCurrentTheme()
    
    If (Not g_Language Is Nothing) Then
        
        If (m_NumOfItems > 0) And (m_LanguageAtLastCheck <> g_Language.GetCurrentLanguage()) Then
            
            m_LanguageAtLastCheck = g_Language.GetCurrentLanguage
            
            Dim i As Long
            If g_Language.TranslationActive Then
                For i = 0 To m_NumOfItems - 1
                    m_Items(i).textTranslated = g_Language.TranslateMessage(m_Items(i).textEn)
                Next i
            Else
                For i = 0 To m_NumOfItems - 1
                    m_Items(i).textTranslated = m_Items(i).textEn
                Next i
            End If
        
            m_RenderDataCorrect = False
            Me.CalculateRenderMetrics
            
        End If
        
    End If

End Sub
