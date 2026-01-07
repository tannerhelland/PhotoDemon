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
'Copyright 2024-2026 by Tanner Helland
'Created: 03/September/24
'Last updated: 23/September/24
'Last update: finish adding all relevant keyboard shortcut features
'
'This class was originally created to enable a tree-view like control in PD's custom hotkey dialog.  It manages
' all tree-related function, but it doesn't do any rendering (for that, refer to the pdTreeView usercontrol).
'
'I've tried to mimic standard Windows treeview behavior as closely as possible.  Please let me know if you
' find features or behaviors that I've missed.
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

'Default height (in pixels) of a single tree item.
' This is controlled by the owner, and we cannot do anything useful until its value is set.
Private m_DefaultHeight As Long

'Total height of the entire treeview, as it would appear without scrolling.
' This is used to calculate scrollbar values.
Private m_TotalHeight As Long

'How many horizontal pixels to offset an item by per level of tree.
' Auto-calculated based on m_DefaultHeight.
Private m_OffsetPerLevel As Long

'Despite being a treeview (where items can be collapsed/hidden), this class still tracks a .ListIndex value.
' This is currently more intuitive for me to track and respond to in PD due to the way tree items are
' managed in the background.
Private m_ListIndex As Long

'Scroll bar values.  This class doesn't do any actual rendering, but it tracks things like
' scroll value and max to make life easier on the owner (and because we have the data required
' for all those calculations).
Private m_ScrollMax As Long, m_ScrollValue As Long
Private m_ListIndexHover As Long, m_MouseInsideList As Boolean
Private m_ListRectF As RectF
Private m_InitMouseX As Single, m_InitMouseY As Single, m_InitScroll As Single
Private m_LMBDown As Boolean

'Typically, adding an item to the treeview requires us to redraw.  This is a waste of time if
' the caller needs to add a bunch of items.  Instead of raising redraws automatically,
' the caller can choose to suspend redraws prior to adding items, then enable redraws after
' all items have been added.  (Please use this capability - it makes a big difference.)
Private m_RedrawAutomatically As Boolean

'Whenever a property changes that affects the on-screen appearance of the tree (e.g. adding an item,
' scrolling the list, expanding/collapsing something), we'll immediately cache the first and last
' elements that need to be drawn on-screen.  Then, when it comes time to render the list, we don't
' have to regenerate that info from scratch.
'
'Because of the way treeviews work, however, note that indices may be included in this range that
' are *not* rendered on-screen (because they are collapsed).  The caller still needs to check
' collapse state of each item prior to rendering!
Private m_FirstRenderIndex As Long, m_LastRenderIndex As Long

'If the list is *not* in automatic redraw mode (e.g. redraw notifications are raised on every list change),
' we don't calculate rendering metrics as we go.  Instead, we'll just mark rendering metrics as dirty,
' and recalculate them when the owner finally requests rendering data.
Private m_RenderDataCorrect As Boolean

'When the tree is initialized, note the current language.  If the language changes on subsequent theme updates,
' we'll raise redraw notifications so that the caller has a chance to re-translate anything localized.
Private m_LanguageAtLastCheck As String

'Treeviews require a lot of mapping back-and-forth between parents and children.  (When a node is closed,
' we have to find all children of a particular parent.  Before rendering, we also need to count the number
' of parent nodes in order to know how far left/right to render a particular item.)  A hash table speeds
' up that mapping by allowing us to quickly correlate IDs and indices.
Private m_ItemHash As pdVariantHash

Private Sub Class_Initialize()
    
    m_DefaultHeight = 0
    m_RedrawAutomatically = True
    m_MouseInsideList = False
    m_ListIndexHover = -1
    Set m_ItemHash = New pdVariantHash
    
    Me.Clear
    
End Sub

'Add an item to the tree.  Unlike listviews, most params are mandatory (to enable hierarchical ordering).
Friend Sub AddItem(ByRef srcItemID As String, ByRef srcItemText As String, Optional ByRef parentID As String = vbNullString, Optional ByVal initialCollapsedState As Boolean = False)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Make sure there's room in the array for this item.
    If (m_NumOfItems > UBound(m_Items)) Then ReDim Preserve m_Items(0 To m_NumOfItems * 2 - 1) As PD_TreeItem
        
    'Unlike listviews, items must currently be added sequentially.
    Dim itemIndex As Long
    itemIndex = m_NumOfItems
    
    'Insert the given item
    With m_Items(itemIndex)
        
        .itemID = srcItemID
        .parentID = parentID
        
        'We only care about collapsed state for parent nodes; child collapsed state is *not* tracked on this
        ' property, but note that we'll perform a second pass through the data to clean this field up (in case
        ' the user passes it wrongly).
        .isCollapsed = initialCollapsedState
        
        'If this node has a parent, find that parent in our lookup table and note it as having children.
        If (LenB(parentID) <> 0) Then
            Dim idxParent As Variant
            If m_ItemHash.GetItemByKey(parentID, idxParent) Then m_Items(idxParent).hasChildren = True
        End If
        
        'Reset all "this-render" collapsed states to FALSE; we'll re-calculate this on first-render
        ' (since hierarchical rendering decisions need to be made for grandchildren, great-grandchildren, etc)
        .isCollapsedThisRender = False
        
        'For an owner-drawn control like this, translations must be handled manually by the caller,
        ' but we do track en-US text so we can return it when useful.
        .textEn = srcItemText
        
    End With
    
    'Add this item to our running hash table (it enables fast index/ID matching)
    m_ItemHash.AddItem srcItemID, itemIndex
    
    'If this is the first item, note the current translation language.
    ' (If this changes, we need to notify the caller so they can re-translate the list as necessary.)
    If (m_NumOfItems = 0) Then
        If (Not g_Language Is Nothing) Then m_LanguageAtLastCheck = g_Language.GetCurrentLanguage
    End If
    
    'Increment the number of list entries, then redraw as necessary
    m_NumOfItems = m_NumOfItems + 1
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
    Set m_ItemHash = New pdVariantHash
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
    Exit Sub
    
FailsafeReset:
    If (newListSize <= 0) Then newListSize = INITIAL_LIST_SIZE
    ReDim m_Items(0 To newListSize - 1) As PD_TreeItem
    
End Sub

'Font size controls the default height of each list item.  When the font size changes, we need to
' recalculate internal size metrics, so it's advisable to set this UP FRONT before doing anything else.
Friend Property Get DefaultItemHeight() As Long
    DefaultItemHeight = m_DefaultHeight
End Property

Friend Property Let DefaultItemHeight(ByVal newHeight As Long)
    If (m_DefaultHeight <> newHeight) Then
        m_DefaultHeight = newHeight
        
        'If a non-standard size mode is in use, we technically need to calculate new positioning metrics
        ' for *all* tree items.  This is stupid, and I'd prefer not to support it - so instead, just set
        ' the damn font size correctly *before* you add items!
        
        'While here, we also calculate offset-per-level (in horizontal pixels) as a ratio of each
        ' item's height.
        m_OffsetPerLevel = m_DefaultHeight
        
    End If
End Property

Friend Function IsMouseInsideTreeView() As Boolean
    IsMouseInsideTreeView = m_MouseInsideList
End Function

'Retrieve a specified tree item (this is currently unused in PD)
Friend Function List(ByVal itemIndex As Long) As String
    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        List = m_Items(itemIndex).textEn
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

'This function does not currently attempt to validate that the target index is actually visible.
' If I someday decide to use this function in PD, that needs to be addressed!
Friend Property Let ListIndex(ByVal newIndex As Long)
    
    If (newIndex < m_NumOfItems) And PDMain.IsProgramRunning() Then
        
        m_ListIndex = newIndex
        If (newIndex >= 0) Then
        
            RaiseEvent Click
        
            'Changing the list index may require us to shift the scrollbar value, so that the newly selected item fits on-screen.
            m_ListIndex = ValidateIndexIsVisible(newIndex, True)
            
            If MakeSureListIndexFitsOnscreen() Then
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

'As a convenience, this function lets the caller retrieve a given ListIndex by mouse position within the container.
' (Various internal functions use this for mouse-hit behavior.)
'
'IMPORTANT NOTE: this function doesn't *change* the ListIndex; it only returns a hypothetical one from position (x, y).
Friend Function ListIndexByPosition(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal checkXAsWell As Boolean = True) As Long
    
    ListIndexByPosition = -1
    
    'First, do a spot-check on srcX.  If it lies outside the list region, skip this whole step.
    If checkXAsWell Then
        If (srcX < m_ListRectF.Left) Or (srcX > m_ListRectF.Left + m_ListRectF.Width) Then
            ListIndexByPosition = -1
            Exit Function
        End If
    End If
    
    'Convert the y-position to an absolute value (accounting for scroll)
    srcY = (srcY - m_ListRectF.Top) + m_ScrollValue
    
    'Iterate the list until we reach the item where the click occurred.  Note that collapsed items
    ' *must* be skipped!
    Dim i As Long
    For i = m_FirstRenderIndex To m_LastRenderIndex
        
        If (i >= 0) And (i < m_NumOfItems) Then
            If (Not m_Items(i).isCollapsedThisRender) Then
                If PDMath.IsPointInRectF(srcX, srcY, m_Items(i).ItemRect) Then
                    ListIndexByPosition = i
                    Exit For
                End If
            End If
        End If
        
    Next i
    
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
' items which changes all rendering data, so a lot of this code triggers full-tree redraws.
Friend Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    'Keys can set a new listindex.  We'll calculate a theoretical new ListIndex, then apply a universal
    ' bounds-check at the end.
    Dim newListIndex As Long
    newListIndex = m_ListIndex
    
    'By default, this function will validate the newly set list index after any keypresses.
    ' (If a keypress doesn't modify the listindex, it should set this to FALSE.)
    Dim validateListIndex As Boolean
    validateListIndex = True
    
    Select Case vkCode
        
        Case VK_DOWN
            
            'Ctrl+Shift means scroll down *without* changing listindex
            If ((Shift And vbCtrlMask) = vbCtrlMask) Then
                Me.ScrollValue = Me.ScrollValue + m_DefaultHeight
                validateListIndex = False
            
            'No modifiers just means move to the next visible item
            Else
                If (newListIndex < 0) Then
                    newListIndex = 0
                Else
                    newListIndex = AdvanceListIndex(True, False)
                End If
            End If
            markEventHandled = True
        
        'See comments for VK_DOWN; this is the same, just in the opposite direction
        Case VK_UP
            If ((Shift And vbCtrlMask) = vbCtrlMask) Then
                Me.ScrollValue = Me.ScrollValue - m_DefaultHeight
                validateListIndex = False
            Else
                If (newListIndex < 0) Then
                    newListIndex = 0
                Else
                    newListIndex = AdvanceListIndex(False, False)
                End If
            End If
            markEventHandled = True
        
        'Advance a page of list items
        Case VK_PAGEDOWN
            newListIndex = AdvanceListIndex(True, True)
            markEventHandled = True
            
        Case VK_PAGEUP
            newListIndex = AdvanceListIndex(False, True)
            markEventHandled = True
        
        'Top, bottom of list
        Case VK_HOME
            newListIndex = ValidateIndexIsVisible(0, False)
            markEventHandled = True
            
        Case VK_END
            newListIndex = ValidateIndexIsVisible(m_NumOfItems - 1, True)
            markEventHandled = True
        
        'Expand this node (if you can)
        Case VK_RIGHT
            If (m_ListIndex >= 0) Then
                If m_Items(m_ListIndex).hasChildren Then
                    If m_Items(m_ListIndex).isCollapsed Then
                        ToggleCollapseStateByIndex m_ListIndex, True, False
                        newListIndex = ValidateIndexIsVisible(m_ListIndex, True)
                    
                    'If this list is already expanded, move to the next child in line.
                    ' (This mirrors behavior in Windows Explorer.)
                    Else
                        newListIndex = AdvanceListIndex(True, False)
                    End If
                
                'Technically, Microsoft does *nothing* if pressing the right-arrow on a node with no children.
                ' However, I think it's more intuitive to move *down* the list rather than do nothing.
                ' (This allows the user to continue moving through the list, expanding nodes as they go.)
                Else
                    newListIndex = AdvanceListIndex(True, False)
                End If
            
            '/m_ListIndex < 0
            Else
                newListIndex = 0
            End If
            
            newListIndex = ValidateIndexIsVisible(newListIndex, True)
            markEventHandled = True
    
        'Collapse the list (if you can)
        Case VK_LEFT
            If (m_ListIndex >= 0) Then
                
                'If this node has children, close it...
                If m_Items(m_ListIndex).hasChildren Then
                    
                    '...unless the node is already closed, in which case jump to the previous parent in the chain
                    If m_Items(m_ListIndex).isCollapsed Then
                        newListIndex = JumpToNearestParent(m_ListIndex)
                    Else
                        ToggleCollapseStateByIndex m_ListIndex, True, True
                        newListIndex = m_ListIndex
                    End If
                Else
                    newListIndex = JumpToNearestParent(m_ListIndex)
                End If
                
            End If
            
            newListIndex = ValidateIndexIsVisible(newListIndex, True)
            markEventHandled = True
        
        '+/- expand/contract the current node
        Case VK_ADD, VK_OEM_PLUS
            If (m_ListIndex >= 0) Then
                If (m_Items(m_ListIndex).hasChildren And m_Items(m_ListIndex).isCollapsed) Then
                    ToggleCollapseStateByIndex m_ListIndex
                End If
            End If
            markEventHandled = True
            
        Case VK_SUBTRACT, VK_OEM_MINUS
            If (m_ListIndex >= 0) Then
                If (m_Items(m_ListIndex).hasChildren And Not m_Items(m_ListIndex).isCollapsed) Then
                    ToggleCollapseStateByIndex m_ListIndex
                End If
            End If
            markEventHandled = True
            
        'The asterisk key (numpad) expands all expandable nodes beneath the current one
        Case VK_MULTIPLY
            
            'For this operation, we want to iterate sequential nodes, and any node that lists this node
            ' as an ancestor gets expanded.
            ExpandAllChildrenOfNode m_ListIndex
            
    End Select
    
    'Any function that modifies the list index requires us to ensure the new listindex is visible on-screen.
    If validateListIndex Then
        
        If (m_NumOfItems = 0) Then
            newListIndex = -1
        Else
            If (newListIndex < 0) Then newListIndex = ValidateIndexIsVisible(0, False)
            If (newListIndex > m_NumOfItems - 1) Then newListIndex = ValidateIndexIsVisible(m_NumOfItems - 1, True)
        End If
        
        Me.ListIndex = newListIndex
        
    End If
        
End Sub

'On numpad asterisk (multiply) keypresses, we need to expand all child nodes of the currently selected node.
Private Sub ExpandAllChildrenOfNode(ByVal idxTarget As Long)
    
    'We will use a hash table for fast parent lookups
    Dim cParents As pdVariantHash
    Set cParents = New pdVariantHash
    
    'Add the target node to the hash table, then open it
    cParents.AddItem m_Items(idxTarget).itemID, idxTarget
    m_Items(idxTarget).isCollapsed = False
    
    'Now, iterate child nodes until we reach one that isn't part of this node's descendants tree.
    Dim i As Long
    For i = idxTarget + 1 To m_NumOfItems - 1
        
        'Start with direct children
        If (m_Items(i).parentID = m_Items(idxTarget).itemID) Then
            
            'If this node has children, add this node to the "valid parent" hash table.
            If m_Items(i).hasChildren Then
                cParents.AddItem m_Items(i).itemID, i
                m_Items(i).isCollapsed = False
            
            '/no Else required (if this node doesn't have children, it'll be expanded automatically).
            End If
        
        'If this node *isn't* a child of the target node, it must be a grandchild (or great-grandchild, etc)
        ' to be expanded; otherwise, we can stop iterating nodes.
        Else
            
            Dim idxSearch As Long, idxNew As Variant
            idxSearch = i
            
            Dim isGrandchild As Boolean: isGrandchild = False
            Do While cParents.GetItemByKey(m_Items(idxSearch).parentID, idxNew)
                
                idxSearch = idxNew
                If m_Items(idxSearch).parentID = m_Items(idxTarget).itemID Then
                
                    'This node is a descendant of the target node.  If it has children, add it to the
                    ' "valid parents" collection and expand it.
                    If m_Items(idxSearch).hasChildren Then
                        cParents.AddItem m_Items(idxSearch).itemID, idxSearch
                        m_Items(idxSearch).isCollapsed = False
                    End If
                    
                    Exit Do
                    
                End If
                
            Loop
            
        End If
        
    Next i
    
    'After all that, redraw the tree
    CalculateRenderMetrics
    
End Sub

'Move the list index around in response to keyboard events.  Returns the new .ListIndex, if any.
Private Function AdvanceListIndex(ByVal directionDown As Boolean, ByVal jumpWholePage As Boolean) As Long
    
    'Start with the current listindex, if any
    AdvanceListIndex = Me.ListIndex
    
    Dim i As Long, numItemsInPage As Long, numItemsPassed As Long
    
    'Figure out how many items are visible at once in the listbox.  (PgUp/Down jumps a whole "page" of entries.)
    numItemsInPage = (m_ListRectF.Height \ m_DefaultHeight)
    
    If directionDown Then
        
        'PgDown
        If jumpWholePage Then
            
            numItemsPassed = 0
            For i = AdvanceListIndex + 1 To m_NumOfItems - 1
                If (Not m_Items(i).isCollapsedThisRender) Then numItemsPassed = numItemsPassed + 1
                If (numItemsPassed >= numItemsInPage) Then Exit For
            Next i
            
            AdvanceListIndex = PDMath.Min2Int(m_NumOfItems - 1, i)
            
        'Down
        Else
            
            For i = AdvanceListIndex + 1 To m_NumOfItems - 1
                If (Not m_Items(i).isCollapsedThisRender) Then Exit For
            Next i
            
            AdvanceListIndex = PDMath.Min2Int(m_NumOfItems - 1, i)
            
        End If
        
    'directionUp
    Else
    
        'PgUp
        If jumpWholePage Then
        
            numItemsPassed = 0
            For i = AdvanceListIndex - 1 To 0 Step -1
                If (Not m_Items(i).isCollapsedThisRender) Then numItemsPassed = numItemsPassed + 1
                If (numItemsPassed >= numItemsInPage) Then Exit For
            Next i
            
            AdvanceListIndex = PDMath.Max2Int(0, i)
            
        'Up
        Else
        
            For i = AdvanceListIndex - 1 To 0 Step -1
                If (Not m_Items(i).isCollapsedThisRender) Then Exit For
            Next i
            
            AdvanceListIndex = PDMath.Max2Int(0, i)
            
        End If
    
    End If
    
    AdvanceListIndex = ValidateIndexIsVisible(AdvanceListIndex, directionDown)
    
End Function

'Validate that a given listindex is visible in the current tree.  If it isn't visible, supply a direction
' and this function will find the next-closest visible listindex possibility.
Private Function ValidateIndexIsVisible(ByVal idxTarget As Long, ByVal directionDown As Boolean) As Long
    
    'Do nothing for out-of-bound indices
    If (idxTarget < 0) Or (idxTarget >= m_NumOfItems) Then
        ValidateIndexIsVisible = idxTarget
        Exit Function
    End If
    
    'Before exiting, make sure our selected item is *not* collapsed.
    ' (This can happen if e.g. the user is on the last-visible item, and it's a collapsed parent item.)
    If m_Items(idxTarget).isCollapsedThisRender Then
        
        Dim i As Long
        
        If directionDown Then
            
            'Retreat backwards until we find a non-collapsed item.
            For i = idxTarget To 0 Step -1
                If (Not m_Items(i).isCollapsedThisRender) Then Exit For
            Next i
            
            If (i < 0) Then
                idxTarget = -1
            Else
                idxTarget = i
            End If
            
        '/directionUp
        Else
        
            'Move forward until we find a non-collapsed item.
            For i = idxTarget To m_NumOfItems - 1
                If (Not m_Items(i).isCollapsedThisRender) Then Exit For
            Next i
            
            If (i < m_NumOfItems - 1) Then
                idxTarget = i
            Else
                idxTarget = -1
            End If
            
        End If
    
    'end "item is collapsed"
    End If
    
    ValidateIndexIsVisible = idxTarget
    
End Function

Friend Sub NotifyKeyUp(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    'Not used at present
End Sub

'Mouse events control a whole bunch of possible things: hover state, .ListIndex, scroll.  As such, their handling is
' fairly involved, despite this class not doing any actual UI rendering.
Friend Sub NotifyMouseClick(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If (Button And pdLeftButton) <> 0 Then
        
        Dim tmpListIndex As Long
        tmpListIndex = Me.ListIndexByPosition(x, y)
        If (tmpListIndex >= 0) Then
            
            'If this item has children, we need to see *where* the click happened (because clicks in the control
            ' area mean we need to expand/collapse this group instead).
            If m_Items(tmpListIndex).hasChildren Then
                If PDMath.IsPointInRectF(x, (y - m_ListRectF.Top) + m_ScrollValue, m_Items(tmpListIndex).controlRect) Then
                    
                    'Expand this node, but importantly, do *not* attempt to bring the currently selected
                    ' node on-screen (as the user may have scrolled elsewhere in the list, then expanded something).
                    ToggleCollapseStateByIndex tmpListIndex
                    
                Else
                    Me.ListIndex = tmpListIndex
                End If
            
            'If this item *doesn't* have children, simply select it.
            Else
                Me.ListIndex = tmpListIndex
            End If
        
        '/end click is over a list item
        End If
    
    '/end left mouse click
    End If
    
End Sub

'Double-clicks toggle expand/contract state, but *only* if the click is *not* over the toggle UI!
Friend Sub NotifyMouseDoubleClick(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    If (Button And pdLeftButton) <> 0 Then
        
        Dim tmpListIndex As Long
        tmpListIndex = Me.ListIndexByPosition(x, y)
        If (tmpListIndex >= 0) Then
            
            'If this item has children, we need to see *where* the click happened (because clicks in the control
            ' area mean we need to expand/collapse this group instead).
            If m_Items(tmpListIndex).hasChildren Then
                If (Not PDMath.IsPointInRectF(x, (y - m_ListRectF.Top) + m_ScrollValue, m_Items(tmpListIndex).controlRect)) Then
                    ToggleCollapseStateByIndex tmpListIndex
                    Me.ListIndex = ValidateIndexIsVisible(Me.ListIndex, True)
                End If
            
            'If this item *doesn't* have children, simply select it.
            Else
                Me.ListIndex = tmpListIndex
            End If
        
        '/end click is over a list item
        End If
    
    '/end left mouse click
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

'When the user presses left on a node that *isn't* a parent, we need to jump up to the last parent in the tree.
'
'This function does nothing if the current node is already a parent.  This function also doesn't validate the
' incoming index; the caller is responsible for that.
'
'Returns: index of the correct parent index to jump to.
Private Function JumpToNearestParent(ByVal idxCurrent As Long) As Long
    
    If (m_NumOfItems < 0) Then Exit Function
    If (idxCurrent < 0) Then
        JumpToNearestParent = 0
        Exit Function
    End If

    Dim i As Long
    For i = idxCurrent - 1 To 0 Step -1
        If m_Items(i).hasChildren Then
            JumpToNearestParent = i
            Exit Function
        End If
    Next i
    
    'If we reach this point, no valid parent was found (will only open on a list consisting only
    ' of top-level items with no children).  Default to the first item in the list.
    JumpToNearestParent = 0
    
End Function

'Toggle a given node's collapse state.  Alternatively, you can use optional parameters to override collapse state
' to a specific value.
'
'Does nothing if the target node is not a parent node.  Does not validate incoming index (for range *or* visibility).
Private Sub ToggleCollapseStateByIndex(ByVal idxTarget As Long, Optional ByVal overrideToggle As Boolean = False, Optional ByVal overrideCollapseValue As Boolean = True)
    
    'Ignore nodes that do not have children
    If (Not m_Items(idxTarget).hasChildren) Then Exit Sub
    
    Dim initState As Boolean
    initState = m_Items(idxTarget).isCollapsed
    
    'Change collapse state for this node, then redraw the list accordingly.
    If overrideToggle Then
        m_Items(idxTarget).isCollapsed = overrideCollapseValue
    Else
        m_Items(idxTarget).isCollapsed = Not m_Items(idxTarget).isCollapsed
    End If
    
    'TODO: a simpler render metrics assessment, advancing only from the collapsed index (because higher
    ' nodes would not be affected by the change?)
    If (m_Items(idxTarget).isCollapsed <> initState) Then CalculateRenderMetrics
    
End Sub

'When a new ListIndex value is set, PD makes sure that item appears on-screen.
' This function will return TRUE if the current scroll value was changed to bring the item on-screen.
Private Function MakeSureListIndexFitsOnscreen() As Boolean

    MakeSureListIndexFitsOnscreen = False
    
    'If the item's current top and bottom values fit with the listview's container area, we don't need to change anything.
    If (m_ListIndex >= 0) And (m_ListIndex < m_NumOfItems) Then
        
        Dim liTop As Single, liBottom As Single
        liTop = m_Items(m_ListIndex).itemTop
        liBottom = liTop + m_DefaultHeight
        
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
                m_ScrollValue = (liTop - (m_ListRectF.Height - m_DefaultHeight)) - 1
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

'Need to render the tree?  Call this first to get rendering limits.
Friend Sub GetRenderingLimits(ByRef firstRenderIndex As Long, ByRef lastRenderIndex As Long, ByRef listIsEmpty As Boolean)
    firstRenderIndex = m_FirstRenderIndex
    lastRenderIndex = m_LastRenderIndex
    listIsEmpty = (m_NumOfItems = 0)
End Sub

'Need to render a specific item?  Call this to retrieve a full copy of a given item's data,
' plus rendering-specific information like the item's literal position in the current view.
'
'RETURNS: TRUE if the requested item is *not* collapsed, FALSE if it is (so you shouldn't render anything
' if this function returns false!)
Friend Function GetRenderingItem(ByVal srcListIndex As Long, ByRef dstListItem As PD_TreeItem, ByRef dstScrollOffsetX As Long, ByRef dstScrollOffsetY As Long) As Boolean
    GetRenderingItem = Not m_Items(srcListIndex).isCollapsedThisRender
    If GetRenderingItem Then
        dstListItem = m_Items(srcListIndex)
        dstScrollOffsetX = 0
        dstScrollOffsetY = m_ScrollValue
    End If
End Function

'While this class doesn't do any actual rendering, it does calculate all relevant scroll bar and positioning values.
' This makes life easier on the caller.

'Obviously, these values may not be correct if the class has not been notified of the list box size and/or not all of
' its contents have been loaded yet.
Friend Property Get ScrollMax() As Long
    If (Not m_RenderDataCorrect) Then
        If m_RedrawAutomatically Then
            CalculateRenderMetrics
        Else
            m_ScrollMax = m_TotalHeight - m_ListRectF.Height
            If (m_ScrollMax < 0) Then m_ScrollMax = 0
        End If
    End If
    ScrollMax = m_ScrollMax
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
    
    'Failsafes (usually only triggered in the IDE)
    If (m_ListRectF.Height <= 0) Or (m_DefaultHeight <= 0) Then Exit Sub
    
    'Current y-offset (accounting for collapse state).
    Dim curOffsetYincCollapse As Long
    curOffsetYincCollapse = 0
    
    'List of parents for the current node.  PD doesn't need this to go especially deep (only two levels at present).
    Const INIT_MAX_PARENTS As Long = 8
    Dim activeParents() As String, idxParent As Long
    ReDim activeParents(0 To INIT_MAX_PARENTS - 1) As String
    idxParent = 0
    
    'Before calculating what to render, we need to determine open/close state.  Because PD doesn't provide
    ' more than a few hundred items in a treeview, we iterate the whole list every time and calculate
    ' collapse state accordingly.  (In the future, we'd really only need to do this from the top-most collapse
    ' item through the next item at the same tree level - anything outside those bounds wouldn't change.)
    Dim inCollapseNow As Boolean: inCollapseNow = False
    Dim idxToCheck As Variant
                    
    Dim i As Long
    For i = 0 To m_NumOfItems - 1
        
        'First, we need to assess the depth of the current node.
        ' (Because this node may not belong to the same parent as the previous node.)
        If (idxParent > 0) Then
            
            'If this item has the same parent as the previous item, we remain at the same tree level
            ' (and can just copy the current collapse state).
            '
            'But if this item has a *different* parent from the last item, it could appear at any
            ' level in the tree.  (Think the last item of the File menu, followed by the Edit menu.)
            If (m_Items(i).parentID <> activeParents(idxParent - 1)) Then
                
                'This node has a parent ID that is different from the currently active parent.
                ' We must retreat through the parent table until we arrive at the correct tree level.
                Dim j As Long
                For j = idxParent - 1 To 0 Step -1
                    
                    'If we reach the top level of the tree, abandon further searching
                    ' (because this is a top-level menu).
                    If (j = 0) Then
                        inCollapseNow = False
                        idxParent = 0
                        Exit For
                    
                    'We are not at the top level of the tree yet...
                    Else
                        
                        If (m_Items(i).parentID = activeParents(j - 1)) Then
                            
                            'Reset the current tree level tracker to match.
                            idxParent = j
                            
                            'Mirror the collapse state of our parent.
                            If m_ItemHash.GetItemByKey(activeParents(j - 1), idxToCheck) Then
                                If m_Items(idxToCheck).isCollapsedThisRender Then
                                    inCollapseNow = True
                                Else
                                    inCollapseNow = m_Items(idxToCheck).isCollapsed
                                End If
                                
                            End If
                            
                            Exit For
                            
                        '/Ignore mismatched parent IDs - we'll keep retreating as necessary
                        End If
                    
                    End If
                    
                Next j
                
                'This node was already fully processed in the previous for loop
                
            End If
            
        'If the previous node was at the top-level, it will have moved us to a deeper level,
        ' or we're still at top-level and can't be collapsed!
        Else
            inCollapseNow = False
        End If
        
        'We now know whether this node is collapsed; assign its value accordingly.
        m_Items(i).isCollapsedThisRender = inCollapseNow
        
        'We now need to deal with whether or not this item has children.  If it does, we need to set up
        ' collapse info for the next node.
        If m_Items(i).hasChildren Then
            
            'Note the current tree level (if any), then advance the parent tracker
            m_Items(i).numParents = idxParent
            activeParents(idxParent) = m_Items(i).itemID
            
            idxParent = idxParent + 1
            If (idxParent > UBound(activeParents)) Then ReDim Preserve activeParents(0 To idxParent * 2 - 1) As String
            
            If (Not inCollapseNow) Then inCollapseNow = m_Items(i).isCollapsed
            
        'If this item *doesn't* have children, we need to compare its parent to the previous node's
        ' and potentially retreat the depth tracker.
        Else
            
            'Nodes without children do not independently track collapse state
            m_Items(i).isCollapsed = False
            m_Items(i).isCollapsedThisRender = inCollapseNow
            m_Items(i).numParents = idxParent
            
        '/end If/Else this node has children
        End If
        
        'We now know enough about the current node to figure out its on-screen rect.  Note that we
        ' calculate a rect *even for collapsed nodes*, because when it comes time to actually render
        ' the tree, we can simply skip collapsed nodes while subtracting their offset as-we-go.
        With m_Items(i)
            
            'Each item has three rects:
            ' 1) itemRect (full width of the list box)
            ' 2) collapseRect (control area for expand/contract only, shifts right by tree level)
            ' 3) captionRect (text area, shifts right by tree level)
            
            'We calculate all three rects for all nodes, and all rects use coordinates as if every node
            ' in the list were expanded.
            .ItemRect.Left = m_ListRectF.Left
            .ItemRect.Width = m_ListRectF.Width
            .ItemRect.Top = curOffsetYincCollapse
            .ItemRect.Height = m_DefaultHeight - 1
            
            .controlRect.Left = m_ListRectF.Left + m_OffsetPerLevel * .numParents
            .controlRect.Width = m_OffsetPerLevel
            .controlRect.Top = curOffsetYincCollapse + 1
            .controlRect.Height = m_DefaultHeight - 3
            
            'Always add an extra offset in case we need to render a caret to the left
            .captionRect.Left = m_ListRectF.Left + m_OffsetPerLevel * (.numParents + 1)
            .captionRect.Width = m_ListRectF.Width - .captionRect.Left
            .captionRect.Top = curOffsetYincCollapse
            .captionRect.Height = m_DefaultHeight - 1
            
            'This standalone item top value *does* include collapse state
            .itemTop = curOffsetYincCollapse
            If (Not .isCollapsedThisRender) Then curOffsetYincCollapse = curOffsetYincCollapse + m_DefaultHeight
            
        End With
        
    Next i
    
    'We now have rendering data available for all items in the tree, both idealized *and* adjusted for collapse state.
    
    'We now want to iterate the tree and figure out which items should appear on the list.
    m_FirstRenderIndex = -1
    m_LastRenderIndex = -1
    Dim idxLastUncollapsed As Long: idxLastUncollapsed = -1
    
    For i = 0 To m_NumOfItems - 1
    
        'Have we reached a visible item, or are we still on the hunt for the *first* visible one?
        If (m_FirstRenderIndex >= 0) Then
            
            'Ignore collapsed items
            If (Not m_Items(i).isCollapsedThisRender) Then
                
                'If this item is outside the displayable area, set
                If (m_Items(i).itemTop >= m_ScrollValue + m_ListRectF.Height) Then
                    m_LastRenderIndex = idxLastUncollapsed
                    Exit For
                
                'If this item *is* inside the displayable area, mark it as the last-checked index.
                Else
                    idxLastUncollapsed = i
                End If
                
            End If
            
        'We haven't found the first displayable item yet
        Else
            
            'If this item is *not* collapsed and *would* appear on-screen, mark it as the first
            ' item to render.
            If (Not m_Items(i).isCollapsedThisRender) Then
                If (m_Items(i).itemTop + m_Items(i).ItemRect.Height >= m_ScrollValue) Then
                    m_FirstRenderIndex = i
                    idxLastUncollapsed = i
                End If
            End If
            
        End If
        
    Next i
    
    'Failsafe for scrolling past the end of the list
    If (m_LastRenderIndex < m_FirstRenderIndex) Then m_LastRenderIndex = m_NumOfItems - 1
    If (idxLastUncollapsed < 0) Then
        idxLastUncollapsed = m_NumOfItems - 1
        m_LastRenderIndex = m_NumOfItems - 1
    End If
    If (m_FirstRenderIndex < 0) Then m_FirstRenderIndex = 0
    
    'Note that our rendering data is up-to-date.  As long as this stays TRUE, we don't have to recalculate rendering data.
    m_RenderDataCorrect = True
    
    'Whenever list contents or container sizes change, we can cache a new scroll bar maximum value.
    Dim newScrollMax As Long
    m_TotalHeight = curOffsetYincCollapse
    
    newScrollMax = m_TotalHeight - m_ListRectF.Height
    If (newScrollMax < 0) Then newScrollMax = 0
    
    If (newScrollMax <> m_ScrollMax) Then
        m_ScrollMax = newScrollMax
        RaiseEvent ScrollMaxChanged
    End If
    
    'If the selected item is now invisible, unselect it
    If (Me.ListIndex >= 0) Then
        If m_Items(Me.ListIndex).isCollapsedThisRender Then Me.ListIndex = -1
    End If
    
    'If automatic redraws are allowed, render everything now
    If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
    
End Sub

Friend Sub UpdateAgainstCurrentTheme()
    
    If (Not g_Language Is Nothing) Then
        
        If (m_NumOfItems > 0) And (m_LanguageAtLastCheck <> g_Language.GetCurrentLanguage()) Then
            m_LanguageAtLastCheck = g_Language.GetCurrentLanguage
            m_RenderDataCorrect = False
            Me.CalculateRenderMetrics
        End If
        
    End If

End Sub
