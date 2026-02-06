VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdListSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon List Box support class
'Copyright 2015-2026 by Tanner Helland
'Created: 22/December/15
'Last updated: 04/June/25
'Last update: fix potential crash in .AddItem when passed an OOB initial index
'
'PD makes use of a lot of custom-drawn list boxes.  To simplify these, I've tried to move a bunch of list-box-agnostic
' code elements into this class, elements that can then be reused by individual boxes.  This can be somewhat confusing,
' because this class does not do any actual UI rendering (that's left to the parent class), but it *does* manage a lot
' of rendering-related data, like scroll bar offsets.
'
'For a good understanding of how this class works, I recommend looking through the pdListBoxView usercontrol.  That UC
' uses this class to implement a generic, text-only list box, and it does so with very little custom code (basically just
' enough to bubble up certain input events to this class, then rendering list item text when this class raises a
' RedrawNeeded() event.)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control raises a few helper events, mostly related to rendering.  If our owner does something that requires a redraw,
' we'll raise a "RedrawNeeded" event, which the owner can then respond to at their leisure.  (Or ignore, if e.g. they're
' not visible at the moment.)
'
'As far as rendering is concerned, please note that some events are always synonymous with a redraw (e.g. Click() always
' means the .ListIndex has changed, which in turn means a redraw is required).  You *do not* have to render content from
' within the Click() event - you will always get a corresponding RedrawNeeded() event to match.
Public Event RedrawNeeded()
Public Event ScrollMaxChanged()
Public Event ScrollValueChanged()
Public Event Click()

'This class can calculate list box element offsets using three different systems:
' - Fixed size (all list elements are the same height)
' - Separators (all list elements are the same height, but some have divider lines after them)
' - Custom size (the owner controls each element height independently)
'By default, fixed size is assumed.
Private m_SizeMode As PD_ListboxHeight

'This class can also adjust how it responds to certain input events
Private m_ListMode As PD_ListSupportMode

'If custom sizes are in use, PD needs to store each listbox item position independently, since we can't infer it.
' Note that this struct (and the corresponding position array) may or may not be in use - or certain parts of it
' may be in use, but not other parts - depending on the combination of parameters sent to us by the user.
Private m_Items() As PD_ListItem

'Current list item count.  May or may not correspond to the size of m_Items, FYI.
Private Const INITIAL_LIST_SIZE As Long = 16&
Private m_NumOfItems As Long

'Default height of a list box item.  This is controlled by the owner, and we cannot do anything useful until its value is set.
Private m_DefaultHeight As Long

'Total height of the entire list box, as it would appear without scrolling.  This is used to calculate scrollbar values.
Private m_TotalHeight As Long

'Divider height is calculated as a fraction of the default line height.  This makes things like DPI handling much simpler.
Private Const SEPARATOR_LINE_HEIGHT As Single = 0.55!

'This class will track .ListIndex for you.  It needs to know that value to auto-suggest things like scroll bar position
' after a keypress.
Private m_ListIndex As Long

'Scroll bar values.  This class doesn't do any actual rendering, but it will track things like scroll value to make life easier
' on the caller (and because we have sufficient data to do all those calculations anyway).
Private m_ScrollMax As Long, m_ScrollValue As Long
Private m_ListIndexHover As Long, m_MouseInsideList As Boolean
Private m_ListRectF As RectF
Private m_InitMouseX As Single, m_InitMouseY As Single, m_InitScroll As Single
Private m_LMBDown As Boolean

'Typically, adding an item to the list box requires us to redraw the list box.  This is a waste of time if the caller needs
' to add a bunch of items all in a row.  Instead of raising redraws automatically, the caller can choose to suspend redraws
' prior to adding items, then enable redraws after all items have been added.
Private m_RedrawAutomatically As Boolean

'Whenever a property changes that affects the on-screen appearance of the list (e.g. adding an item, scrolling the list),
' we'll immediately cache the first and last elements that need to be drawn on-screen.  Then, when it comes time to render
' the list, we don't have to regenerate that info from scratch.
Private m_FirstRenderIndex As Long, m_LastRenderIndex As Long

'If the list is *not* in automatic redraw mode (where redraw notifications are raised on every list change), we won't
' calculate rendering metrics as we go.  Instead, we'll just mark rendering metrics as dirty, and recalculate them when
' the owner finally requests rendering data.
Private m_RenderDataCorrect As Boolean

'When the list is initialized, we note the current language.  If the language changes on subsequent theme updates,
' we'll re-translate the listbox to match.
Private m_LanguageAtLastCheck As String

Private Sub Class_Initialize()
    
    m_ListMode = PDLM_ListBox
    m_SizeMode = PDLH_Fixed
    
    m_DefaultHeight = 0
    m_RedrawAutomatically = True
    m_MouseInsideList = False
    m_ListIndexHover = -1
    
    Me.Clear
    
End Sub

'Add an item to the list.  Note that all parameters are optional; if the owner is managing a list with custom data, for example,
' they do not need to pass strings to this function.  Similarly, things like separators and custom height can be happily
' ignored if those features are not required.
Friend Sub AddItem(Optional ByRef srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1, Optional ByVal itemShouldBeTranslated As Boolean = True)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Make sure there's room in the array for this item.
    If (m_NumOfItems > UBound(m_Items)) Then ReDim Preserve m_Items(0 To m_NumOfItems * 2 - 1) As PD_ListItem
    
    'Change the rendering mode, as necessary
    If (itemHeight <> -1) And (m_SizeMode <> PDLH_Custom) Then
        m_SizeMode = PDLH_Custom
    ElseIf (hasTrailingSeparator And (m_SizeMode <> PDLH_Separators)) Then
        m_SizeMode = PDLH_Separators
    End If
    
    'If this item is being inserted anywhere but the end of the list, move other items around to make room.
    If (itemIndex <> -1) And (itemIndex < m_NumOfItems) Then
        
        'Migrate all objects downward in the list.
        Dim i As Long
        For i = m_NumOfItems To (itemIndex + 1) Step -1
            m_Items(i) = m_Items(i - 1)
        Next i
        
    'If no insert index has been specified, assume the insertion happens at the end of the list
    ElseIf (itemIndex = -1) Or (itemIndex > m_NumOfItems) Then
        itemIndex = m_NumOfItems
    End If
    
    'Insert the given item
    With m_Items(itemIndex)
        .isSeparator = hasTrailingSeparator
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
        
        'Calculating height is a bit more involved...
        If (itemHeight = -1) Then
            
            If .isSeparator Then
                .itemHeight = m_DefaultHeight + (m_DefaultHeight * SEPARATOR_LINE_HEIGHT)
            Else
                .itemHeight = m_DefaultHeight
            End If
            
        'If the user specifies a height, assume it's correct.  Any positioning issues are theirs to deal with.
        ' (NOT IMPLEMENTED YET!)
        Else
            .itemHeight = m_DefaultHeight
        End If
        
        'Increase the net height of the entire list
        m_TotalHeight = m_TotalHeight + .itemHeight
        
        'Positioning values are really only used if separators are active, or if user-specified heights are involved,
        ' but right now we set those positions correctly for any insertion action.  (We can revisit in the future
        ' if performance becomes an issue.)
        If (itemIndex < m_NumOfItems) Then
            
            'Set the current item's position.
            If (itemIndex = 0) Then
                .itemTop = 0
            Else
                .itemTop = m_Items(itemIndex - 1).itemTop + m_Items(itemIndex - 1).itemHeight
            End If
            
            'Add this item's height to all subsequent positions.
            For i = itemIndex + 1 To m_NumOfItems
                m_Items(i).itemTop = m_Items(i).itemTop + .itemHeight
            Next i
        
        'If this item is being inserted at the end of the list, simply plug it into place.
        Else
            If (itemIndex > 0) Then
                .itemTop = m_Items(itemIndex - 1).itemTop + m_Items(itemIndex - 1).itemHeight
            Else
                .itemTop = 0
            End If
        End If
        
    End With
    
    'If this is the first item, note the current translation language.  (If this changes, we need to re-translate the list.)
    If (m_NumOfItems = 0) Then
        If (Not g_Language Is Nothing) Then m_LanguageAtLastCheck = g_Language.GetCurrentLanguage
    End If
    
    'Increment the number of list entries
    m_NumOfItems = m_NumOfItems + 1
    
    'If this item is beneath the list index, bump up the list index by one
    If (itemIndex < m_ListIndex) Then
        m_ListIndex = m_ListIndex + 1
        RaiseEvent Click
    End If
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
End Sub

'Reset the current list.  An optional starting list size can be passed; if it is not passed, it will default to INITIAL_LIST_SIZE.
Friend Sub Clear(Optional ByVal newListSize As Long = INITIAL_LIST_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Failsafe bounds check
    If (newListSize <= 0) Then newListSize = INITIAL_LIST_SIZE
    
    'Reset the array (but only if necessary!)
    If (m_NumOfItems = 0) Then
        ReDim m_Items(0 To newListSize - 1) As PD_ListItem
    Else
        If UBound(m_Items) = newListSize - 1 Then
            Dim i As Long
            For i = 0 To UBound(m_Items)
                With m_Items(i)
                    .isSeparator = False
                    .itemHeight = 0
                    .itemTop = 0
                    .textEn = vbNullString
                    .textTranslated = vbNullString
                End With
            Next i
        Else
            ReDim m_Items(0 To newListSize - 1) As PD_ListItem
        End If
    End If
    
    'Reset some obvious things (that don't require special handling)
    m_ListIndex = -1
    m_NumOfItems = 0
    m_TotalHeight = 0
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
    Exit Sub
    
FailsafeReset:
    If (newListSize <= 0) Then newListSize = INITIAL_LIST_SIZE
    ReDim m_Items(0 To newListSize - 1) As PD_ListItem
    
End Sub

'Want to clone some other list instance?  This function handles it in one fell swoop, including initializing our internal
' rendering parameters.
Friend Sub CloneExternalListSupport(ByRef srcListSupport As pdListSupport, Optional ByVal desiredListIndexTop As Long = 0, Optional ByVal newListSupportMode As PD_ListSupportMode = PDLM_LB_Inside_DD)
    
    'The order in which we copy data from the target list support is crucial; copying in the
    ' wrong order will cause crashes, especially because we are copying the data using direct
    ' pointer manipulation.
    
    'In other words: don't mess with this function.
    
    'Start by grabbing the raw m_Items array
    m_NumOfItems = srcListSupport.ListCount
    If (m_NumOfItems > 0) Then
        
        ReDim m_Items(0 To srcListSupport.GetSizeOfInternalListStruct) As PD_ListItem
        
        Dim i As Long
        For i = 0 To m_NumOfItems - 1
            m_Items(i) = srcListSupport.GetDirectListItem(i)
        Next i
    
    Else
        Me.Clear
    End If
    
    'We do not clone all elements of the source pdListSupport class.  Things like font size (and corresponding height values)
    ' are specific to *our* parent object, so they remain in place.
    
    'Clone the list index and a few cumbersome-to-recreate internal trackers
    m_ListIndex = srcListSupport.ListIndex
    m_TotalHeight = srcListSupport.GetHeightOfAllListItems
    m_SizeMode = srcListSupport.GetInternalSizeMode
    m_ListMode = newListSupportMode
    
    'With those main trackers in place, we now need to totally recalculate all rendering metrics, as they're specific to
    ' our owner's dimensions, fontsize, etc.
    CalculateRenderMetrics
    
    'Normally, we calculate a new scrollbar maximum during the CalculateRenderMetrics step, but during a clone, we want to
    ' set it early so we can position the scroll bar perfectly, according to our parent control's requested position.
    If (desiredListIndexTop <> 0) Then
        
        Dim newScrollValue As Single
        If (m_ListIndex >= 0) And (m_NumOfItems > 0) Then newScrollValue = (m_Items(m_ListIndex).itemTop - desiredListIndexTop)
        
        If (newScrollValue < 0) Then newScrollValue = 0
        If (newScrollValue > m_ScrollMax) Then newScrollValue = m_ScrollMax
        Me.ScrollValue = newScrollValue
        
    Else
        Me.ScrollValue = 0
    End If
    
End Sub

'Need to render a specific list item?  Call this to retrieve a full copy of a given list item's data, plus
' rendering-specific information like the item's literal position in the current list box.
Friend Function GetDirectListItem(ByVal srcListIndex As Long) As PD_ListItem
    GetDirectListItem = m_Items(srcListIndex)
End Function

Friend Function GetHeightOfAllListItems() As Long
    GetHeightOfAllListItems = m_TotalHeight
End Function

Friend Function GetInternalSizeMode() As PD_ListboxHeight
    GetInternalSizeMode = m_SizeMode
End Function

Friend Function GetSizeOfInternalListStruct() As Long
    GetSizeOfInternalListStruct = UBound(m_Items)
End Function

'Font size controls the default height of each list item.  When the font size changes, we need to recalculate a number of
' internal size metrics, so it's advisable to set this UP FRONT before doing anything else.
Friend Property Get DefaultItemHeight() As Long
    DefaultItemHeight = m_DefaultHeight
End Property

Friend Property Let DefaultItemHeight(ByVal newHeight As Long)
    If (m_DefaultHeight <> newHeight) Then
        m_DefaultHeight = newHeight
        
        'If a non-standard size mode is in use, we technically need to calculate new positioning metrics for all list items.
        ' This is stupid, and I'd prefer not to support it - so instead, just set the damn font size correctly *before* you
        ' add items to the list box!
        
    End If
End Property

Friend Function GetSeparatorHeight() As Single
    GetSeparatorHeight = (m_DefaultHeight * SEPARATOR_LINE_HEIGHT)
End Function

Friend Function DoesItemHaveSeparator(ByVal itemIndex As Long) As Boolean
    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        DoesItemHaveSeparator = m_Items(itemIndex).isSeparator
    End If
End Function

Friend Function IsMouseInsideListBox() As Boolean
    IsMouseInsideListBox = m_MouseInsideList
End Function

'Retrieve a specified list item
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

'You can set the ListIndex to -1 to indicate "nothing is selected", but this is stupid for a whole bunch of reasons, so please
' don't do it.  (I mention this because this function rejects invalid ListIndex requests, with that sole exception.)
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
    
    'On a fixed-size list, this calculation can be automated.
    If m_SizeMode = PDLH_Fixed Then
        ListIndexByPosition = srcY \ m_DefaultHeight
        If ListIndexByPosition >= m_NumOfItems Then ListIndexByPosition = -1
        
    'On a variable-size list, this calculation is more complicated
    Else
        
        Dim tmpRect As RectF
        tmpRect.Left = m_ListRectF.Left
        tmpRect.Width = m_ListRectF.Width
        
        'Because the (x, y) position may lie outside the visible container area, we need to perform a comprehensive hit search.
        Dim i As Long
        For i = 0 To m_NumOfItems - 1
            
            tmpRect.Top = m_Items(i).itemTop + m_ListRectF.Top
            tmpRect.Height = m_Items(i).itemHeight
            
            If checkXAsWell Then
                If PDMath.IsPointInRectF(srcX, srcY, tmpRect) Then
                    ListIndexByPosition = i
                    Exit For
                End If
            Else
                If PDMath.IsPointInRectF(1&, srcY, tmpRect) Then
                    ListIndexByPosition = i
                    Exit For
                End If
            End If
        Next i
        
    End If
    
End Function

'This property exists purely for improving UI feedback.  It may or may not be identical to the default .ListIndex value.
Friend Property Get ListIndexHovered() As Long
    ListIndexHovered = m_ListIndexHover
End Property

Friend Property Get ListSupportMode() As PD_ListSupportMode
    ListSupportMode = m_ListMode
End Property

Friend Property Let ListSupportMode(ByVal newMode As PD_ListSupportMode)
    m_ListMode = newMode
End Property

Friend Function UpdateItem(ByVal itemIndex As Long, ByVal newItemText As String) As Boolean

    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        
        If Strings.StringsNotEqual(newItemText, m_Items(itemIndex).textEn, vbBinaryCompare) Then
        
            m_Items(itemIndex).textEn = newItemText
            
            'Add a translated copy of the string as well; this will be the string actually rendered onto the screen.
            If (Not g_Language Is Nothing) Then
                If g_Language.TranslationActive Then
                    m_Items(itemIndex).textTranslated = g_Language.TranslateMessage(newItemText)
                Else
                    m_Items(itemIndex).textTranslated = newItemText
                End If
            Else
                m_Items(itemIndex).textTranslated = newItemText
            End If
            
        Else
            UpdateItem = False
        End If
        
    Else
        UpdateItem = False
    End If
    
    If (m_RedrawAutomatically And UpdateItem) Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
End Function

'While this class doesn't actually render anything (that's left to the parent), it has enough information to manage a lot
' of the backend rendering details automatically.  This makes it much easier to construct custom list boxes, as things like
' hit detection can be handled here - ensuring that all list boxes throughout the program, even custom ones, behave identically.
Friend Sub NotifyParentRectF(ByRef srcListRectF As RectF)
    If (m_ListRectF.Width <> srcListRectF.Width) Or (m_ListRectF.Height <> srcListRectF.Height) Or (m_ListRectF.Top <> srcListRectF.Top) Or (m_ListRectF.Left <> srcListRectF.Left) Then
        m_ListRectF = srcListRectF
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    End If
End Sub

'Key events primarily affect the current .ListIndex property.  Note that some shortcuts (like PAGEDOWN/PAGEUP) require
' up-to-date rendering information, as we need to know which item lies on the next "page" of the listview - information that
' directly correlates to the on-screen size of the listbox's constraining rect.
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
            If (m_ListMode = PDLM_ListBox) Or (m_ListMode = PDLM_LB_Inside_DD) Then
                newListIndex = newListIndex + (m_LastRenderIndex - m_FirstRenderIndex)
            Else
                newListIndex = newListIndex + 1
            End If
            markEventHandled = True
            
        Case VK_PAGEUP
            If (m_ListMode = PDLM_ListBox) Or (m_ListMode = PDLM_LB_Inside_DD) Then
                newListIndex = newListIndex - (m_LastRenderIndex - m_FirstRenderIndex)
            Else
                newListIndex = newListIndex - 1
            End If
            markEventHandled = True
            
        Case VK_HOME
            newListIndex = 0
            markEventHandled = True
            
        Case VK_END
            newListIndex = m_NumOfItems - 1
            markEventHandled = True
            
        'Return only matters for listboxes embedded inside combo boxes.
        Case VK_RETURN, VK_SPACE
            If (m_ListMode = PDLM_LB_Inside_DD) Then
                markEventHandled = True
                RaiseEvent Click
                Exit Sub
            Else
                markEventHandled = False
                Exit Sub
            End If
            
        'For listboxes embedded inside combo boxes, the escape key can hide the dropdown.  Raise a Click() event to mimic this.
        Case VK_ESCAPE
            If (m_ListMode = PDLM_LB_Inside_DD) Then
                markEventHandled = True
                RaiseEvent Click
                Exit Sub
            End If
            
    End Select
    
    If (m_NumOfItems = 0) Then
        newListIndex = -1
    Else
        If (newListIndex < 0) Then newListIndex = 0
        If (newListIndex > m_NumOfItems - 1) Then newListIndex = m_NumOfItems - 1
    End If
    
    'Listboxes dynamically raised by combo boxes should not raise Click() events on keypresses.  This will erroneously cause
    ' the list box to close.  Instead, they simply reassign their .ListIndex and redraw themselves immediately.
    If (m_ListMode = PDLM_LB_Inside_DD) Then
        
        m_ListIndex = newListIndex
        
        If MakeSureListIndexFitsOnscreen Then
            RaiseEvent ScrollValueChanged
            If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
        Else
            If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
        End If
        
    Else
        Me.ListIndex = newListIndex
    End If
    
End Sub

Friend Sub NotifyKeyUp(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

End Sub

'Mouse events control a whole bunch of possible things: hover state, .ListIndex, scroll.  As such, their handling is
' fairly involved, despite this class not doing any actual UI rendering.
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
    
    'The behavior of the mousewheel varies by the current "listmode", e.g. if our parent is a scrollable listview, we behave
    ' differently than if our parent is a closed dropdown box.
    If (m_ListMode = PDLM_ListBox) Or (m_ListMode = PDLM_LB_Inside_DD) Then
        
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
        
    'Dropdown mode
    Else
            
        If (m_NumOfItems > 1) Then
            
            Dim newListIndex As Long: newListIndex = Me.ListIndex
            If (scrollAmount > 0) Then newListIndex = newListIndex - 1 Else newListIndex = newListIndex + 1
            
            If (newListIndex < 0) Then newListIndex = 0
            If (newListIndex > m_NumOfItems - 1) Then newListIndex = m_NumOfItems - 1
            Me.ListIndex = newListIndex
            
        End If
    
    End If
    
End Sub

'When a new ListIndex value is set, PD makes sure that item appears on-screen.  This function will return TRUE if the current
' scroll value was changed to bring the item on-screen.
Private Function MakeSureListIndexFitsOnscreen() As Boolean

    MakeSureListIndexFitsOnscreen = False
    
    'If the item's current top and bottom values fit with the listview's container area, we don't need to change anything.
    If (m_ListIndex >= 0) And (m_ListIndex < m_NumOfItems) And (m_ListMode <> PDLM_DropDown) Then
        
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
        ScrollValue = tmpScrollValue
    End If
End Sub

'When the mouse enters, leaves, or moves inside the underlying control, we will update the current hover index (provided our
' parent relays those events to us, obviously).
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

'Remove an item from the combo box
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

'Need to render the list?  Call this first to get rendering limits.
Friend Sub GetRenderingLimits(ByRef firstRenderIndex As Long, ByRef lastRenderIndex As Long, ByRef listIsEmpty As Boolean)
    firstRenderIndex = m_FirstRenderIndex
    lastRenderIndex = m_LastRenderIndex
    listIsEmpty = (m_NumOfItems = 0)
End Sub

'Need to render a specific list item?  Call this to retrieve a full copy of a given list item's data,
' plus rendering-specific information like the item's literal position in the current list box.
Friend Sub GetRenderingItem(ByVal srcListIndex As Long, ByRef dstListItem As PD_ListItem, ByRef dstItemTop As Long, ByRef dstItemHeight As Long, ByRef dstItemHeightWithoutSeparator As Long)
    
    dstListItem = m_Items(srcListIndex)
    
    If (m_SizeMode = PDLH_Fixed) Then
        dstItemTop = (srcListIndex * m_DefaultHeight) - m_ScrollValue
        dstItemHeight = m_DefaultHeight
        dstItemHeightWithoutSeparator = m_DefaultHeight
    Else
        dstItemTop = m_Items(srcListIndex).itemTop - m_ScrollValue
        dstItemHeight = m_Items(srcListIndex).itemHeight
        dstItemHeightWithoutSeparator = m_DefaultHeight
    End If
    
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
            If m_ScrollMax < 0 Then m_ScrollMax = 0
        End If
    End If
End Property

Friend Property Get ScrollValue() As Long
    ScrollValue = m_ScrollValue
End Property

'When assigning a new scroll value, you should probably double-check the passed newValue.  This class will automatically reset the
' value to an appropriate range if it's too small or too large.
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

'The caller can suspend automatic redraws caused by things like adding an item to the list box.  Just make sure to enable redraws
' once you're ready, or you'll never get rendering requests!
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
    
    'Lists with uniform item sizes can skip a lot of messy handling.
    If (m_SizeMode = PDLH_Fixed) Then
    
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
        
    
    'If list entries have variable height, we need to loop through all list entries and look for ones that overlap the viewport.
    ' This step could be optimized by using something like a binary search to detect the first overlapping item, but because PD
    ' only uses short lists, I'm not particularly concerned with optimizing this right now.
    Else
        
        'To spare us from looping through unnecessary tail-end entries, if we've already raised at least one draw event, and we
        ' encounter a list item that exists outside the list, we can immediately exit the function.
        Dim oneItemFound As Boolean
        oneItemFound = False
        
        Dim itemTop As Single, itemBottom As Single
        m_FirstRenderIndex = LONG_MAX
        m_LastRenderIndex = -1 * LONG_MAX
        
        For i = 0 To m_NumOfItems - 1
            
            'Calculate top and bottom values *relative to the viewport* (e.g. with scrolling factored in)
            itemTop = (m_Items(i).itemTop - m_ScrollValue)
            itemBottom = itemTop + m_Items(i).itemHeight
            
            'If overlap is found, render the rect
            If ((itemTop >= 0) And (itemTop <= m_ListRectF.Height)) Or ((itemBottom >= 0) And (itemBottom <= m_ListRectF.Height)) Then
                
                If (i < m_FirstRenderIndex) Then m_FirstRenderIndex = i
                If (i > m_LastRenderIndex) Then m_LastRenderIndex = i
                
                oneItemFound = True
            
            'If no overlap is found, but overlap was found for a PREVIOUS entry, exit immediately, because no further hits
            ' will be detected.
            Else
                If oneItemFound Then Exit For
            End If
            
        Next i
        
    End If
    
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
