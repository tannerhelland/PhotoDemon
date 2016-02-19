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
'Copyright 2015-2016 by Tanner Helland
'Created: 22/December/15
'Last updated: 18/February/16
'Last update: continued integration with pdListBoxView
'
'PD makes use of a lot of custom-drawn list boxes.  To simplify these, I've tried to move a bunch of list-box-agnostic
' code elements into this class, elements that can then be reused by individual boxes.
'
'TODO: support standard height mode that is NOT based on text (e.g. the owner can specify their own custom height)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control raises a few helper events, mostly related to rendering.  If our owner does something that requires a redraw,
' we'll raise a "RedrawNeeded" event, which the owner can then respond to at their leisure.  (Or ignore, if e.g. they're
' not visible at the moment.)
Public Event RedrawNeeded()

'This class can calculate list box element offsets using three different systems:
' - Fixed size (all list elements are the same height)
' - Separators (all list elements are the same height, but some have divider lines after them)
' - Custom size (the owner controls each element height independently)
'By default, fixed size is assumed.
Private m_SizeMode As PD_LISTBOX_HEIGHT

'If custom sizes are in use, PD needs to store each listbox item position independently, since we can't infer it.
' Note that this struct (and the corresponding position array) may or may not be in use - or certain parts of it
' may be in use, but not other parts - depending on the combination of parameters sent to us by the user.
Private m_Items() As PD_LISTITEM

'Current list item count.  May or may not correspond to the size of m_Items, FYI.
Private Const INITIAL_LIST_SIZE As Long = 16&
Private m_NumOfItems As Long

'Default height of a list box item.  This is controlled by the owner, and we cannot do anything useful until its value is set.
Private m_DefaultHeight As Long

'Total height of the entire list box, as it would appear without scrolling.  This is used to calculate scrollbar values.
Private m_TotalHeight As Long

'Divider height is calculated as a fraction of the default line height.  This makes things like DPI handling much simpler.
Private Const SEPARATOR_LINE_HEIGHT As Single = 0.75

'Font size is used to determine default line height.  You can ignore it if the box will be fully custom-drawn.
Private m_FontSize As Single

'This class will track .ListIndex for you.  It needs to know that value to auto-suggest things like scroll bar position
' after a keypress.
Private m_ListIndex As Long

'Scroll bar values.  This class doesn't do any actual rendering, but it will track things like scroll value to make life easier
' on the caller (and because we have sufficient data to do all those calculations anyway).
Private m_ScrollMax As Long, m_ScrollValue As Long, m_ContainerHeight As Long, m_ContainerWidth As Long

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

Private Sub Class_Initialize()
    m_DefaultHeight = 0
    m_SizeMode = PDLH_FIXED
    m_RedrawAutomatically = True
    Me.Clear
End Sub

'Add an item to the list.  Note that all parameters are optional; if the owner is managing a list with custom data, for example,
' they do not need to pass strings to this function.  Similarly, things like separators and custom height can be happily
' ignored if those features are not required.
Public Sub AddItem(Optional ByVal srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1)
    
    'Make sure there's room in the array for this item.
    If m_NumOfItems > UBound(m_Items) Then ReDim Preserve m_Items(0 To m_NumOfItems * 2 - 1) As PD_LISTITEM
    
    'Change the rendering mode, as necessary
    If CBool(itemHeight <> -1) And CBool(m_SizeMode <> PDLH_CUSTOM) Then
        m_SizeMode = PDLH_CUSTOM
    ElseIf hasTrailingSeparator And CBool(m_SizeMode <> PDLH_SEPARATORS) Then
        m_SizeMode = PDLH_SEPARATORS
    End If
    
    'If this item is being inserted anywhere but the end of the list, move other items around to make room.
    Dim i As Long
    If CBool(itemIndex <> -1) And CBool(itemIndex < m_NumOfItems) Then
        
        'Migrate all objects downward in the list.
        For i = m_NumOfItems To (itemIndex + 1) Step -1
            m_Items(i) = m_Items(i - 1)
        Next i
        
    'If no insert index has been specified, assume the insertion happens at the end of the list
    ElseIf CBool(itemIndex = -1) Then
        itemIndex = m_NumOfItems
    End If
    
    'Insert the given item
    With m_Items(itemIndex)
        .isSeparator = hasTrailingSeparator
        .textEn = srcItemText
        
        'Add a translated copy of the string as well; this will be the string actually rendered onto the screen.
        If Not (g_Language Is Nothing) Then
            If g_Language.translationActive Then
                .textTranslated = g_Language.TranslateMessage(srcItemText)
            Else
                .textTranslated = srcItemText
            End If
        Else
            .textTranslated = srcItemText
        End If
        
        'Calculating height is a bit more involved...
        If CBool(itemHeight = -1) Then
            
            If .isSeparator Then
                .itemHeight = m_DefaultHeight + (m_DefaultHeight * SEPARATOR_LINE_HEIGHT)
            Else
                .itemHeight = m_DefaultHeight
            End If
            
        'If the user specifies a height, assume it's correct.  Any positioning issues are theirs to deal with.
        Else
            .itemHeight = m_DefaultHeight
        End If
        
        'Increase the net height of the entire list
        m_TotalHeight = m_TotalHeight + .itemHeight
        
        'Positioning values are really only used if separators are active, or if user-specified heights are involved,
        ' but right now we right those positions correctly for any insertion action.  (We can revisit in the future
        ' if performance becomes an issue.)
        If itemIndex < m_NumOfItems Then
            
            'Set the current item's position.
            If itemIndex = 0 Then
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
            If itemIndex > 0 Then
                .itemTop = m_Items(itemIndex - 1).itemTop + m_Items(itemIndex - 1).itemHeight
            Else
                .itemTop = 0
            End If
        End If
        
    End With
    
    'If this item is beneath the list index, bump up the list index by one
    If itemIndex < m_ListIndex Then m_ListIndex = m_ListIndex + 1
    
    'Increment the number of list entries
    m_NumOfItems = m_NumOfItems + 1
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
End Sub

'Reset the current list.  An optional starting list size can be passed; if it is not passed, it will default to INITIAL_LIST_SIZE.
Public Sub Clear(Optional ByVal newListSize As Long = INITIAL_LIST_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Reset some obvious things (that don't require special handling)
    m_ListIndex = 0
    m_NumOfItems = 0
    m_TotalHeight = 0
    
    'Failsafe bounds check
    If newListSize <= 0 Then newListSize = INITIAL_LIST_SIZE
    
    'Reset the array (but only if necessary!)
    If m_NumOfItems = 0 Then
        ReDim m_Items(0 To newListSize - 1) As PD_LISTITEM
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
            ReDim m_Items(0 To newListSize - 1) As PD_LISTITEM
        End If
    End If
    
    m_NumOfItems = 0
    
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
    Exit Sub
    
FailsafeReset:
    If newListSize <= 0 Then newListSize = INITIAL_LIST_SIZE
    ReDim m_Items(0 To newListSize - 1) As PD_LISTITEM
    
End Sub

'Font size controls the default height of each list item.  When the font size changes, we need to recalculate a number of
' internal size metrics, so it's advisable to set this UP FRONT before doing anything else.
Public Property Get DefaultItemHeight() As Single
    DefaultItemHeight = m_DefaultHeight
End Property

Public Property Let DefaultItemHeight(ByVal newHeight As Single)
    If m_DefaultHeight <> newHeight Then
        m_DefaultHeight = newHeight
        
        'If a non-standard size mode is in use, we technically need to calculate new positioning metrics for all list items.
        ' This is stupid, and I'd prefer not to support it - so instead, just set the damn font size correctly *before* you
        ' add items to the list box!
        
    End If
End Property

'Retrieve a specified list item
Public Function List(ByVal itemIndex As Long, Optional ByVal returnTranslatedText As Boolean = False) As String
    
    If (itemIndex >= 0) And (itemIndex < m_NumOfItems) Then
        If returnTranslatedText Then
            List = m_Items(itemIndex).textEn
        Else
            List = m_Items(itemIndex).textTranslated
        End If
    Else
        List = vbNullString
    End If
    
End Function

Public Function ListCount() As Long
    ListCount = m_NumOfItems
End Function

Public Property Get ListIndex() As Long
    ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal newIndex As Long)
    If (newIndex >= 0) And (newIndex < m_NumOfItems) Then
        m_ListIndex = newIndex
        
        'TODO: if the listindex lies outside the current rendering area, raise a request for a new scroll value
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
        
    End If
End Property

'As a convenience, this class also lets the caller set a ListIndex by string.  (If a match cannot be found,
' the list index will not be changed.)
'The optional ReadValueOnly parameter can be used to simply RETURN the value of the discovered ListIndex, without actually
' changing the ListIndex property.
Public Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare, Optional ByVal ReadValueOnly As Boolean = False) As Long
    
    ListIndexByString = -1
    
    If m_NumOfItems > 0 Then
        
        Dim newIndex As Long
        newIndex = -1
        
        Dim i As Long
        For i = 0 To m_NumOfItems - 1
            If StrComp(srcString, m_Items(i).textEn, compareMode) = 0 Then
                newIndex = i
                Exit For
            End If
        Next i
        
        'If a match was found, change the list index now
        If (newIndex >= 0) And (newIndex <> m_ListIndex) Then
            If Not ReadValueOnly Then m_ListIndex = newIndex
            ListIndexByString = newIndex
        End If
        
    End If
    
End Function

'As a convenience, this class also lets the caller set a ListIndex by mouse position within the container.  (If a match cannot
' be found, the list index will not be changed.)
'The optional ReadValueOnly parameter can be used to simply RETURN the value of the discovered ListIndex, without actually
' changing the ListIndex property.
Public Function ListIndexByPosition(ByVal srcX As Single, ByVal srcY As Single, Optional ByVal checkXAsWell As Boolean = True) As Long
    
    ListIndexByPosition = -1
    
    'First, do a spot-check on srcX.  If it lies outside the list region, skip this whole step.
    If checkXAsWell Then
        If (srcX < 0) Or (srcX > m_ContainerWidth) Then
            ListIndexByPosition = -1
            Exit Function
        End If
    End If
    
    'Convert the y-position to an absolute value
    srcY = srcY + m_ScrollValue
    
    'On a fixed-size list, this calculation can be automated.
    If m_SizeMode = PDLH_FIXED Then
        ListIndexByPosition = srcY \ m_DefaultHeight
        If ListIndexByPosition > m_TotalHeight Then ListIndexByPosition = -1
    
    'On a variable-size list, this calculation is more complicated
    Else
        
        Dim tmpRect As RECTF
        tmpRect.Left = 0
        tmpRect.Width = m_ContainerWidth
        
        'Because the (x, y) position may lie outside the visible container area, we need to perform a comprehensive hit search.
        Dim i As Long
        For i = 0 To m_NumOfItems - 1
            
            tmpRect.Top = m_Items(i).itemTop
            tmpRect.Height = m_Items(i).itemHeight
            
            If checkXAsWell Then
                If Math_Functions.IsPointInRectF(srcX, srcY, tmpRect) Then
                    ListIndexByPosition = i
                    Exit For
                End If
            Else
                If Math_Functions.IsPointInRectF(1&, srcY, tmpRect) Then
                    ListIndexByPosition = i
                    Exit For
                End If
            End If
        Next i
        
    End If
    
End Function

'Remove an item from the combo box
Public Sub RemoveItem(ByVal itemIndex As Long)
    
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
        If itemIndex <= m_ListIndex Then m_ListIndex = m_ListIndex - 1
        If m_ListIndex < 0 Then m_ListIndex = 0
        
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
        
    End If
    
End Sub

'Need to render the list?  Call this first to get rendering limits.
Public Sub GetRenderingLimits(ByRef firstRenderIndex As Long, ByRef lastRenderIndex As Long, ByRef listIsEmpty As Boolean)
    firstRenderIndex = m_FirstRenderIndex
    lastRenderIndex = m_LastRenderIndex
    listIsEmpty = CBool(m_NumOfItems = 0)
End Sub

'Need to render a specific list item?  Call this to retrieve a full copy of a given list item's data, plus
' rendering-specific information like the item's literal position in the current list box.
Friend Sub GetRenderingItem(ByVal srcListIndex As Long, ByRef dstListItem As PD_LISTITEM, ByRef dstItemTop As Long, ByRef dstItemHeight As Long)
    
    dstListItem = m_Items(srcListIndex)
    
    If m_SizeMode = PDLH_FIXED Then
        dstItemTop = (srcListIndex * m_DefaultHeight) - m_ScrollValue
        dstItemHeight = m_DefaultHeight
    Else
        dstItemTop = m_Items(srcListIndex).itemTop
        dstItemHeight = m_Items(srcListIndex).itemHeight
    End If
    
End Sub

'While this class doesn't do any actual rendering, it does calculate all relevant scroll bar and positioning values.
' This makes life easier on the caller.
Public Function ScrollMax(ByVal pxWidthOfList As Long, ByVal pxHeightOfList As Long) As Long
    If (m_ContainerWidth <> pxWidthOfList) Or (m_ContainerHeight <> pxHeightOfList) Then
        m_ContainerWidth = pxWidthOfList
        m_ContainerHeight = pxHeightOfList
        m_ScrollMax = m_TotalHeight - pxHeightOfList
        If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
        ScrollMax = m_ScrollMax
    End If
End Function

Public Property Get ScrollValue() As Long
    ScrollValue = m_ScrollValue
End Property

'When assigning a new scroll value, you should probably double-check the passed newValue.  This class will automatically reset the
' value to an appropriate range if it's too small or too large.
Public Property Let ScrollValue(ByRef newValue As Long)
    
    'Range-check the incoming value
    If newValue < 0 Then newValue = 0
    If newValue > m_ScrollMax Then newValue = m_ScrollMax
    m_ScrollValue = newValue
    
    'Changing the scroll value changes the on-screen position of list elements, so we need to recalculate rendering data.
    If m_RedrawAutomatically Then CalculateRenderMetrics Else m_RenderDataCorrect = False
    
End Property

'The caller can suspend automatic redraws caused by things like adding an item to the list box.  Just make sure to enable redraws
' once you're ready, or you'll never get rendering requests!
Public Sub SetAutomaticRedraws(ByVal newState As Boolean, Optional ByVal raiseRedrawImmediately As Boolean = False)
    m_RedrawAutomatically = newState
    If raiseRedrawImmediately Then
        If m_RenderDataCorrect Then
            RaiseEvent RedrawNeeded
        Else
            Me.CalculateRenderMetrics
        End If
    End If
End Sub

'Call this sub to request a full redraw of the list.  This sub doesn't actually perform any drawing; instead, it raises a series of
' RenderListItem() events, which the caller can then handle on their own terms.
Public Sub CalculateRenderMetrics()
    
    'Prior to requesting a redraw, determine the indices of the first and last items our owner needs to draw.
    ' We'll cache these, so we don't have to calculate them again (until something changes, at least).
    Dim i As Long
    
    If m_DefaultHeight = 0 Then Exit Sub
    
    'Lists with uniform item sizes can skip a lot of messy handling.
    If m_SizeMode = PDLH_FIXED Then
    
        'Calculate the first overlapping item that overlaps the viewable area (and cache it)
        i = m_ScrollValue \ m_DefaultHeight
        m_FirstRenderIndex = i
        
        'Search for the first item that doesn't overlap the existing container area
        Do
            
            'Move to the next entry in the list
            i = i + 1
            
            'If we're past the number of items in the list, exit immediately
            If i >= m_NumOfItems Then Exit Do
            
        Loop While (i * m_DefaultHeight) < m_ScrollValue + m_ContainerHeight
        
        m_LastRenderIndex = i - 1
        
    
    'If list entries have variable height, we need to loop through all list entries and look for ones that overlap the viewport.
    ' This step could be optimized by using something like a binary search to detect the first overlapping item, but because PD
    ' only uses short lists, I'm not particularly concerned with optimizing this right now.
    Else
        
        'To spare us from looping through unnecessary tail-end entries, if we've already raised at least one draw event, and we
        ' encounter a list item that exists outside the list, we can immediately exit the function.
        Dim oneItemFound As Boolean
        oneItemFound = False
        
        'Hit detection is easily performed using rectangle intersection.
        ' (Note that the .Top parameter must always be set mathematically, later in this function.)
        Dim tmpRect As RECTF
        tmpRect.Left = 0
        tmpRect.Width = m_ContainerWidth
        tmpRect.Height = m_DefaultHeight
        
        'Create a dummy rect at the size of the current viewport.  We'll use this for intersection testing.
        Dim listRect As RECTF
        With listRect
            .Left = 0
            .Width = m_ContainerWidth
            .Top = m_ScrollValue
            .Height = m_ContainerHeight
        End With
        
        Dim intersectionRect As RECTF
        
        For i = 0 To m_NumOfItems - 1
            
            m_FirstRenderIndex = LONG_MAX
            m_LastRenderIndex = -1 * LONG_MAX
            
            'Populate our temporary rect.
            tmpRect.Top = m_Items(i).itemTop
            tmpRect.Height = m_Items(i).itemHeight
            
            'If overlap is found, render the rect
            If GDI_Plus.IntersectRectF(intersectionRect, tmpRect, listRect) Then
                
                If i < m_FirstRenderIndex Then m_FirstRenderIndex = i
                If i > m_LastRenderIndex Then m_LastRenderIndex = i
                
                oneItemFound = True
            
            'If no overlap is found, but overlap was found for a PREVIOUS entry, exit immediately
            Else
                If oneItemFound Then Exit For
            End If
            
        Next i
        
    End If
    
    'Note that our rendering data is up-to-date.  As long as this stays TRUE, we don't have to recalculate rendering data.
    m_RenderDataCorrect = True
    
    If m_RedrawAutomatically Then RaiseEvent RedrawNeeded
    
End Sub
