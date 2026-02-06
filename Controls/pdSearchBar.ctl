VERSION 5.00
Begin VB.UserControl pdSearchBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   ClipControls    =   0   'False
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
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   ToolboxBitmap   =   "pdSearchBar.ctx":0000
   Begin PhotoDemon.pdListBoxOD lbPrimary 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
   End
End
Attribute VB_Name = "pdSearchBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Search Bar control
'Copyright 2019-2026 by Tanner Helland
'Created: 25/April/19
'Last updated: 09/January/24
'Last update: harden against potential errors (see https://github.com/tannerhelland/PhotoDemon/issues/509)
'
'This is PD's version of a "search box" - an edit box that raises a neighboring list window with a list
' of "hits" that match the current search query.  Search matching is left up to the parent window, which
' allows us to drop this control wherever we want without worrying about implementation details.
'
'This control is also very similar in construction to the pdListBox object, including its reliance on a
' separate pdListSupport class for managing its data.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'By design, this textbox raises fewer events than a standard text box
Public Event Change()
Public Event Click(ByRef bestSearchHit As String)
Public Event KeyPress(ByVal vKey As Long, ByRef preventFurtherHandling As Boolean)
Public Event Resize()
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'After language changes, the search list needs to be refreshed (as the current results, if any,
' will still be in the *old* language!)
Public Event RequestSearchList()

'The actual common control edit box is handled by a dedicated class
Private WithEvents m_EditBox As pdEditBoxW
Attribute m_EditBox.VB_VarHelpID = -1

'Some mouse states relative to the edit box are tracked, so we can render custom borders around the embedded box
Private m_MouseOverEditBox As Boolean

'Tracks whether the control (any component) has focus.  This is helpful as we must synchronize between VB's focus events and API
' focus events.  This value is deliberately kept separate from m_HasFocus, above, as we only use this value to raise our own
' Got/Lost focus events when the *entire control* loses focus (vs any one individual component).
Private m_ControlHasFocus As Boolean

'If the user resizes an edit box, the control's back buffer needs to be redrawn.  If we resize the edit box as part of an internal
' AutoSize calculation, however, we will already be in the midst of resizing the backbuffer - so we override the behavior of the
' UserControl_Resize event, using this variable.
Private m_InternalResizeState As Boolean

'All searches are performed against this string stack.  The caller must supply this (obviously).
Private m_SearchStack As pdStringStack

'All search matches are placed into this stack.  The caller can retrieve this and do whatever
' they want with it; we just display it in the dropdown list.
Private m_SearchResults As pdStringStack

'If the list of search matches changes due to user input, we will reset the current dropdown listindex
' to the "best match" string.  If, however, the current list of results has *not* changed (typical
' when typing something like a space character), we will preserve the existing listindex, if any,
' between updates.
Private m_LastResults As pdStringStack, m_ResultsChanged As Boolean

'If the user hits enter (while in the text box) or clicks a specific list entry, we raise a corresponding
' _Click() event and return the clicked (or in the case of Enter presses, best-matched) string.  That string
' is also cached locally, and can be manually retrieved for other purposes.
Private m_BestMatchString As String

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDDROPDOWNFONT_COLOR_LIST
    [_First] = 0
    PDDD_Background = 0
    PDDD_ComboFill = 1
    PDDD_ComboBorder = 2
    PDDD_DropDownCaption = 3
    PDDD_DropArrow = 4
    PDDD_ListCaption = 5
    PDDD_ListBorder = 6
    [_Last] = 6
    [_Count] = 7
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'Padding distance (in px) between the user control edges and the edit box edges
Private Const EDITBOX_BORDER_PADDING As Long = 2&

'Everything below this line comes from pdDropDown, and should be (hypothetically) tied to raising the
' dynamic list box.
Private Declare Function GetWindowRect Lib "user32" (ByVal srcHWnd As Long, ByRef dstRectL As RectL) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal ptrToRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Private m_WindowStyleHasBeenSet As Boolean
Private m_OriginalWindowBits As Long, m_OriginalWindowBitsEx As Long
Private m_popupRectCopy As RectL

Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20

'When the popup listbox is raised, we subclass the parent control.  If it is moved or sized or clicked, we automatically
' unload the dropdown listview.  (This workaround is necessary for modal dialogs, among other things.)
Implements ISubclass
Private m_ParentHWnd As Long
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_WINDOWPOSCHANGING As Long = &H46&

'Change this value to control the maximum number of visible items in the dropped box.  (Note that it's technically
' this value + 1, with the +1 representing the currently selected item.)
Private Const NUM_ITEMS_VISIBLE As Long = 8

'Padding around the currently selected list item when painted to the combo box.  These values are also added to the
' default font metrics to arrive at a default control size.
Private Const LIST_PADDING_HORIZONTAL As Single = 4!
Private Const LIST_PADDING_VERTICAL As Single = 2!

'The rectangle where the combo portion of the control is actually rendered
Private m_ComboRect As RectF

'When the popup listbox is visible, this is set to TRUE.  (Also, as a failsafe the list box hWnd is cached.)
Private m_PopUpVisible As Boolean, m_PopUpHwnd As Long

'If we previously pushed the edit box (awkwardly) to the side so that extra-long entries fit on screen,
' this will be set to TRUE.  As the dropdown list tends to shrink as the user types more entries,
' we want to reorient the list box as necessary.
Private m_PopUpForciblyFit As Boolean

'List box support class.  Handles data storage and coordinate math for rendering, but for this control, we primarily
' use the data storage aspect.  (Note that when the combo box is clicked and the corresponding listbox window is raised,
' we hand a copy of this class over to the list view so it can clone it and mirror our data.)
Private WithEvents listSupport As pdListSupport
Attribute listSupport.VB_VarHelpID = -1

'If something forces us to release our subclass while in the midst of the subclass proc, we want to delay the request until
' the subclass exits.  If we don't do this, PD will crash.
Private m_InSubclassNow As Boolean, m_SubclassActive As Boolean
Private WithEvents m_SubclassReleaseTimer As pdTimer
Attribute m_SubclassReleaseTimer.VB_VarHelpID = -1

'Persistent brush and pen objects for rendering list elements
Private m_Brush As pd2DBrush, m_Pen As pd2DPen

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_SearchBar
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    If Not (m_EditBox Is Nothing) Then
        m_EditBox.Enabled = newValue
        RelayUpdatedColorsToEditBox
    End If
    UserControl.Enabled = newValue
    If PDMain.IsProgramRunning() Then RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Public Property Get FontSize() As Single
    If (Not m_EditBox Is Nothing) Then FontSize = m_EditBox.FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If Not (m_EditBox Is Nothing) Then
        If (newSize <> m_EditBox.FontSize) Then
            m_EditBox.FontSize = newSize
            listSupport.DefaultItemHeight = Fonts.GetDefaultStringHeight(newSize) + Interface.FixDPI(LIST_PADDING_VERTICAL) * 2
            lbPrimary.ListItemHeight = listSupport.DefaultItemHeight
            PropertyChanged "FontSize"
        End If
    End If
End Property

Public Property Get HasFocus() As Boolean
    HasFocus = ucSupport.DoIHaveFocus() Or m_EditBox.HasFocus()
End Property

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

Public Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestNewSize newWidth, newHeight, True
End Sub

'External functions can call this to set focus to the search box
Public Sub SetFocusToSearchBox()
    m_EditBox.SetFocusToEditBox
    m_EditBox.SelectAll
End Sub

'External functions can call this to fully select the text box's contents
Public Sub SelectAll()
    If (Not m_EditBox Is Nothing) Then m_EditBox.SelectAll
End Sub

'SelStart is used by some PD functions to control caret positioning after automatic text updates (as used in the text up/down)
Public Property Get SelStart() As Long
    If (Not m_EditBox Is Nothing) Then SelStart = m_EditBox.SelStart
End Property

Public Property Let SelStart(ByVal newPosition As Long)
    If (Not m_EditBox Is Nothing) Then m_EditBox.SelStart = newPosition
End Property

Public Property Get Text() As String
    If (Not m_EditBox Is Nothing) Then Text = m_EditBox.Text
End Property

Public Property Let Text(ByRef newString As String)
    If (Not m_EditBox Is Nothing) Then
        m_EditBox.Text = newString
        If PDMain.IsProgramRunning() Then
            RaiseEvent Change
        Else
            PropertyChanged "Text"
        End If
    End If
End Property

'You *MUST* call this before the user starts typing; otherwise, the control won't have anything to search!
Public Sub SetSearchList(ByRef srcStringStack As pdStringStack)
    Set m_SearchStack = srcStringStack
    If m_PopUpVisible Then RefreshSearchResults
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    m_InSubclassNow = True
    
    'If certain events occur in our parent window, and our list box is visible, release it
    If m_PopUpVisible Then
        If (uiMsg = WM_ENTERSIZEMOVE) Or (uiMsg = WM_WINDOWPOSCHANGING) Then
            HideListBox
        ElseIf (uiMsg = WM_LBUTTONDOWN) Or (uiMsg = WM_RBUTTONDOWN) Or (uiMsg = WM_MBUTTONDOWN) Then
            HideListBox
        ElseIf (uiMsg = WM_NCDESTROY) Then
            HideListBox
            Set m_SubclassReleaseTimer = Nothing
            VBHacks.StopSubclassing hWnd, Me
            m_ParentHWnd = 0
        End If
    End If
    
    'Never eat parent window messages; just peek at them
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    m_InSubclassNow = False
    
End Function

'On dropdown-click, return the string (*not* the index) that was clicked
Private Sub lbPrimary_Click()
    m_BestMatchString = lbPrimary.List(lbPrimary.ListIndex)
    RaiseEvent Click(m_BestMatchString)
End Sub

Private Sub lbPrimary_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    If (m_SearchResults Is Nothing) Then Exit Sub
    
    'Cache colors in advance, so we can simply reuse them in the inner loop
    Dim itemFillColor As Long, itemFillBorderColor As Long, itemFontColor As Long
    itemFillColor = m_Colors.RetrieveColor(PDDD_ComboFill, Me.Enabled, itemIsSelected, itemIsHovered)
    itemFillBorderColor = m_Colors.RetrieveColor(PDDD_ListBorder, Me.Enabled, itemIsSelected, itemIsHovered)
    itemFontColor = m_Colors.RetrieveColor(PDDD_ListCaption, Me.Enabled, itemIsSelected, itemIsHovered)
    
    'Grab the rendering rect
    Dim tmpRectF As RectF
    If (ptrToRectF <> 0) Then CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    'Paint the fill and border
    Dim cSurface As pd2DSurface: Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundDC bufferDC
    cSurface.SetSurfaceCompositing P2_CM_Overwrite
    
    If (m_Brush Is Nothing) Then Set m_Brush = New pd2DBrush
    m_Brush.SetBrushColor itemFillColor
    PD2D.FillRectangleF_FromRectF cSurface, m_Brush, tmpRectF
    
    If (m_Pen Is Nothing) Then
        Set m_Pen = New pd2DPen
        m_Pen.SetPenLineJoin P2_LJ_Miter
    End If
    m_Pen.SetPenColor itemFillBorderColor
    PD2D.DrawRectangleF_FromRectF cSurface, m_Pen, tmpRectF
    Set cSurface = Nothing
    
    'Paint the font name in the default UI font
    Dim tmpFont As pdFont, textPadding As Single
    Set tmpFont = Fonts.GetMatchingUIFont(Me.FontSize)
    textPadding = LIST_PADDING_HORIZONTAL
    
    Dim tmpString As String
    tmpString = m_SearchResults.GetString(itemIndex)
    
    'Failsafe only
    If (LenB(tmpString) > 0) Then
        tmpFont.SetFontColor itemFontColor
        tmpFont.AttachToDC bufferDC
        tmpFont.SetTextAlignment vbLeftJustify
        tmpFont.FastRenderTextWithClipping tmpRectF.Left + textPadding, tmpRectF.Top + LIST_PADDING_VERTICAL, tmpRectF.Width - LIST_PADDING_HORIZONTAL, tmpRectF.Height - LIST_PADDING_VERTICAL, tmpString, False, True, False
        tmpFont.ReleaseFromDC
    End If
    
End Sub

Private Sub lbPrimary_GotFocusAPI()
    ComponentGotFocus
End Sub

Private Sub lbPrimary_LostFocusAPI()
    ComponentLostFocus
End Sub

Private Sub listSupport_RedrawNeeded()
    If ucSupport.AmIVisible Then RedrawBackBuffer
End Sub

Private Sub m_EditBox_Change()

    If (PDMain.IsProgramRunning()) Then
        
        RefreshSearchResults
        
        'Finally, notify our parent
        RaiseEvent Change
        
    End If
    
End Sub

Private Sub m_EditBox_GotFocusAPI()
    
    'If the dropdown isn't visible, make it visible now
    If (Not m_ControlHasFocus) And (Not m_SearchResults Is Nothing) Then
        If (LenB(m_EditBox.Text) <> 0) And (m_SearchResults.GetNumOfStrings > 0) Then RaiseListBox
    End If
    
    ComponentGotFocus
    
End Sub

Private Sub m_EditBox_KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)

    'Many edit boxes defer to PD's central hotkey handler for Ctrl+A; the search bar, however,
    ' is one where we definitely want to handle Ctrl+A ourselves.
    If ((vKey = vbKeyA) And (Shift = vbCtrlMask)) Then
        m_EditBox.SelectAll
        preventFurtherHandling = True
    End If
    
End Sub

Private Sub m_EditBox_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    'Enter raises a Click event with the current best-match search result (if any)
    If (vKey = pdnk_Enter) Then
        
        'Make sure we have usable search results for the current query
        PerformSearch
        If (m_SearchResults Is Nothing) Then
            preventFurtherHandling = True
            Exit Sub
        End If
        
        If (m_SearchResults.GetNumOfStrings > 0) Then
            
            'If the list box is dropped, query it for a list index; if the user has used arrow keys
            ' to select a *different* entry, use that instead.  (This allows for navigating the dropdown
            ' using keyboard only, no mouse required.)
            Dim targetIndex As Long: targetIndex = 0
            If m_PopUpVisible Then
                If (lbPrimary.ListIndex >= 0) Then targetIndex = lbPrimary.ListIndex
            End If
            
            m_BestMatchString = m_SearchResults.GetString(targetIndex)
            RaiseEvent Click(m_BestMatchString)
            
        Else
            If (Not NavKey.NotifyNavKeypress(Me, vKey, Shift)) Then RaiseEvent KeyPress(vKey, preventFurtherHandling)
        End If
    
    'Esc/Tab keypresses are checked for navigation usefulness
    ElseIf ((vKey = pdnk_Escape) Or (vKey = pdnk_Tab)) Then
        If (Not NavKey.NotifyNavKeypress(Me, vKey, Shift)) Then RaiseEvent KeyPress(vKey, preventFurtherHandling)
    
    'Up/down keypresses are forwarded to the dropped listbox
    ElseIf (((vKey = vbKeyDown) Or (vKey = vbKeyUp)) And m_PopUpVisible) Then
        lbPrimary.NotifyKeyDown Shift, vKey, preventFurtherHandling
    
    'Other keypresses are passed, uninterrupted, to our parent
    Else
        RaiseEvent KeyPress(vKey, preventFurtherHandling)
    End If
    
End Sub

Private Sub m_EditBox_LostFocusAPI()
    ComponentLostFocus
End Sub

Private Sub m_EditBox_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverEditBox = True
    RedrawBackBuffer
End Sub

Private Sub m_EditBox_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverEditBox = False
    RedrawBackBuffer
End Sub

Private Sub m_SubclassReleaseTimer_Timer()
    If (Not m_InSubclassNow) Then
        m_SubclassReleaseTimer.StopTimer
        RemoveSubclass
    End If
End Sub

Private Sub ucSupport_LostFocusAPI()
    If m_PopUpVisible Then HideListBox
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo And (Not m_InternalResizeState) Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    
    If (Not m_EditBox Is Nothing) Then
        If (m_EditBox.hWnd = 0) Then CreateEditBoxAPIWindow
        m_EditBox.Visible = newVisibility
    Else
        If newVisibility Then
            listSupport.SetAutomaticRedraws True, True
        Else
            If m_PopUpVisible Then HideListBox
        End If
    End If
    
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    RaiseEvent Resize
End Sub

'Call this to perform a search against the current search object.  Should generally be called
' on all edit box _Change() events.
Private Sub PerformSearch()
    
    'Errors are not expected; this is purely a failsafe against unexpected localization surprises
    On Error GoTo BadSearch
    
    'Failsafe checks to make sure the caller gave us something to search
    If (m_SearchStack Is Nothing) Or (m_EditBox Is Nothing) Then
        Set m_SearchResults = Nothing
        Exit Sub
    End If
    
    If (m_SearchStack.GetNumOfStrings <= 0) Then
        Set m_SearchResults = Nothing
        Exit Sub
    End If
    
    Dim strSource As String
    strSource = Trim$(m_EditBox.Text)
    
    If (LenB(strSource) = 0) Then
        Set m_SearchResults = Nothing
        Exit Sub
    End If
    
    Set m_SearchResults = New pdStringStack
    
    'If the source search string contains multiple words, split it into individual words before continuing.
    ' (NOTE: this uses the space char as a word delimiter, which isn't correct across all locales -
    '  e.g. languages like Thai will not be covered by this, and a more comprehensive library like Uniscribe
    '  should really be used for word-breaking.  TODO!)
    Dim lstSearchTerms As pdStringStack
    Set lstSearchTerms = Strings.GetListOfWordsFromString(strSource)
    
    'Remove any menu separator "words" as these just add noise to the results list (e.g. ">")
    Dim i As Long, j As Long
    Dim tmpSearchList As pdStringStack
    Set tmpSearchList = New pdStringStack
    Do While lstSearchTerms.PopString(strSource)
        strSource = Trim$(strSource)
        If (LenB(strSource) > 0) Then
            If (strSource <> ">") Then tmpSearchList.AddString strSource
        End If
    Loop
    
    Set lstSearchTerms = tmpSearchList
    
    'Cancel search if there are no useable search terms in the current search string
    If (lstSearchTerms.GetNumOfStrings <= 0) Then Exit Sub
    
    'Search is pretty simple: iterate the list the caller provided, and see if the search string occurs
    ' inside any of the strings we were passed.  Exact matches are given priority over partial matches,
    ' but otherwise, no special search "ranking" is currently performed.
    Dim alreadyAdded() As Boolean
    ReDim alreadyAdded(0 To m_SearchStack.GetNumOfStrings - 1) As Boolean
    
    'First, look for an exact match of the given search term.
    Dim testString As String
    For i = 0 To m_SearchStack.GetNumOfStrings - 1
        testString = m_SearchStack.GetString(i)
        If Strings.StringsEqual(strSource, testString, True) Then
            If (LenB(testString) > 0) Then
                m_SearchResults.AddString m_SearchStack.GetString(i)
                alreadyAdded(i) = True
            End If
        End If
    Next i
    
    'Next, look for partial matches of one or more words in the search list.
    Dim curHits As Long, maxHits As Long
        
    'Perform a first pass to see if we get *any* hits for *any* of the search terms.
    For i = 0 To m_SearchStack.GetNumOfStrings - 1
        
        'Skip already added items
        If (Not alreadyAdded(i)) Then
            
            curHits = 0
            
            'Iterate all separate search words, and count how many hits we get
            testString = m_SearchStack.GetString(i)
            For j = 0 To lstSearchTerms.GetNumOfStrings - 1
                If (InStr(1, testString, lstSearchTerms.GetString(j), vbTextCompare) <> 0) Then curHits = curHits + 1
            Next j
            
            If (curHits > maxHits) Then maxHits = curHits
            
        End If
        
    Next i
    
    'If any partial matches were found, we now want to add them to the search results queue -
    ' but IMPORTANTLY, we want to add them according to *how many* partial matches were found.
    ' (e.g. if the user searches for "bright contrast", we want to return
    ' "Adjustments > Brightness and Contrast"
    ' ...ahead of...
    ' "Auto-correct contrast"
    If (maxHits > 0) Then
        
        'As part of this segment of the search engine, we want to put results in one of three
        ' categories: "good", "better", and "best" results.
        
        ' - "Best" results start with the text of the search query.
        ' - "Better" results do not start with the text of the search query, but a word within
        '   the search text *does* start with the search query.
        ' - "Good" results are results that include any words in the search query.
        
        ' As a concrete example, if the user searches for "gra"...
        ' - "Gradient tool" is a best result (it starts with "gra")
        ' - "Monochrome to grayscale" is a better result (a word other than the first one
        '   starts with "gra")
        ' - "Histogram" is just a good result (because "gra" appears in the text, but it's in
        '   the middle of a word rather than the start).
        
        'After sorting results into these categories, we will append results in "best",
        ' "better", "good" order.
        Dim goodResults As pdStringStack: Set goodResults = New pdStringStack
        Dim betterResults As pdStringStack: Set betterResults = New pdStringStack
        Dim bestResults As pdStringStack: Set bestResults = New pdStringStack
        
        Dim loopHits As Long, minHits As Long, hitPosition As Long
        Dim resultIsBest As Boolean, resultIsBetter As Boolean
        
        'If the search query contains (n) words, we want to match at least (n-1) of them.
        ' This is useful after the user performs a search query, as the edit box will read something
        ' like "Effects > Stylize > Vignetting".  If we allow one-word matches, everything in the
        ' "Effects" menu will appear in the results box - but using an n-1 system, we will first
        ' return the exact match "Effects > Stylize > Vignetting", followed by all other items in
        ' the "Effects > Stylize" submenu - but no other items from the base "Effects" menu.
        minHits = 1
        If (maxHits > 2) Then minHits = maxHits - 1
        
        For loopHits = maxHits To minHits Step -1
            For i = 0 To m_SearchStack.GetNumOfStrings - 1
                If (Not alreadyAdded(i)) Then
                    
                    curHits = 0
                    resultIsBetter = False
                    resultIsBest = False
                    
                    'Cache the target search string for improved performance, and perform additional failsafe checks
                    ' (just in case).
                    testString = m_SearchStack.GetString(i)
                    If (LenB(testString) > 0) Then
                        
                        For j = 0 To lstSearchTerms.GetNumOfStrings - 1
                            
                            hitPosition = InStr(1, testString, lstSearchTerms.GetString(j), vbTextCompare)
                            If (hitPosition <> 0) Then
                                
                                curHits = curHits + 1
                                
                                'Classify this as a "good" or "better" result based on where in the string the match
                                ' was found (remember: start-of-first-word is prioritized over start-of-later-word
                                ' which is prioritized over match-in-middle-of-word).
                                If (hitPosition = 1) Then
                                    resultIsBest = True
                                Else
                                    Const SPACE_CHAR As String = " "
                                    If (Mid$(testString, hitPosition - 1, 1) = SPACE_CHAR) Then resultIsBetter = True
                                End If
                                
                            End If
                            
                        Next j
                        
                        If (curHits = loopHits) Then
                            If resultIsBest And (curHits = maxHits) Then
                                bestResults.AddString testString
                            ElseIf resultIsBest Or resultIsBetter Then
                                betterResults.AddString testString
                            Else
                                goodResults.AddString testString
                            End If
                            alreadyAdded(i) = True
                        End If
                        
                    '/Failsafe length check on testString (m_SearchStack.GetString(i))
                    End If
                        
                End If
            Next i
        Next loopHits
        
        If (bestResults.GetNumOfStrings > 0) Then m_SearchResults.AppendStack bestResults
        If (betterResults.GetNumOfStrings > 0) Then m_SearchResults.AppendStack betterResults
        If (goodResults.GetNumOfStrings > 0) Then m_SearchResults.AppendStack goodResults
        
    End If
    
    'Other matching mechanisms could be performed here in the future (e.g. phonetic algorithms), but they
    ' are not implemented currently as internationalization concerns terrify me.  English searches would
    ' be easy enough to handle, but other languages... I'd definitely need outside help.
    
BadSearch:
    
End Sub

'After search results change, we need to update the corresponding list object
Private Sub UpdateResultsList()

    listSupport.Clear
    If (m_SearchResults Is Nothing) Then Exit Sub
    If (m_SearchResults.GetNumOfStrings > 0) Then
        
        'Update each item in the list, and - importantly! - note that items do *NOT* need to
        ' be translated by the language engine (as we've already received translated strings
        ' from the menu manager
        Dim i As Long
        For i = 0 To m_SearchResults.GetNumOfStrings - 1
            listSupport.AddItem m_SearchResults.GetString(i), itemShouldBeTranslated:=False
        Next i
        
    End If
    
End Sub

Private Sub RefreshSearchResults()
    
    If PDMain.IsProgramRunning() Then
    
        'First, perform a search to see if we have any matches
        PerformSearch
        
        'Compare the old and new search results list to see if any changes were made;
        ' this affects how we assign a list index in the dropdown window
        m_ResultsChanged = True
        
        'Ensure at least 1 search hit exists
        If (Not m_SearchResults Is Nothing) Then
            If (m_SearchResults.GetNumOfStrings > 0) Then
                
                'Only perform a comparison against previous results if we actually have a previous result to compare
                If (Not m_LastResults Is Nothing) Then
                    If (m_LastResults.GetNumOfStrings = m_SearchResults.GetNumOfStrings) Then
                        
                        'Compare lists for equality
                        Dim i As Long, mismatchFound As Boolean
                        For i = 0 To m_LastResults.GetNumOfStrings - 1
                            If Strings.StringsNotEqual(m_LastResults.GetString(i), m_SearchResults.GetString(i)) Then
                                mismatchFound = True
                                Exit For
                            End If
                        Next i
                        
                        m_ResultsChanged = mismatchFound
                        
                    End If
                End If
            End If
        End If
        
        'Make a backup copy of the current search results list; we use this in the previous step
        ' to detect changes to the current list of search results (which again, affects how we assign
        ' a listindex in the dropdown window - if the list of results changes, we default to position 0,
        ' or the "best match", but if the list of results *hasn't* changed, we preserve the user's
        ' current selection, if any)
        Set m_LastResults = New pdStringStack
        If (Not m_SearchResults Is Nothing) Then m_LastResults.CloneStack m_SearchResults
        
        'If we do, forward the matches to the listbox and display it
        UpdateResultsList
        
        'If search results exist, raise the list box; otherwise, hide it unconditionally
        If (m_SearchResults Is Nothing) Then
            HideListBox
            Exit Sub
        Else
            If (m_SearchResults.GetNumOfStrings > 0) Then RaiseListBox Else HideListBox
        End If
        
    End If
        
End Sub

'Sometimes, we want to change the UC's size to match the edit box.  Other times, we want to change the edit box's size to
' match the UC.  Use this two functions to update the appropriate size; if "editBoxGetsMoved" is TRUE, we'll forcibly set
' it to match our desired size.
Private Sub SynchronizeSizes()
    
    If (Not m_EditBox Is Nothing) Then
        
        Dim needToMove As Boolean
        needToMove = True
        
        'Start by determining a rect that we ideally want the edit box to fit within.  Note that x2 and y2 in this measurement
        ' are RIGHT AND BOTTOM measurements, not WIDTH AND HEIGHT.
        Dim tmpRect As winRect
        CalculateDesiredEditBoxRect tmpRect
        
        'Next, retrieve the edit box's current rect.  If it's already in an ideal position, skip the move step entirely.
        Dim curRect As winRect
        If m_EditBox.GetPositionRect(curRect) Then
            
            If (tmpRect.x1 = curRect.x1) And (tmpRect.x2 = curRect.x2) And (tmpRect.y1 = curRect.y1) And (tmpRect.y2 = curRect.y2) Then
                needToMove = False
            End If
            
        End If
        
        'Apply the move conditionally
        If needToMove Then
            m_InternalResizeState = True
            m_EditBox.Move tmpRect.x1, tmpRect.y1, tmpRect.x2 - tmpRect.x1, tmpRect.y2 - tmpRect.y1
            m_InternalResizeState = False
        End If
        
    End If
    
End Sub

'When one of this control's components (either the underlying UC or the edit box) gets focus, call this function to update
' trackers and UI accordingly.
Private Sub ComponentGotFocus()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'If a component already had focus, ignore this step, as focus is just changing internally within the control
    If (Not m_ControlHasFocus) Then
        m_ControlHasFocus = True
        RaiseEvent GotFocusAPI
    End If
    
    'The user control itself should never have focus.  Forward it to the API edit box as necessary.
    If (Not m_EditBox Is Nothing) Then
        If (Not m_EditBox.HasFocus) Then m_EditBox.SetFocusToEditBox
    End If
    
    'Regardless of component state, redraw the control "just in case"
    RelayUpdatedColorsToEditBox
    RedrawBackBuffer
    
End Sub

'When one of this control's components (either the underlying UC or the edit box) loses focus, call this function to update
' trackers and UI accordingly.
Private Sub ComponentLostFocus()
    
    'If focus has simply moved to another component within the control, ignore this step
    If m_ControlHasFocus And Not ucSupport.DoIHaveFocus Then
        If (Not m_EditBox Is Nothing) Then
            If (Not m_EditBox.HasFocus) Then
                m_ControlHasFocus = False
                RaiseEvent LostFocusAPI
            End If
        End If
    End If
    
    If (Not m_ControlHasFocus) Then HideListBox
    
    'Regardless of component state, redraw the control "just in case"
    RelayUpdatedColorsToEditBox
    RedrawBackBuffer
    
End Sub

Private Sub CalculateDesiredEditBoxRect(ByRef targetRect As winRect)
    With targetRect
        .x1 = EDITBOX_BORDER_PADDING
        .y1 = EDITBOX_BORDER_PADDING
        .x2 = ucSupport.GetControlWidth - EDITBOX_BORDER_PADDING
        .y2 = ucSupport.GetControlHeight - EDITBOX_BORDER_PADDING
    End With
End Sub

Public Function PixelWidth() As Long
    PixelWidth = ucSupport.GetControlWidth
End Function

Public Function PixelHeight() As Long
    PixelHeight = ucSupport.GetControlHeight
End Function

'Generally speaking, the underlying API edit box management class recreates itself as needed, but we need to request its
' initial creation.  During this stage, we also auto-size ourself to match the edit box's suggested size.
Private Sub CreateEditBoxAPIWindow()
    
    If Not (m_EditBox Is Nothing) Then
        
        Dim tmpRect As winRect
        
        'Make sure all edit box settings are up-to-date prior to creation
        m_EditBox.Enabled = Me.Enabled
        RelayUpdatedColorsToEditBox
        
        'Resize ourselves vertically to match the edit box's suggested size.
        m_InternalResizeState = True
        ucSupport.RequestNewSize ucSupport.GetControlWidth, m_EditBox.SuggestedHeight + EDITBOX_BORDER_PADDING * 2, False
        m_InternalResizeState = False
        
        'Now that we're the proper size, determine where we're gonna stick the edit box (relative to this control instance)
        CalculateDesiredEditBoxRect tmpRect
        
        'Ask the edit box to create itself!
        m_EditBox.CreateEditBox UserControl.hWnd, tmpRect.x1, tmpRect.y1, tmpRect.x2 - tmpRect.x1, tmpRect.y2 - tmpRect.y1, False
        
        'Because control sizes may have changed, we need to repaint everything
        RedrawBackBuffer
        
        'Creating the edit box may have caused this control to resize itself, so as a failsafe, raise a
        ' Resize() event manually
        RaiseEvent Resize
    
    End If
    
End Sub

Private Sub UserControl_Hide()
    If (Not m_EditBox Is Nothing) Then m_EditBox.Visible = False
End Sub

Private Sub UserControl_Initialize()
    
    'Note that we are not currently responsible for any resize events
    m_InternalResizeState = False
    
    'Initialize an edit box support class
    Set m_EditBox = New pdEditBoxW
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDDROPDOWNFONT_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDDropDownFont", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Initialize a helper list class; it manages the actual list data, and a bunch of rendering and layout decisions
    Set listSupport = New pdListSupport
    listSupport.SetAutomaticRedraws False
    listSupport.ListSupportMode = PDLM_DropDown
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    FontSize = 10
    Text = vbNullString
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
        Text = .ReadProperty("Text", vbNullString)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Terminate()
    
    Set m_EditBox = Nothing
    
    'As a failsafe, immediately release the popup box.  (If we don't do this, PD will crash.)
    If m_PopUpVisible Then HideListBox
    If Not (m_SubclassReleaseTimer Is Nothing) Then m_SubclassReleaseTimer.StopTimer
    SafelyRemoveSubclass
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", Me.FontSize, 10
        .WriteProperty "Text", Me.Text, vbNullString
    End With
End Sub

Private Sub UpdateControlLayout()
    
    SynchronizeSizes
    
    'Notify the list manager of our new size.  (Note that this isn't necessary from a rendering standpoint, as we don't
    ' render a normal list-type UI to the dropdown - but the listSupport class won't raise Redraw events if it has an
    ' invalid rendering rect.)
    listSupport.NotifyParentRectF m_ComboRect
    
    RedrawBackBuffer
    
End Sub

'After the back buffer has been correctly sized and positioned, this function handles the actual painting.  Similarly, for state changes
' that don't require a resize (e.g. gain/lose focus), this function should be used.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDDD_Background, Me.Enabled, m_ControlHasFocus, m_MouseOverEditBox))
    If (bufferDC = 0) Then Exit Sub
    
    'This control's render code relies on GDI+ exclusively, so there's no point calling it in the IDE - sorry!
    If PDMain.IsProgramRunning() Then
    
        'Relay any recently changed/modified colors to the edit box, so it can repaint itself to match
        RelayUpdatedColorsToEditBox
        
        'Retrieve DPI-aware control dimensions from the support class
        Dim bWidth As Long, bHeight As Long
        bWidth = ucSupport.GetBackBufferWidth
        bHeight = ucSupport.GetBackBufferHeight
        
        'The edit box doesn't actually have a border; we render a pseudo-border onto the underlying UC, as necessary.
        Dim halfPadding As Long
        halfPadding = 1
        
        Dim borderWidth As Single
        If Not (m_EditBox Is Nothing) Then
            If m_EditBox.HasFocus Or m_MouseOverEditBox Then borderWidth = 3! Else borderWidth = 1!
        Else
            borderWidth = 1!
        End If
        
        Dim cSurface As pd2DSurface, cPen As pd2DPen
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC bufferDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfaceCompositing P2_CM_Overwrite
        
        Set cPen = New pd2DPen
        cPen.SetPenWidth borderWidth
        cPen.SetPenColor m_Colors.RetrieveColor(PDDD_ComboBorder, Me.Enabled, m_ControlHasFocus, m_MouseOverEditBox)
        cPen.SetPenLineJoin P2_LJ_Miter
        
        PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, halfPadding, halfPadding, (bWidth - 1) - halfPadding, (bHeight - 1) - halfPadding
        Set cSurface = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    If (Not PDMain.IsProgramRunning()) Then UserControl.Refresh
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
        
    'Color list retrieval is pretty darn easy - just load each color one at a time, and leave the rest to the color class.
    ' It will build an internal hash table of the colors we request, which makes rendering much faster.
    With m_Colors
        .LoadThemeColor PDDD_Background, "Background", IDE_WHITE
        .LoadThemeColor PDDD_ComboFill, "ComboFill", IDE_WHITE
        .LoadThemeColor PDDD_ComboBorder, "ComboBorder", IDE_GRAY
        .LoadThemeColor PDDD_DropDownCaption, "Caption", IDE_GRAY
        .LoadThemeColor PDDD_DropArrow, "DropArrow", IDE_GRAY
        .LoadThemeColor PDDD_ListCaption, "ListCaption", IDE_GRAY
        .LoadThemeColor PDDD_ListBorder, "ListBorder", IDE_GRAY
    End With
    
    RelayUpdatedColorsToEditBox
    
End Sub

'When this control has special knowledge of a state change that affects the edit box's visual appearance, call this function.
' It will relay the relevant themed colors to the edit box class.
Private Sub RelayUpdatedColorsToEditBox()
    If (Not m_EditBox Is Nothing) Then
        m_EditBox.BackColor = m_Colors.RetrieveColor(PDDD_Background, Me.Enabled, m_ControlHasFocus, m_MouseOverEditBox)
        m_EditBox.TextColor = m_Colors.RetrieveColor(PDDD_DropDownCaption, Me.Enabled, False, False)
    End If
End Sub

'Display the search results list box; typically done on receiving focus and/or the first edit box "Change" event.
Private Sub RaiseListBox()
    
    On Error GoTo UnexpectedListBoxTrouble
    
    If (Not ucSupport.AmIVisible) Or (Not ucSupport.AmIEnabled) Or (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'This sub is called whenever the list box is meant to be shown (e.g. every edit box change event),
    ' which means there are two possible starting points:
    ' 1) The list box is currently invisible, and needs to be made visible
    ' 2) The list box is already visible, and its size potentially needs to be adjusted (because the
    '     number of search results may have changed, thanks to the user typing more text)
    '
    'Because these two scenarios have quite different requirements, we handle them separately.
    
    'First, retrieve the edit box's window coordinates *in the screen's coordinate space*.
    ' (We need this to know how to position the listbox next to it, and to establish a lower
    ' limit on the dropdown's width.)
    Dim myRect As RectL
    GetWindowRect m_EditBox.hWnd, myRect
    
    'Start with the case where the listbox is *not* visible; this requires more work as we have to establish
    ' an initial x/y position, and also setup a bunch of window bit changes to make the child control
    ' work as a popup window.
    Dim popupRect As RectF, showingFirstTime As Boolean
    
    If (Not m_PopUpVisible) Then
    
        'We now want to figure out the idealized coordinates for the pop-up rect.
        'Start by assuming a successful drop-down.  (If this fails, we'll "drop-up" instead.)
        With popupRect
            .Left = myRect.Left - 3     '-3 so that text aligns (instead of just window chrome)
            .Top = myRect.Bottom + 1
            
            'Width and height are TBD, contingent on list contents
        End With
        
        'The first time we raise the window, we want to cache its current window longs
        ' as whatever VB has set.  (We may need to restore these before the window is unloaded,
        ' to prevent issues with VB.)
        m_PopUpHwnd = lbPrimary.hWnd
        m_ParentHWnd = UserControl.Parent.hWnd
        If (Not m_WindowStyleHasBeenSet) Then
            m_WindowStyleHasBeenSet = True
            m_OriginalWindowBits = g_WindowManager.GetWindowLongWrapper(m_PopUpHwnd)
            m_OriginalWindowBitsEx = g_WindowManager.GetWindowLongWrapper(m_PopUpHwnd, True)
        End If
        
        'We also want to make the listbox a top-level window (SetParent null) and while we're at it,
        ' apply any other relevant window styles.  The top-level window is especially important,
        ' as it allows the listbox to be positioned outside the boundary rect of this control.
        SetParent m_PopUpHwnd, 0&
        g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_EX_PALETTEWINDOW, False, True
        
        'Normally, you need to reset the popup and child flags when you make a window top-level.
        ' Unfortunately, this breaks the window terribly, and I'm not sure why; it's probably an
        ' internal VB thing.  At any rate, the current solution seems to work, so we ignore this for now.
        'g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_CHILD, True, False
        'g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_POPUP, False, False
        
        showingFirstTime = True
        
    Else
        
        showingFirstTime = False
        
        'Initialize our window rect to the list box's *current* on-screen rect.
        Dim tmpWinRect As winRect
        g_WindowManager.GetWindowRect_API m_PopUpHwnd, tmpWinRect
        
        With popupRect
            .Left = tmpWinRect.x1
            .Top = tmpWinRect.y1
            .Width = tmpWinRect.x2 - tmpWinRect.x1
            .Height = tmpWinRect.y2 - tmpWinRect.y1
        End With
        
        'If the edit box was previously forced aside due to screen boundaries, reset it to
        ' the edit box's coordinates; it will be moved again in a subsequent step, as necessary.
        If m_PopUpForciblyFit Then popupRect.Left = myRect.Left
        
    End If
    
    'Regardless of whether the listbox is visible or not, we now want to make sure it is large enough
    ' to fit all search entries (both horizontally and vertically).
    
    'Let's deal with height first.
    
    'Height is obviously contingent on how many entries we need to show.  We want to show as many
    ' as possible, up to the limit of NUM_ITEMS_VISIBLE.  (Past that point, we'll use a scrollbar.)
    ' Note that we calculation both an "untouched" amount to show (amtShowOriginal), and an "actual"
    ' amount to show (amtShow); later in the function we'll use any difference between these to
    ' know if a scrollbar is required.
    Dim amtShow As Long, amtShowOriginal As Long
    amtShow = m_SearchResults.GetNumOfStrings
    amtShowOriginal = amtShow
    If (amtShow > NUM_ITEMS_VISIBLE) Then amtShow = NUM_ITEMS_VISIBLE
    
    'Make sure we're showing at least one item; if we're not, hide the window instead and bail
    If (amtShow <= 0) Then
        HideListBox
        Exit Sub
    End If
        
    'We now know there's at least one item in the results list.
    
    'Instead of doing a cheap size calculation (itemHeight * count), iterate through the list;
    ' this is separator-compatible (if we decide to use separators in the future).
    Dim sizeChange As Single, i As Long
    sizeChange = amtShow * listSupport.DefaultItemHeight
    
    If (listSupport.GetInternalSizeMode = PDLH_Separators) Then
        For i = 0 To amtShow - 1
            If listSupport.DoesItemHaveSeparator(i) Then sizeChange = sizeChange + listSupport.GetSeparatorHeight
        Next i
    End If
    
    'Use this as our baseline for the list box's height, and add enough space for window chrome
    popupRect.Height = sizeChange + 3
    
    'Next, we want to calculate width.  There's no trivial way to do this; instead, we need to
    ' manually iterate the list and find the longest string being displayed.
    
    'Create a temporary DIB so we don't have to constantly re-select the font into a DC of its own making.
    Dim tmpDC As Long
    tmpDC = GDI.GetMemoryDC()
    
    'Font names are rendered in the current UI font
    Dim curFont As pdFont
    Set curFont = Fonts.GetMatchingUIFont(m_EditBox.FontSize)
    curFont.AttachToDC tmpDC
    
    'Find the longest font name
    Dim tmpWidth As Long, maxWidth As Long
    For i = 0 To m_SearchResults.GetNumOfStrings() - 1
        tmpWidth = curFont.GetWidthOfString(m_SearchResults.GetString(i))
        If (tmpWidth > maxWidth) Then maxWidth = tmpWidth
    Next i
    
    curFont.ReleaseFromDC
    GDI.FreeMemoryDC tmpDC
    
    'If the max width is not greater than the width of our parent edit box, use its width instead
    If (myRect.Right - myRect.Left) > maxWidth Then maxWidth = (myRect.Right - myRect.Left)
    popupRect.Width = maxWidth
    
    'If the listbox requires a scroll bar, factor that into the width calculation; otherwise, add just
    ' enough for window chrome padding.
    If (amtShowOriginal > amtShow) Then
        popupRect.Width = popupRect.Width + Interface.FixDPI(28)
    Else
        popupRect.Width = popupRect.Width + Interface.FixDPI(12)
    End If
    
    'We now want to make sure the popup box doesn't lie off-screen.  If it does, we want to flip it
    ' to appear "above" the search bar.
    Dim estimatedDesktopBottom As Long
    estimatedDesktopBottom = (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight) - g_Displays.GetTaskbarHeight
        
    If (popupRect.Top + popupRect.Height > estimatedDesktopBottom) Then popupRect.Top = myRect.Top - popupRect.Height
    
    'Same with left/right differences
    If (popupRect.Left < g_Displays.GetDesktopLeft) Then
        sizeChange = g_Displays.GetDesktopLeft - popupRect.Left
        popupRect.Left = g_Displays.GetDesktopLeft
        m_PopUpForciblyFit = True
    ElseIf (popupRect.Left + popupRect.Width > g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth) Then
        sizeChange = (popupRect.Left + popupRect.Width) - (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth)
        popupRect.Left = popupRect.Left - sizeChange
        m_PopUpForciblyFit = True
    End If
    
    'The list box is now ready to go.
    
    'If the listbox is already visible, compare its newly calculated position to its current position.
    ' If they are identical, we can exit immediately - while still making sure to sync any changes
    ' to the underlying list of search results!
    If (Not showingFirstTime) Then
        
        With m_popupRectCopy
            If (.Left = Int(popupRect.Left)) And (.Top = Int(popupRect.Top)) Then
                If (.Right = Int(popupRect.Left + popupRect.Width + 0.999999)) Then
                    If (.Bottom = Int(popupRect.Top + popupRect.Height + 0.999999)) Then
                        
                        'See if we can reuse the current listindex, if any
                        If m_ResultsChanged Then
                            listSupport.ListIndex = 0
                        ElseIf (lbPrimary.ListIndex >= 0) Then
                            listSupport.ListIndex = listSupport.ListIndexByString(lbPrimary.List(lbPrimary.ListIndex), vbBinaryCompare)
                            If (listSupport.ListIndex < 0) Then listSupport.ListIndex = 0
                        End If
                        
                        lbPrimary.CloneExternalListSupport listSupport, , PDLM_LB_Inside_DD
                        Exit Sub
                        
                    End If
                End If
            End If
        End With
        
    End If
    
    'Move the listbox into its new position *but do not activate it* (we don't want to steal focus
    ' from the edit box where the user is typing!)
    With popupRect
        SetWindowPos m_PopUpHwnd, 0&, .Left, .Top, .Width, .Height, SWP_NOACTIVATE
    End With
    
    'We also need to cache the popup rect's position; when the listbox is closed, we will manually
    ' invalidate windows beneath it (only on certain OS + theme combinations; Aero handles this correctly).
    With m_popupRectCopy
        .Left = Int(popupRect.Left)
        .Top = Int(popupRect.Top)
        .Right = Int(popupRect.Left + popupRect.Width + 0.999999)
        .Bottom = Int(popupRect.Top + popupRect.Height + 0.999999)
    End With
    
    'Clone our list's contents; note that we cannot do this until *after* the list size has been established,
    ' as the scroll bar's maximum value is contingent on the available pixel size of the dropdown.
    
    'Also, while we're here, set a default listindex - this makes it clear what will happen if the user
    ' hits "Enter" inside the edit box
    If showingFirstTime Then
        listSupport.ListIndex = 0
    
    'See if we can reuse the current listindex, if any
    Else
        If m_ResultsChanged Then
            listSupport.ListIndex = 0
        ElseIf (lbPrimary.ListIndex >= 0) Then
            listSupport.ListIndex = listSupport.ListIndexByString(lbPrimary.List(lbPrimary.ListIndex), vbBinaryCompare)
            If (listSupport.ListIndex < 0) Then listSupport.ListIndex = 0
        End If
    End If
    
    lbPrimary.CloneExternalListSupport listSupport, , PDLM_LB_Inside_DD
    
    'Now we can show the window; we also notify the window of its changed window style bits
    With popupRect
        SetWindowPos m_PopUpHwnd, 0&, .Left, .Top, .Width, .Height, SWP_SHOWWINDOW Or SWP_FRAMECHANGED
    End With
    
    'One last thing: because this is a (fairly?  mostly?  extremely?) hackish way to emulate a combo box,
    ' we need to cover the case where the user selects outside the raised list box, but *not* on an object
    ' that can receive focus (e.g. an exposed section of an underlying form).  Focusable objects are taken
    ' care of automatically, because a LostFocus event will fire, but non-focusable clicks are problematic.
    ' To solve this, we subclass our parent control and watch for mouse events. Also, since we're subclassing
    ' the control anyway, we'll also hide the ListBox if the parent window is moved.
    If (m_ParentHWnd <> 0) And PDMain.IsProgramRunning() Then
        
        'Make sure we're not currently trying to release a previous subclass attempt
        Dim subclassActive As Boolean: subclassActive = False
        If Not (m_SubclassReleaseTimer Is Nothing) Then
            If m_SubclassReleaseTimer.IsActive Then
                m_SubclassReleaseTimer.StopTimer
                subclassActive = True
            End If
        End If
        
        If (Not subclassActive) And (Not m_SubclassActive) Then
            VBHacks.StartSubclassing m_ParentHWnd, Me
            m_SubclassActive = True
        End If
        
    End If
    
    'As an additional failsafe, we also notify the central UserControl tracker that a list box is active.
    ' If any other PD control receives focus, that tracker will automatically unload our list box as well,
    ' "just in case".
    UserControls.NotifyDropDownChangeState Me.hWnd, m_PopUpHwnd, True
    
    m_PopUpVisible = True
    
    Exit Sub
    
UnexpectedListBoxTrouble:
    PDDebug.LogAction "WARNING!  pdDropDown.RaiseListBox failed because of Err # " & Err.Number & ", " & Err.Description
    
End Sub

Private Sub HideListBox()

    If m_PopUpVisible And (m_PopUpHwnd <> 0) Then
        
        'Notify the central UserControl tracker that our list box is now inactive.
        UserControls.NotifyDropDownChangeState Me.hWnd, m_PopUpHwnd, False
        
        m_PopUpVisible = False
        SetParent m_PopUpHwnd, Me.hWnd
        If (m_OriginalWindowBits <> 0) Then g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, m_OriginalWindowBits, , , True
        If (m_OriginalWindowBitsEx <> 0) Then g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, m_OriginalWindowBits, , True, True
        g_WindowManager.SetVisibilityByHWnd m_PopUpHwnd, False
        
        m_PopUpHwnd = 0
        
        'If Aero theming is not active, hiding the list box may cause windows beneath the current one to render incorrectly.
        If (OS.IsVistaOrLater And (Not g_WindowManager.IsDWMCompositionEnabled)) Then
            InvalidateRect 0&, VarPtr(m_popupRectCopy), 0&
        End If
        
        'Note that termination may result in the client site not being available.  If this happens, we simply want
        ' to continue; the subclasser will handle clean-up automatically.
        SafelyRemoveSubclass
        
    End If
    
End Sub

'If a hook exists, uninstall it.  DO NOT CALL THIS FUNCTION if the class is currently inside the hook proc.
Private Sub RemoveSubclass()
    On Error GoTo UnsubclassUnnecessary
    If ((m_ParentHWnd <> 0) And m_SubclassActive) Then
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHWnd = 0
        m_SubclassActive = False
    End If
UnsubclassUnnecessary:
End Sub

'Release the edit box's keyboard hook.  In some circumstances, we can't do this immediately, so we set a timer that will
' release the hook as soon as the system allows.
Private Sub SafelyRemoveSubclass()
    If m_InSubclassNow Then
        If (m_SubclassReleaseTimer Is Nothing) Then Set m_SubclassReleaseTimer = New pdTimer
        m_SubclassReleaseTimer.Interval = 16
        m_SubclassReleaseTimer.StartTimer
    Else
        RemoveSubclass
    End If
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        listSupport.UpdateAgainstCurrentTheme
        lbPrimary.UpdateAgainstCurrentTheme
        If PDMain.IsProgramRunning() Then
            NavKey.NotifyControlLoad Me, hostFormhWnd
            ucSupport.UpdateAgainstThemeAndLanguage
            RaiseEvent RequestSearchList
        End If
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    If (Not m_EditBox Is Nothing) Then
        Dim targetHWnd As Long
        If m_EditBox.hWnd = 0 Then targetHWnd = UserControl.hWnd Else targetHWnd = m_EditBox.hWnd
        ucSupport.AssignTooltip targetHWnd, newTooltip, newTooltipTitle
    End If
End Sub
