VERSION 5.00
Begin VB.UserControl pdComboBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFF80&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ClipControls    =   0   'False
   FillColor       =   &H00FF00FF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ToolboxBitmap   =   "pdComboBox.ctx":0000
   Begin VB.Timer tmrHookRelease 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "pdComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Combo Box control
'Copyright ©2013-2014 by Tanner Helland
'Created: 14/November/14
'Last updated: 14/November/14
'Last update: continued work on initial build
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this drop-down (combo) box control, specifically:
'
' 1) Unlike other PD custom controls, this is simply a wrapper to a system combo box.
' 2) The idea with this control was not to expose all combo box properties, but simply those most relevant to PD.
' 3) Focus is the real nightmare for this control, and as you will see, some complicated tricks are required to work
'    around VB's handling of tabstops in particular.
' 4) To allow use of arrow keys and other control keys, this control must hook the keyboard.  (If it does not, VB will
'    eat control keypresses, because it doesn't know about windows created via the API!)  A byproduct of this is that
'    accelerators flat-out WILL NOT WORK while this control has focus.  I haven't yet settled on a good way to handle
'    this; what I may end up doing is manually forwarding any key combinations that use Alt to the default window
'    handler, but I'm not sure this will help.  TODO!
' 5) Dynamic hooking can occasionally cause trouble in the IDE, particularly when used with break points.  It should
'    be fine once compiled.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'By design, this combo box raises fewer events than a standard combo box.  I would prefer the Click() event to actually be Change(),
' but I have used Click() throughout VB due to the behavior of the old combo box - and rather than rewrite all that code, I've simply
' used the same semantics here.  Note, however, that "Click" will also return changes to the combo box that originate from the keyboard.
Public Event Click()

'Flicker-free window painter; note that two painters are (probably? TODO!) required: one for the edit portion of the control (including its button),
' and another for the drop-down list (only the border is relevant here, as individual items draw their own background).
Private WithEvents cPainterBox As pdWindowPainter
Attribute cPainterBox.VB_VarHelpID = -1

'Window styles
Private Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOACTIVATE = &H8000000
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_TOOLWINDOW = &H80&
    WS_EX_MDICHILD = &H40
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

#If False Then
    Private Const WS_BORDER = &H800000, WS_CAPTION = &HC00000, WS_CHILD = &H40000000, WS_CLIPCHILDREN = &H2000000, WS_CLIPSIBLINGS = &H4000000, WS_DISABLED = &H8000000, WS_DLGFRAME = &H400000, WS_EX_ACCEPTFILES = &H10&, WS_EX_DLGMODALFRAME = &H1&, WS_EX_NOPARENTNOTIFY = &H4&, WS_EX_TOPMOST = &H8&, WS_EX_TRANSPARENT = &H20&, WS_EX_TOOLWINDOW = &H80&, WS_GROUP = &H20000, WS_HSCROLL = &H100000, WS_MAXIMIZE = &H1000000, WS_MAXIMIZEBOX = &H10000, WS_MINIMIZE = &H20000000, WS_MINIMIZEBOX = &H20000, WS_OVERLAPPED = &H0&, WS_POPUP = &H80000000, WS_SYSMENU = &H80000, WS_TABSTOP = &H10000, WS_THICKFRAME = &H40000, WS_VISIBLE = &H10000000, WS_VSCROLL = &H200000, WS_EX_MDICHILD = &H40, WS_EX_WINDOWEDGE = &H100, WS_EX_CLIENTEDGE = &H200, WS_EX_CONTEXTHELP = &H400, WS_EX_RIGHT = &H1000, WS_EX_LEFT = &H0, WS_EX_RTLREADING = &H2000, WS_EX_LTRREADING = &H0, WS_EX_LEFTSCROLLBAR = &H4000, WS_EX_RIGHTSCROLLBAR = &H0, WS_EX_CONTROLPARENT = &H10000, WS_EX_STATICEDGE = &H20000, WS_EX_APPWINDOW = &H40000
    Private Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE), WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
#End If

'Updating the font is done via WM_SETFONT
Private Const WM_SETFONT = &H30

'These constants can be used as the second parameter of the ShowWindow API function
Private Enum showWindowOptions
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
End Enum

#If False Then
    Private Const SW_HIDE = 0, SW_SHOWNORMAL = 1, SW_SHOWMINIMIZED = 2, SW_SHOWMAXIMIZED = 3, SW_SHOWNOACTIVATE = 4, SW_SHOW = 5, SW_MINIMIZE = 6, SW_SHOWMINNOACTIVE = 7, SW_SHOWNA = 8, SW_RESTORE = 9, SW_SHOWDEFAULT = 10, SW_FORCEMINIMIZE = 11
#End If

'System window handling APIs
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As RECTL) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As RECTL) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As RECTL, ByVal bErase As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hndWindow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hndWindow As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As showWindowOptions) As Long

'Getting/setting window data
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpStringPointer As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Handle to the system combo box wrapped by this control
Private m_ComboBoxHwnd As Long

'pdFont handles the creation and maintenance of the font used to render the combo box.  It is also used to determine control
' height, as the control is auto-sized to fit the current font.
Private curFont As pdFont

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single

'Alt-key state can be tracked a number of different ways, but I find GetAsyncKeyState to be the easiest.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_KEYUP As Long = &H101
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COMMAND As Long = &H111
Private Const WM_NEXTDLGCTL As Long = &H28
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_SIZE As Long = &H5
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202

Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ALT As Long = &H12    'Note that VK_ALT is referred to as VK_MENU in MSDN documentation!

'Obviously, we're going to be doing a lot of subclassing inside this control.
Private cSubclass As cSelfSubHookCallback

'Mouse input handler helps with things like enter/leave behavior
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cMouseEventsListBox As pdInputMouse
Attribute cMouseEventsListBox.VB_VarHelpID = -1

'Unlike the combo box, which may be recreated multiple times as properties change, we only need to subclass the parent window once.
' After it has been subclassed, this will be set to TRUE.
Private m_ParentHasBeenSubclassed As Boolean

'Dynamic hooking requires us to track focus events with care.  When focus is lost, we must relinquish control of the keyboard.
' This value will be set to TRUE if the API edit box currently has focus.
Private m_HasFocus As Boolean

'Because our API combo box is not tied into VB's default tab stop handling, we must jump through some hoops to forward focus correctly.
' Our hook proc will capture the Tab key that causes focus to enter the control, but mistakenly assume it is a Tab keypress from
' *within* the control.  To prevent this from happening, we enforce a slight time delay from when our hook procedure begins, to when
' we capture Tab keypresses.  This prevents faulty Tab-key handling.
Private m_TimeAtFocusEnter As Long
Private m_FocusDirection As Long

'If the user resizes a combo box, the control's back buffer needs to be redrawn.  If we resize the combo box as part of an internal
' AutoSize calculation, however, we will already be in the midst of resizing the backbuffer - so we override the behavior of the
' UserControl_Resize event, using this variable.
Private m_InternalResizeState As Boolean

'The system handles drawing of the combo box.  This persistent brush handle is returned to the relevant window message,
' and WAPI uses it to draw the control's background.
Private m_ComboBoxBrush As Long

'While inside the hook event, this will be set to TRUE.  Because we raise events directly from the hook, we sometimes need to postpone
' crucial actions (like releasing the hook) until the hook proc has exited.
Private m_InHookNow As Boolean

'If the user attempts to add items to the combo box before it is created (e.g. when the control is invisible), we will back up the
' items to this string.  When the combo box is created, this list will be automatically be added to the control.
Private Type backupComboEntry
    entryIndex As Long
    entryString As String
End Type

Private m_BackupEntries() As backupComboEntry
Private m_NumBackupEntries As Long
Private m_BackupListIndex As Long

'Additional helpers for rendering themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'Combo box interaction functions
Private Const CB_ADDSTRING As Long = &H143
Private Const CB_DELETESTRING As Long = &H144
Private Const CB_INSERTSTRING As Long = &H14A
Private Const CB_RESETCONTENT As Long = &H14B
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_SETITEMDATA As Long = &H151

Private Const CBN_SELCHANGE As Long = 1
Private Const CBN_DROPDOWN As Long = 7
Private Const CBN_CLOSEUP As Long = 8

Private Const CBS_SIMPLE As Long = &H1
Private Const CBS_DROPDOWN As Long = &H2
Private Const CBS_DROPDOWNLIST As Long = &H3

Private Const CBS_AUTOHSCROLL As Long = &H40&
Private Const CBS_HASSTRINGS As Long = &H200&
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400&
Private Const CBS_OWNERDRAWFIXED As Long = &H10&
Private Const CBS_OWNERDRAWVARIABLE As Long = &H20&

'Owner-drawn combo boxes require us to fill and/or use these structs during painting
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type

'A DRAWITEMSTRUCT instance will specify one or more of these draw actions; as such, make sure to mask the values when checking them
Private Const ODA_DRAWENTIRE As Long = &H1    'Redraw the whole item from scratch
Private Const ODA_SELECT As Long = &H2        'Select state has changed (note: particularly relevant for checkbox-style drop-downs)
Private Const ODA_FOCUS As Long = &H4         'Focus has changed

'A DRAWITEMSTRUCT instance will return one or more of these states; as such, make sure to mask the values when checking them
Private Const ODS_CHECKED As Long = &H8
Private Const ODS_DISABLED As Long = &H4
Private Const ODS_FOCUS As Long = &H10
Private Const ODS_GRAYED As Long = &H2
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_COMBOBOXEDIT As Long = &H1000
Private Const ODS_HOTLIGHT As Long = &H40&

Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    ItemState As Long
    hWndItem As Long
    hDC As Long
    rcItem As RECTL
    ItemData As Long
End Type

'The MeasureItemStruct struct, above, will identify the combo box using this constant in the CtlType field.
Private Const ODT_COMBOBOX As Long = &H3

Private Type COMBOBOXINFO
    cbSize As Long
    rcItem As RECTL
    rcButton As RECTL
    lStateButton As Long
    hWndCombo As Long
    hWndEdit As Long
    hWndList As Long
End Type

'The COMBOBOXINFO struct, above, will report button state using one of the following constants:
Private Const STATE_SYSTEM_UNPRESSED = &H0&
Private Const STATE_SYSTEM_PRESSED = &H8&
Private Const STATE_SYSTEM_INVISIBLE = &H8000&

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hWndCombo As Long, ByRef pcbi As COMBOBOXINFO) As Long

'When creating a window, we assign it a unique ID.  This is handled via GetTickCount, which is as close to a pseudo-random ID as I care to implement.
Private m_ComboBoxWindowID As Long

'Because the control is owner-drawn, we must perform our own measurements.  We calculate these against a test string when creating the combo box;
' the results are stored to this variable, and used for any subsequent measurements
Private m_ItemHeight As Long

'For consistency with other PD controls, the .ListIndex property will always poll the API window for its current value.  However, we need to do
' some separate internal tracking to simplify the rendering process (since this controls is fully owner-drawn).  The last ListIndex change
' notification will set this module-level variable to match the current .ListIndex; this value is used when draw notifications are received, to
' differentiate between hovered items and the actually selected current item (which are not differentiated in the draw struct).
Private m_CurrentListIndex As Long

'If the mouse is currently over the combo box area, this will be set to TRUE
Private m_MouseOverComboBox As Boolean

'Basic combo box interaction functions

'Add an item to the combo box
Public Sub AddItem(ByVal srcItem As String, Optional ByVal itemIndex As Long = -1)
    
    'Because Windows automatically converts to ANSI the Unicode contents of a Unicode combo box inside an ANSI project, we cannot simply add
    ' Unicode strings directly to the combo box.  Instead, we manage our own Unicode-friendly copies of the combo box contents, as a backup.
    ' This way, the combo box manages storage (with support for accessibility, just ANSI-only), and as far as rendering the actual strings goes,
    ' we handle that part ourselves.
    
    'Resize the backup array as necessary
    If m_NumBackupEntries = 0 Then ReDim m_BackupEntries(0 To 15) As backupComboEntry
    If m_NumBackupEntries > UBound(m_BackupEntries) Then ReDim Preserve m_BackupEntries(0 To m_NumBackupEntries * 2 - 1) As backupComboEntry
    
    'Add this item to the backup array
    m_BackupEntries(m_NumBackupEntries).entryIndex = itemIndex
    m_BackupEntries(m_NumBackupEntries).entryString = srcItem
    m_NumBackupEntries = m_NumBackupEntries + 1
    
    'Add this item to the API combo box.
    copyStringToComboBox m_NumBackupEntries - 1
    
End Sub

'Remove an item from the combo box
Public Sub RemoveItem(ByVal itemIndex As Long)
    
    'First, make sure the requested index is valid
    If (itemIndex >= 0) And (itemIndex < m_NumBackupEntries) Then
    
        'Cache the current .ListIndex; we will need this later on, to prevent the "-1" .ListIndex scenario
        Dim priorListIndex As Long
        priorListIndex = Me.ListIndex
    
        'Modify our internal string storage first
        Dim i As Long
        For i = itemIndex To m_NumBackupEntries - 1
        
            If (i + 1) < UBound(m_BackupEntries) Then
                m_BackupEntries(i) = m_BackupEntries(i + 1)
                
                'Copy the index value of our backup array (for this string) into the new itemdata slot.  This allows us to retrieve the Unicode
                ' version of the string, if any, at render time.
                SendMessage m_ComboBoxHwnd, CB_SETITEMDATA, i, ByVal i
                
            End If
            
        Next i
        
        m_NumBackupEntries = m_NumBackupEntries - 1
        
        'Forward the request to the API box as well
        SendMessage m_ComboBoxHwnd, CB_DELETESTRING, itemIndex, ByVal 0&
        
        'Resize the drop-down portion to match; see copyStringToComboBox, below, for details on how this works
        If g_IsProgramCompiled And g_IsVistaOrLater Then
            SendMessage m_ComboBoxHwnd, CB_SETMINVISIBLE, SendMessage(m_ComboBoxHwnd, CB_GETCOUNT, 0, ByVal 0&), ByVal 0&
        Else
            Dim comboRect As RECTL
            GetClientRect Me.hWnd, comboRect
            MoveWindow m_ComboBoxHwnd, comboRect.Left, comboRect.Top, comboRect.Right - comboRect.Left, (comboRect.Bottom - comboRect.Top) + ((SendMessage(m_ComboBoxHwnd, CB_GETCOUNT, 0, ByVal 0&) + 1) * SendMessage(m_ComboBoxHwnd, CB_GETITEMHEIGHT, 0, ByVal 0&)), 1
        End If
        
        'If the removal affected the current ListIndex, update it to match
        If itemIndex <= priorListIndex Then
            
            'Make sure there is at least one valid entry in the drop-down
            If Me.ListCount >= 0 Then
            
                If priorListIndex > 0 Then
                    Me.ListIndex = priorListIndex - 1
                Else
                    Me.ListIndex = 0
                End If
                
            End If
            
        End If
    
    End If
    
End Sub

'Duplicate a given string inside the API combo box.  We don't actually use this copy of the string (we use our own, so we can support Unicode),
' but this provides a fallback for accessibility technology.
Private Sub copyStringToComboBox(ByVal strIndex As Long)

    'Add this item to the combo box exists
    If (m_ComboBoxHwnd <> 0) Then
    
        'If no index is specified, let the default combo box handler decide order; otherwise, request the placement we were given.
        Dim newIndex As Long
        
        If m_BackupEntries(strIndex).entryIndex = -1 Then
            newIndex = SendMessage(m_ComboBoxHwnd, CB_ADDSTRING, 0, ByVal StrPtr(m_BackupEntries(strIndex).entryString))
        Else
            newIndex = SendMessage(m_ComboBoxHwnd, CB_INSERTSTRING, m_BackupEntries(strIndex).entryIndex, ByVal StrPtr(m_BackupEntries(strIndex).entryString))
        End If
        
        'Copy the index value of our backup array (for this string) into the itemdata slot.  This allows us to retrieve the Unicode
        ' version of the string, if any, at render time.
        SendMessage m_ComboBoxHwnd, CB_SETITEMDATA, newIndex, ByVal strIndex
        
        'Set the combo box to always display the full list amount in the drop-down; this is only applicable if a manifest is present,
        ' so it will have no effect in the IDE.
        If g_IsProgramCompiled And g_IsVistaOrLater Then
            SendMessage m_ComboBoxHwnd, CB_SETMINVISIBLE, SendMessage(m_ComboBoxHwnd, CB_GETCOUNT, 0, ByVal 0&), ByVal 0&
        
        'If a manifest is not present, we can achieve the same thing by manually setting the window size to match the number of
        ' entries in the combo box.
        Else
        
            'Rather than forcing combo boxes to a predetermined size, we dynamically adjust their size as items are added.
            ' To do this, we must start by getting the window rect of the current combo box.
            Dim comboRect As RECTL
            GetClientRect Me.hWnd, comboRect
            
            'Next, resize the combo box to match the number of items currently in the box.
            MoveWindow m_ComboBoxHwnd, comboRect.Left, comboRect.Top, comboRect.Right - comboRect.Left, (comboRect.Bottom - comboRect.Top) + ((SendMessage(m_ComboBoxHwnd, CB_GETCOUNT, 0, ByVal 0&) + 1) * SendMessage(m_ComboBoxHwnd, CB_GETITEMHEIGHT, 0, ByVal 0&)), 1
            
        End If
        
    End If

End Sub

'Clear all entries from the combo box
Public Sub Clear()
    If m_ComboBoxHwnd <> 0 Then
        SendMessage m_ComboBoxHwnd, CB_RESETCONTENT, 0, ByVal 0&
    Else
        m_NumBackupEntries = 0
        ReDim m_BackupEntries(0) As backupComboEntry
    End If
End Sub

'Number of items currently in the combo box list
Public Function ListCount() As Long
    
    'We do not track ListCount persistently.  It is requested on-demand from the combo box.
    If m_ComboBoxHwnd <> 0 Then
        ListCount = SendMessage(m_ComboBoxHwnd, CB_GETCOUNT, 0, ByVal 0&)
    Else
        ListCount = m_NumBackupEntries
    End If
    
End Function

'Retrieve a specified list item
Public Property Get List(ByVal indexOfItem As Long) As String
    
    'We do not track ListCount persistently.  It is requested on-demand from the combo box.
    If (indexOfItem >= 0) And (indexOfItem < m_NumBackupEntries) Then
        List = m_BackupEntries(indexOfItem).entryString
    Else
        List = ""
    End If
    
End Property

'Get/set the currently active item.
' NB: unlike the default VB combo box, we do not raise an error if an invalid index is requested.
Public Property Get ListIndex() As Long
    
    'We do not track ListIndex persistently.  It is requested on-demand from the combo box.
    If m_ComboBoxHwnd <> 0 Then
        ListIndex = SendMessage(m_ComboBoxHwnd, CB_GETCURSEL, 0, ByVal 0&)
    End If
    
End Property

Public Property Let ListIndex(ByVal newIndex As Long)

    'Make a backup of the new listindex
    m_CurrentListIndex = newIndex

    If m_ComboBoxHwnd <> 0 Then
        
        'See if new ListIndex is different from the current ListIndex.  (We can skip the assignment step if they match.)
        If newIndex <> SendMessage(m_ComboBoxHwnd, CB_GETCURSEL, 0, ByVal 0&) Then
            
            'Request the new list index
            SendMessage m_ComboBoxHwnd, CB_SETCURSEL, newIndex, ByVal 0&
            
            'Request an immediate repaint; without this, there may be a delay, based on the caller's handling of the Click event
            cPainterBox.requestRepaint
            
            'Notify the user of the change
            RaiseEvent Click
            
        End If
        
    'If the combo box doesn't exist yet, maek a backup of any ListIndex requests
    Else
        m_BackupListIndex = newIndex
    End If
    
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
' TODO: disable API box as well
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    
    'If the control is disabled, the BackColor property actually becomes relevant (because the edit box will allow the back color
    ' to "show through").  As such, set it now, and note that we can use VB's internal property, because it simply wraps the
    ' matching GDI function(s).
    If g_IsProgramRunning Then
        If NewValue Then
            UserControl.BackColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        Else
            UserControl.BackColor = g_Themer.getThemeColor(PDTC_GRAY_HIGHLIGHT)
        End If
    End If
    
    If m_ComboBoxHwnd <> 0 Then EnableWindow m_ComboBoxHwnd, IIf(NewValue, 1, 0)
    UserControl.Enabled = NewValue
    
    PropertyChanged "Enabled"
    
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        
        If Not (curFont Is Nothing) Then
            
            'Recreate the font object
            curFont.releaseFromDC
            curFont.setFontSize m_FontSize
            curFont.createFontObject
            
            'Combo box sizes are set by the system, at creation time, so we don't have a choice but to recreate the box now
            createComboBox
            
        End If
        
        PropertyChanged "FontSize"
        
    End If
    
End Property

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseOverComboBox = True
    cPainterBox.requestRepaint
    
    'Set a hand cursor
    cMouseEvents.setSystemCursor IDC_HAND
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseOverComboBox = False
    cPainterBox.requestRepaint
    
    'Reset the cursor
    cMouseEvents.setSystemCursor IDC_ARROW
    
End Sub

Private Sub cMouseEventsListBox_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEventsListBox.setSystemCursor IDC_HAND
End Sub

Private Sub cMouseEventsListBox_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEventsListBox.setSystemCursor IDC_ARROW
End Sub

Private Sub cMouseEventsListBox_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEventsListBox.setSystemCursor IDC_HAND
End Sub

'Flicker-free paint requests for the main control box (e.g. NOT the drop-down list portion)
Private Sub cPainterBox_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    drawComboBox True
End Sub

Private Sub tmrHookRelease_Timer()

    'If a hook is active, this timer will repeatedly try to kill it.  Do not enable it until you are certain the hook needs to be released.
    ' (This is used as a failsafe if we cannot immediately release the hook when focus is lost, for example if we are currently inside an
    '  external event, as happens in the Layer toolbox, which hides the active text box inside the KeyPress event.)
    If (m_ComboBoxHwnd <> 0) And (Not m_InHookNow) Then
        RemoveHookConditional
        tmrHookRelease.Enabled = False
    End If
    
End Sub

'When the control receives focus, forcibly forward focus to the edit window
Private Sub UserControl_GotFocus()
    
    'The user control itself should never have focus.  Forward it to the API edit box.
    If m_ComboBoxHwnd <> 0 Then
        SetForegroundWindow m_ComboBoxHwnd
        SetFocus m_ComboBoxHwnd
    End If
    
End Sub

'When the user control is hidden, we must hide the edit box window as well
Private Sub UserControl_Hide()
    If m_ComboBoxHwnd <> 0 Then ShowWindow m_ComboBoxHwnd, SW_HIDE
End Sub

Private Sub UserControl_Initialize()

    m_ComboBoxHwnd = 0
    
    Set curFont = New pdFont
        
    'Note that we are not currently responsible for any resize events
    m_InternalResizeState = False
    
    'At run-time, initialize a subclasser
    If g_IsProgramRunning Then
        Set cSubclass = New cSelfSubHookCallback
        Set cPainterBox = New pdWindowPainter
    End If
    
    'When not in design mode, initialize a tracker for mouse events
    If g_IsProgramRunning Then
                        
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        
        Set g_Themer = New pdVisualThemes
        
    End If
    
    'Create an initial font object
    refreshFont
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    FontSize = 10
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
    End With

End Sub

'Show the control and the combo box.  (This is the first place the combo box is typically created, as well.)
Private Sub UserControl_Show()
    
    'If we have not yet created the combo box, do so now.
    If m_ComboBoxHwnd = 0 Then
        
        createComboBox
        
    'The combo box has already been created, so we just need to show it.  Note that we explicitly set flags to NOT activate
    ' the window, as we don't want it stealing focus.
    Else
        If m_ComboBoxHwnd <> 0 Then ShowWindow m_ComboBoxHwnd, SW_SHOWNA
    End If
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).  Note that this has the ugly side-effect of
    ' permanently erasing the extender's tooltip, so FOR THIS CONTROL, TOOLTIPS MUST BE SET AT RUN-TIME!
    '
    'TODO!  Add helper functions for setting the tooltip to the created hWnd, instead of the VB control
    m_ToolString = Extender.ToolTipText

    If m_ToolString <> "" Then

        Set m_ToolTip = New clsToolTip
        With m_ToolTip

            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool Me, m_ToolString
            Extender.ToolTipText = ""

        End With

    End If

End Sub

'TODO: solve drawing for the combo box.  We probably don't need a border, like we used for the edit box...
Private Sub getComboBoxRect(ByRef targetRect As winRect)

    With targetRect
        .x1 = 0
        .y1 = 0
        .x2 = UserControl.ScaleWidth
        .y2 = UserControl.ScaleHeight
    End With

End Sub

'Create a brush for drawing the box background
Private Sub createComboBoxBrush()

    If m_ComboBoxBrush <> 0 Then DeleteObject m_ComboBoxBrush
    
    If g_IsProgramRunning Then
        m_ComboBoxBrush = CreateSolidBrush(g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT))
    Else
        m_ComboBoxBrush = CreateSolidBrush(RGB(128, 255, 255))
    End If

End Sub

'After curFont has been created, this function can be used to return the "ideal" height of a string rendered via the current font.
Private Function getIdealStringHeight() As Long

    Dim attachedDC As Long
    attachedDC = curFont.getAttachedDC
    curFont.releaseFromDC
    
    'Create a temporary DC
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank 1, 1, 24
    
    'Select the current font into that DC
    curFont.attachToDC tmpDIB.getDIBDC
    
    'Determine a standard string height
    getIdealStringHeight = curFont.getHeightOfString("abc123")
    
    'Remove the font and release our temporary DIB
    curFont.releaseFromDC
    curFont.attachToDC attachedDC
    
    'tmpDIB will be automatically released
    
End Function

'As the wrapped system combo box may need to be recreated when certain properties are changed, this function is used to
' automate the process of destroying an existing window (if any) and recreating it anew.
Private Function createComboBox() As Boolean
    
    'If the combo box already exists, kill it
    ' TODO: do we ever actually need to do this, or can we get away with a single creation??
    destroyComboBox
    
    'Create a brush for drawing the box background
    createComboBoxBrush
    
    'Figure out which flags to use, based on the control's properties
    Dim flagsWinStyle As Long, flagsWinStyleExtended As Long, flagsComboControl As Long
    flagsWinStyle = WS_VISIBLE Or WS_CHILD Or WS_VSCROLL Or WS_HSCROLL
    flagsWinStyleExtended = 0
    
    'PhotoDemon only supports simple drop-downs.  Similarly, all drop-down entries are coerced into strings, so we can use the CBS_HASSTRINGS
    ' setting, which instructs the drop-down to manage its own string memory (instead of us doing it manually).  This is a much better solution
    ' for accessibility interoperability; see http://msdn.microsoft.com/en-us/library/windows/desktop/dd318073%28v=vs.85%29.aspx
    flagsComboControl = CBS_DROPDOWNLIST Or CBS_HASSTRINGS Or CBS_OWNERDRAWFIXED
    
    'The underlying user control should ignore any height values set from the IDE; instead, it should be forced to an ideal height,
    ' using the current font as our guide.  We check for this here, prior to creating the combo box, as we can't easily
    ' access our font object once we assign it to the combo box.
    If Not (curFont Is Nothing) Then
        
        Dim idealHeight As Long
        idealHeight = getIdealStringHeight()
        
        'Cache this value at module-level; we will need it for subsequent WM_MEASUREITEM requests sent to the parent
        m_ItemHeight = idealHeight
        
        'Resize the user control accordingly; the formula for height is the string height + 5px of borders.
        ' (5px = 2px on top, 3px on bottom.)  User control width is not changed.
        m_InternalResizeState = True
        
        'If the program is running (e.g. NOT design-time) resize the user control to match.  This improves compile-time performance, as there
        ' are a lot of instances in this control, and their size events will be fired during compilation.
        If g_IsProgramRunning Then
            UserControl.Height = ScaleY(idealHeight + 8, vbPixels, vbTwips)
        End If
                
        m_InternalResizeState = False
                    
    End If
    
    'Determine a unique ID for this combo box instance.  This is needed to identify this control against other owner-drawn controls on the same parent.
    m_ComboBoxWindowID = GetTickCount()
    
    'Prior to creating the combo box, we need to subclass the parent window.  It is important to do this now, because the combo box is owner-drawn,
    ' so when it is initiated, the parent needs to supply measurement data - so we can't wait until post-creation to subclass the parent.
    If g_IsProgramRunning Then
        If Not (cSubclass Is Nothing) Then
            
            'Subclass the parent user control.
            If Not m_ParentHasBeenSubclassed Then
                cSubclass.ssc_Subclass UserControl.hWnd, 0, 1, Me, True, True, False
                cSubclass.ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_CTLCOLOREDIT, WM_COMMAND, WM_MEASUREITEM, WM_DRAWITEM
                m_ParentHasBeenSubclassed = True
            End If
            
        End If
    End If
    
    'Generate the combo box's window rect, which is positioned relative to the underlying DC
    Dim tmpRect As winRect
    tmpRect.x1 = 0
    tmpRect.y1 = 0
    tmpRect.x2 = UserControl.ScaleWidth
    tmpRect.y2 = idealHeight + 6
    
    'Creating a combo box window is a little different from other windows, because the drop-down height must be factored into the initial
    ' size calculation.  We start at zero, then increase the combo box size as additional items are added.
    With tmpRect
        m_ComboBoxHwnd = CreateWindowEx(flagsWinStyleExtended, ByVal StrPtr("COMBOBOX"), ByVal StrPtr(""), flagsWinStyle Or flagsComboControl, _
        .x1, .y1, .x2, .y2, UserControl.hWnd, m_ComboBoxWindowID, App.hInstance, ByVal 0&)
    End With
    
    'Enable the window per the current UserControl's extender setting
    EnableWindow m_ComboBoxHwnd, IIf(Me.Enabled, 1, 0)
    
    'Assign a subclasser to enable proper tab and arrow key support
    If g_IsProgramRunning Then
    
        If Not (cSubclass Is Nothing) Then
            
            'Subclass the combo box
            cSubclass.ssc_Subclass m_ComboBoxHwnd, 0, 1, Me, True, True, True
            cSubclass.ssc_AddMsg m_ComboBoxHwnd, MSG_BEFORE, WM_KEYDOWN, WM_SETFOCUS, WM_KILLFOCUS, WM_MOUSEACTIVATE
            cSubclass.ssc_AddMsg m_ComboBoxHwnd, MSG_AFTER, WM_SIZE
            
            'Subclass the user control as well.  This is necessary for handling update messages from the edit box
            If Not m_ParentHasBeenSubclassed Then
                cSubclass.ssc_Subclass UserControl.hWnd, 0, 1, Me, True, True, False
                cSubclass.ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_CTLCOLOREDIT, WM_COMMAND
                m_ParentHasBeenSubclassed = True
            End If
            
        End If
        
        'Assign a second subclasser for the window painter
        If Not (cPainterBox Is Nothing) Then
            cPainterBox.startPainter m_ComboBoxHwnd
        End If
        
        '...and a third subclasser for mouse events
        If (cMouseEvents Is Nothing) Then
            Set cMouseEvents = New pdInputMouse
            cMouseEvents.addInputTracker m_ComboBoxHwnd, True, , , True, True
            cMouseEvents.setSystemCursor IDC_HAND
        End If
        
    End If
    
    'Assign the default font to the combo box
    refreshFont True
    
    'If we backed up previous combo box entries at some point, we must restore those entries now.
    If m_NumBackupEntries > 0 Then
        
        Dim i As Long
        For i = 0 To m_NumBackupEntries - 1
            copyStringToComboBox i
        Next i
        
        'Also set a list index, if any.  (If none were specifed, the first entry in the list wil be selected.)
        Me.ListIndex = m_BackupListIndex
        
    End If
    
    'Finally, synchronize the underlying user control to whatever size the system has created the combo box at
    syncUserControlSizeToComboSize
    
    'Return TRUE if successful
    createComboBox = (m_ComboBoxHwnd <> 0)

End Function

'If an edit box currently exists, this function will destroy it.
Private Function destroyComboBox() As Boolean

    If m_ComboBoxHwnd <> 0 Then
        
        If Not cSubclass Is Nothing Then
            cSubclass.ssc_UnSubclass m_ComboBoxHwnd
            cSubclass.shk_TerminateHooks
        End If
        
        DestroyWindow m_ComboBoxHwnd
        
    End If
    
    destroyComboBox = True

End Function

Private Sub UserControl_Terminate()
    
    'Release the edit box background brush
    If m_ComboBoxBrush <> 0 Then DeleteObject m_ComboBoxBrush
    
    'Destroy the edit box, as necessary
    destroyComboBox
    
    'Release any extra subclasser(s)
    If Not cSubclass Is Nothing Then cSubclass.ssc_Terminate
    
End Sub

'When the font used for the edit box changes in some way, it can be recreated (refreshed) using this function.  Note that font
' creation is expensive, so it's worthwhile to avoid this step as much as possible.
Private Sub refreshFont(Optional ByVal forceRefresh As Boolean = False)
    
    Dim fontRefreshRequired As Boolean
    fontRefreshRequired = curFont.hasFontBeenCreated
    
    'Update each font parameter in turn.  If one (or more) requires a new font object, the font will be recreated as the final step.
    
    'Font face is always set automatically, to match the current program-wide font
    If (Len(g_InterfaceFont) <> 0) And (StrComp(curFont.getFontFace, g_InterfaceFont, vbTextCompare) <> 0) Then
        fontRefreshRequired = True
        curFont.setFontFace g_InterfaceFont
    End If
    
    'In the future, I may switch to GDI+ for font rendering, as it supports floating-point font sizes.  In the meantime, we check
    ' parity using an Int() conversion, as GDI only supports integer font sizes.
    If Int(m_FontSize) <> Int(curFont.getFontSize) Then
        fontRefreshRequired = True
        curFont.setFontSize m_FontSize
    End If
        
    'Request a new font, if one or more settings have changed
    If fontRefreshRequired Or forceRefresh Then
        
        curFont.createFontObject
        
        'Whenever the font is recreated, we need to reassign it to the combo box.  This is done via the WM_SETFONT message.
        If m_ComboBoxHwnd <> 0 Then SendMessage m_ComboBoxHwnd, WM_SETFONT, curFont.getFontHandle, IIf(UserControl.Extender.Visible, 1, 0)
                    
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", m_FontSize, 10
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub updateAgainstCurrentTheme()
    
    If g_IsProgramRunning Then
        
        'Create a brush for drawing the box background
        createComboBoxBrush
        
        'Update the current font, as necessary
        refreshFont
        
        'Force an immediate repaint
        cPainterBox.requestRepaint
                
    End If
    
End Sub

'If an object of type Control is capable of receiving focus, this will return TRUE.
Private Function isControlFocusable(ByRef Ctl As Control) As Boolean

    If Not (TypeOf Ctl Is Timer) And Not (TypeOf Ctl Is Line) And Not (TypeOf Ctl Is pdLabel) And Not (TypeOf Ctl Is Frame) And Not (TypeOf Ctl Is Shape) And Not (TypeOf Ctl Is Image) And Not (TypeOf Ctl Is vbalHookControl) And Not (TypeOf Ctl Is ShellPipe) And Not (TypeOf Ctl Is bluDownload) Then
        isControlFocusable = True
    Else
        isControlFocusable = False
    End If

End Function

'Iterate through all sibling controls in our container, and if one is capable of receiving focus, activate it.  I had *really* hoped
' to bypass this kind of manual handling by using WM_NEXTDLGCTL, but I failed to get it working reliably with all types of VB forms.
' I'm honestly not sure whether VB even uses that message, or whether it uses some internal mechanism for focus tracking; the latter
' might explain why a manual approach like this seems to be necessary for us as well.
Private Sub forwardFocusManually(ByVal focusDirectionForward As Boolean)

    'If the user has deactivated tab support, or we are invisible/disabled, ignore this completely
    If UserControl.Extender.TabStop And UserControl.Extender.Visible And UserControl.Enabled Then
        
        'Iterate through all controls in the container, looking for the next TabStop index
        Dim myIndex As Long
        myIndex = UserControl.Extender.TabIndex
        
        Dim newIndex As Long
        Const MAX_INDEX As Long = 99999
        
        'Forward and back focus checks require different search strategies
        If focusDirectionForward Then
            newIndex = MAX_INDEX
        Else
            newIndex = myIndex
        End If
        
        'Some controls may not have a TabStop property.  That's okay - just keep iterating if it happens.
        On Error GoTo NextControlCheck
        
        Dim Ctl As Control, targetControl As Control
        For Each Ctl In Parent.Controls
            
            'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
            ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
            ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
            ' is focusable.
            If isControlFocusable(Ctl) Then
            
                'Ignore controls whose TabStop property is False, who are not visible, or who are disabled
                If Ctl.TabStop And Ctl.Visible And Ctl.Enabled Then
                        
                    If focusDirectionForward Then
                    
                        'Check the tab index of this control.  We're looking for the lowest tab index that is also larger than our tab index.
                        If (Ctl.TabIndex > myIndex) And (Ctl.TabIndex < newIndex) Then
                            newIndex = Ctl.TabIndex
                            Set targetControl = Ctl
                        End If
                        
                    Else
                    
                        'Check the tab index of this control.  We're looking for the highest tab index that is also larger than our tab index.
                        If (Ctl.TabIndex > newIndex) Then
                            newIndex = Ctl.TabIndex
                            Set targetControl = Ctl
                        End If
                    
                    End If
    
                End If
                
            End If
            
NextControlCheck:
        Next
        
        'When moving focus forward, we now have one of two possibilites:
        ' 1) NewIndex represents the tab index of a valid control whose index is higher than us.
        ' 2) NewIndex is still MAX_INDEX, because no control with a valid tab index was found.
        
        'When moving focus backward, we also have two possibilities:
        ' 1) NewIndex represents the tab index of a valid control whose index is higher than us.  (Required if Shift+Tab will push the
        '     TabIndex below 0.)
        ' 2) NewIndex is still MY_INDEX, because no control with a valid tab index was found.
        
        'Handle case 2 now.
        If (focusDirectionForward And (newIndex = MAX_INDEX)) Or (Not focusDirectionForward) Then
            
            If focusDirectionForward Then
                newIndex = myIndex
            Else
                newIndex = -1
            End If
            
            'Some controls may not have a TabStop property.  That's okay - just keep iterating if it happens.
            On Error GoTo NextControlCheck2
            
            'If our control is last in line for tabstops, we need to now find the LOWEST tab index to forward focus to.
            For Each Ctl In Parent.Controls
                
                'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
                ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
                ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
                ' is focusable.
                If isControlFocusable(Ctl) Then
                    
                    'Ignore controls whose TabStop property is False, who are not visible, or who are disabled
                    If Ctl.TabStop And Ctl.Visible And Ctl.Enabled Then
                            
                        If focusDirectionForward Then
                        
                            'Check the tab index of this control.  We're looking for the lowest valid tab index.
                            If (Ctl.TabIndex < myIndex) And (Ctl.TabIndex < newIndex) Then
                                newIndex = Ctl.TabIndex
                                Set targetControl = Ctl
                            End If
                            
                        Else
                        
                            'Check the tab index of this control.  We're looking for the lowest valid tab index, if one exists.
                            If (Ctl.TabIndex < myIndex) And (Ctl.TabIndex > newIndex) Then
                                newIndex = Ctl.TabIndex
                                Set targetControl = Ctl
                            End If
                        
                        End If
                    
                    End If
                    
                End If
                
NextControlCheck2:
            Next
        
        End If
        
        If (Not focusDirectionForward) Then
            If newIndex = -1 Then newIndex = myIndex
        End If
        
        'Regardless of focus direction, we once again have one of two possibilites.
        ' 1) NewIndex represents the tab index of the next valid control in VB's tab order.
        ' 2) NewIndex = our index, because no control with a valid tab index was found.
        
        'SetFocus can fail under a variety of circumstances, so error handling is still required
        On Error GoTo NoFocusRecipient
        
        'Ignore the second case completely, as tab should have no effect
        If newIndex <> myIndex Then
            targetControl.SetFocus
        
NoFocusRecipient:
        
        End If
        
    End If

End Sub

'Given a virtual keycode, return TRUE if the keycode is a command key that must be manually forwarded to a combo box.  Command keys include
' arrow keys at present, but in the future, additional keys can be added to this list.
'
'NOTE FOR OUTSIDE USERS!  These key constants are declared publicly in PD, because they are used many places.  You can find virtual keycode
' declarations in PD's Public_Constants module, or at this MSDN link:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
Private Function doesVirtualKeyRequireSpecialHandling(ByVal vKey As Long) As Boolean
    
    Select Case vKey
    
        Case VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN
            doesVirtualKeyRequireSpecialHandling = True
            
        Case VK_TAB
            doesVirtualKeyRequireSpecialHandling = True
        
        Case Else
            doesVirtualKeyRequireSpecialHandling = False
    
    End Select
        
End Function

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function isVirtualKeyDown(ByVal vKey As Long) As Boolean
    isVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'Render the combo box area (not the list!)
Private Sub drawComboBox(Optional ByVal srcIsWMPAINT As Boolean = True)

    'Before painting, retrieve detailed information on the combo box
    Dim cbiCombo As COMBOBOXINFO
    cbiCombo.cbSize = LenB(cbiCombo)
    
    'Make sure the combo box exists
    If m_ComboBoxHwnd <> 0 Then
    
        If GetComboBoxInfo(m_ComboBoxHwnd, cbiCombo) <> 0 Then
        
            'cbiCombo now contains a bunch of useful information, including hWnds for each portion of the combo box control, and a rect for both
            ' the button area, and the edit area.  We will render each of these in turn.
            
            'Start by retrieving a DC for the edit area.
            Dim targetDC As Long
            
            If srcIsWMPAINT Then
                targetDC = cPainterBox.getPaintStructDC()
            Else
                targetDC = GetDC(m_ComboBoxHwnd)
            End If
            
            If targetDC <> 0 Then
            
                'Next, determine paint colors, contingent on enablement and other settings.
                Dim cboBorderColor As Long, cboFillColor As Long, cboTextColor As Long, cboButtonColor As Long
                
                If Me.Enabled Then
                    
                    'When the mouse is over the combo box, the border and drop-down arrow are highlighted
                    If m_MouseOverComboBox Then
                        cboBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
                        cboButtonColor = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
                    Else
                        cboBorderColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
                        cboButtonColor = g_Themer.getThemeColor(PDTC_TEXT_DEFAULT)
                    End If
                    
                    If m_HasFocus Then
                        cboFillColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
                        cboTextColor = g_Themer.getThemeColor(PDTC_TEXT_INVERT)
                        cboButtonColor = cboTextColor
                    Else
                        cboFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
                        cboTextColor = g_Themer.getThemeColor(PDTC_TEXT_EDITBOX)
                    End If
                    
                Else
                
                    cboBorderColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
                    cboFillColor = g_Themer.getThemeColor(PDTC_GRAY_HIGHLIGHT)
                    cboTextColor = g_Themer.getThemeColor(PDTC_TEXT_DEFAULT)
                    cboButtonColor = cboTextColor
                
                End If
                
                'Retrieve the full client area
                Dim fullWinRect As RECTL
                GetClientRect m_ComboBoxHwnd, fullWinRect
                
                'Paint the full control background
                Dim tmpBrush As Long
                tmpBrush = CreateSolidBrush(cboFillColor)
                FillRect targetDC, fullWinRect, tmpBrush
                DeleteObject tmpBrush
                
                'Paint the border
                tmpBrush = CreateSolidBrush(cboBorderColor)
                FrameRect targetDC, fullWinRect, tmpBrush
                DeleteObject tmpBrush
                
                'Painting the combo box will also paint over the currently selected item, unfortunately.  To work around this, we must
                ' paint that item manually, after the background has already been rendered.
                
                'Retrieve the string for the active combo box entry.
                Dim stringIndex As Long, tmpString As String
                stringIndex = SendMessage(m_ComboBoxHwnd, CB_GETITEMDATA, m_CurrentListIndex, ByVal 0&)
                If stringIndex >= 0 Then
                    tmpString = m_BackupEntries(stringIndex).entryString
                End If
                
                'Prepare a font renderer, then render the text
                If Not curFont Is Nothing Then
                    
                    curFont.releaseFromDC
                    curFont.setFontColor cboTextColor
                    curFont.attachToDC targetDC
                    
                    With cbiCombo.rcItem
                        curFont.fastRenderTextWithClipping .Left + 4, .Top, (.Right - .Left) - 4, (.Bottom - .Top) - 2, tmpString, False
                    End With
                    
                    curFont.releaseFromDC
                    
                End If
                
                'Draw the button
                Dim buttonPt1 As POINTFLOAT, buttonPt2 As POINTFLOAT, buttonPt3 As POINTFLOAT
                
                'Start with the downward-pointing arrow
                buttonPt1.x = fullWinRect.Right - fixDPIFloat(16)
                buttonPt1.y = (fullWinRect.Bottom - fullWinRect.Top) / 2 - fixDPIFloat(1)
                
                buttonPt3.x = fullWinRect.Right - fixDPIFloat(8)
                buttonPt3.y = buttonPt1.y
                
                buttonPt2.x = buttonPt1.x + (buttonPt3.x - buttonPt1.x) / 2
                buttonPt2.y = buttonPt1.y + fixDPIFloat(3)
                
                GDI_Plus.GDIPlusDrawLineToDC targetDC, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y, cboButtonColor, 255, 2, True, LineCapRound
                GDI_Plus.GDIPlusDrawLineToDC targetDC, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y, cboButtonColor, 255, 2, True, LineCapRound
                
                'For an OSX-type look, we can mirror the arrow across the control's center line, then draw it again; I personally prefer
                ' this behavior (as the list box may extend up or down), but I'm not sold on implementing it just yet, because it's out of place
                ' next to regular Windows drop-downs...
                'buttonPt1.y = fullWinRect.Bottom - buttonPt1.y
                'buttonPt2.y = fullWinRect.Bottom - buttonPt2.y
                'buttonPt3.y = fullWinRect.Bottom - buttonPt3.y
                '
                'GDI_Plus.GDIPlusDrawLineToDC targetDC, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y, cboButtonColor, 255, 2, True, LineCapRound
                'GDI_Plus.GDIPlusDrawLineToDC targetDC, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y, cboButtonColor, 255, 2, True, LineCapRound
                                
                'Release the DC
                If Not srcIsWMPAINT Then
                    ReleaseDC m_ComboBoxHwnd, targetDC
                End If
                    
            End If
            
        End If
    
    End If

End Sub

'Given a DRAWITEMSTRUCT object, draw the corresponding item.  This function returns TRUE if drawing was successful.
Private Function drawComboBoxEntry(ByRef srcDIS As DRAWITEMSTRUCT) As Boolean

    Dim drawSuccess As Boolean
    drawSuccess = False
    
    'The control type should always be combo box, but it doesn't hurt to check
    If srcDIS.CtlType = ODT_COMBOBOX Then
        
        'If the ItemID is -1, the combo box is empty; this case is important to check, because an empty combo box won't have any text data,
        ' so attempting to retrieve text entries will fail.
        If (srcDIS.itemID <> -1) Then
            
            'Determine colors.  Obviously these vary depending on the selection state of the current entry
            Dim itemBackColor As Long, itemTextColor As Long
            Dim isMouseOverItem As Boolean
            isMouseOverItem = ((srcDIS.ItemState And ODS_SELECTED) <> 0)
            
            'If the current entry is also the ListIndex, and the control has focus, render it inversely
            If (srcDIS.itemID = m_CurrentListIndex) And m_HasFocus Then
                itemTextColor = g_Themer.getThemeColor(PDTC_TEXT_INVERT)
                itemBackColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            
            'If this entry is not the ListIndex, or the control does not have focus, render the item normally.
            Else
                
                'If the mouse is currently over this item, highlight the text.  This is in keeping with other hover behavior in PD.
                If isMouseOverItem And m_HasFocus Then
                    itemTextColor = g_Themer.getThemeColor(PDTC_TEXT_HYPERLINK)
                Else
                    itemTextColor = g_Themer.getThemeColor(PDTC_TEXT_EDITBOX)
                End If
                
                itemBackColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
                
            End If
            
            'Fill the background
            Dim tmpBackBrush As Long
            tmpBackBrush = CreateSolidBrush(itemBackColor)
            FillRect srcDIS.hDC, srcDIS.rcItem, tmpBackBrush
            DeleteObject tmpBackBrush
            
            'Retrieve the string for the active combo box entry.
            Dim stringIndex As Long, tmpString As String
            stringIndex = SendMessage(m_ComboBoxHwnd, CB_GETITEMDATA, srcDIS.itemID, ByVal 0&)
            tmpString = m_BackupEntries(stringIndex).entryString
            
            'Prepare a font renderer, then render the text
            If Not curFont Is Nothing Then
                
                curFont.releaseFromDC
                curFont.setFontColor itemTextColor
                curFont.attachToDC srcDIS.hDC
                
                With srcDIS.rcItem
                    curFont.fastRenderTextWithClipping .Left + 4, .Top, (.Right - .Left) - 4, (.Bottom - .Top) - 2, tmpString, False
                End With
                
                curFont.releaseFromDC
                
            End If
                        
            'If the item has focus, draw a rectangular frame around the item.
            If isMouseOverItem Then
                tmpBackBrush = CreateSolidBrush(g_Themer.getThemeColor(PDTC_ACCENT_SHADOW))
                FrameRect srcDIS.hDC, srcDIS.rcItem, tmpBackBrush
                DeleteObject tmpBackBrush
            End If
            
            'Note that we have handled the draw request successfully
            drawSuccess = True
            
        'The combo box is empty (TODO: do we still need to render a background here??)
        Else
        
            drawSuccess = False
        
        End If
    
    'The item type is not combo box - not sure what to do here (TODO)
    Else
    
        Debug.Print "draw request received for a non-combo-box item!"
        drawSuccess = False
    
    End If
    
    drawComboBoxEntry = drawSuccess

End Function

'When creating the combo box (or when the combo box is resized due to some external event), use this function to sync the underlying usercontrol size
Private Sub syncUserControlSizeToComboSize()

    'Get the window rect of the combo box
    Dim comboRect As RECTL
    GetWindowRect m_ComboBoxHwnd, comboRect
    
    'Resize the user control, as necessary
    With UserControl
    
        If g_IsProgramCompiled Then
            MoveWindow UserControl.hWnd, UserControl.Extender.Left, UserControl.Extender.Top, comboRect.Right - comboRect.Left, comboRect.Bottom - comboRect.Top, 1
        Else
   
            If (comboRect.Bottom - comboRect.Top) <> .ScaleHeight Or (comboRect.Right - comboRect.Left) <> .ScaleWidth Then
                .Size .ScaleX((comboRect.Right - comboRect.Left), vbPixels, vbTwips), .ScaleY((comboRect.Bottom - comboRect.Top), vbPixels, vbTwips)
            End If
   
        End If
    
    End With
    
    'Repaint the control
    If Not (cPainterBox Is Nothing) Then cPainterBox.requestRepaint

End Sub

'Install a keyboard hook in our window
Private Sub InstallHookConditional()

    'Check for an existing hook
    If Not m_HasFocus Then
    
        'Note the time.  This is used to prevent keypresses occurring immediately prior to the hook, from being
        ' caught within our hook proc!
        m_TimeAtFocusEnter = GetTickCount
        
        'Note that this window is now active
        m_HasFocus = True
        cPainterBox.requestRepaint
        
        'No hook exists.  Hook the control now.
        cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_ComboBoxHwnd, 2, Me, , True
            
    End If

End Sub

Private Sub RemoveHookConditional()

    'Check for an existing hook
    If m_HasFocus Then
        
        'Note that this window is now considered inactive
        m_HasFocus = False
        cPainterBox.requestRepaint
        
        'A hook exists.  Uninstall it now.
        cSubclass.shk_UnHook WH_KEYBOARD
                
    End If
    
End Sub

'This routine MUST BE KEPT as the next-to-last routine for this form. Its ordinal position determines its ability to hook properly.
Private Sub myHookProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    m_InHookNow = True
    bHandled = False
        
    If (lHookType = WH_KEYBOARD) And m_HasFocus Then
        
        'MSDN states that negative codes must be passed to the next hook, without processing
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx)
        If nCode >= 0 Then
            
            'Bit 31 of lParam is 0 if a key is being pressed, and 1 if it is being released.  We use this to raise
            ' separate KeyDown and KeyUp events, as necessary.
            If lParam < 0 Then
                
                'Dialog keys (e.g. arrow keys) get eaten by VB, so we must manually catch them in this hook, and forward them directly
                ' to the API control.
                If doesVirtualKeyRequireSpecialHandling(wParam) Then
                    
                    'Key up events will be raised twice; once in a transitionary stage, and once again in a final stage.
                    ' To prevent double-raising of KeyUp events, we check the transitionary state before proceeding
                    If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                        
                        'Non-tab keys that require special handling are text-dependent keys (e.g. arrow keys).  Simply forward these
                        ' directly to the API box, and it will take care of the rest.
                        
                        'WM_KEYUP requires that certain lParam bits be set.  See http://msdn.microsoft.com/en-us/library/windows/desktop/ms646281%28v=vs.85%29.aspx
                        SendMessage m_ComboBoxHwnd, WM_KEYUP, wParam, ByVal (lParam And &HDFFFFF81 Or &HC0000000)
                        bHandled = True
                        
                    End If
                
                End If
                
            Else
            
                'Dialog keys (e.g. arrow keys) get eaten by VB, so we must manually catch them in this hook, and forward them directly
                ' to the API control.
                If doesVirtualKeyRequireSpecialHandling(wParam) Then
                    
                    'Tab key is used to redirect focus to a new window.
                    If (wParam = VK_TAB) And ((GetTickCount - m_TimeAtFocusEnter) > 250) Then
                                
                        'Set a module-level shift state, and a flag that tells the hook to deactivate after it eats this keypress.
                        If isVirtualKeyDown(VK_SHIFT) Then m_FocusDirection = 2 Else m_FocusDirection = 1
                                                        
                        'Forward focus to the next control
                        forwardFocusManually (m_FocusDirection = 1)
                        m_FocusDirection = 0
                        
                        bHandled = True
                        
                    Else
                    
                        'WM_KEYDOWN requires that certain bits be set.  See http://msdn.microsoft.com/en-us/library/windows/desktop/ms646280%28v=vs.85%29.aspx
                        SendMessage m_ComboBoxHwnd, WM_KEYDOWN, wParam, ByVal (lParam And &H51111111)
                        bHandled = True
                        
                    End If
                    
                End If
                
            End If
            
        End If
            
    End If
    
    'Per MSDN, return the value of CallNextHookEx, contingent on whether or not we handled the keypress internally.
    ' Note that if we do not manually handle a keypress, this behavior allows the default keyhandler to deal with
    ' the pressed keys (and raise its own WM_CHAR events, etc).
    If (Not bHandled) Then
        lReturn = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    Else
        lReturn = 1
    End If
    
    m_InHookNow = False
    
End Sub

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    'FYI: two types of messages can be received here: notification messages sent to the parent window (such as control state changes),
    ' and internal combo box messages.
    Select Case uMsg
        
        'The parent receives this message for all kinds of things; we subclass it to track when the edit box's contents have changed.
        ' (And when we don't handle the message, it is *very important* that we forward it correctly!
        Case WM_COMMAND
        
            'Make sure the command is relative to *our* combo box, and not another one
            If lParam = m_ComboBoxHwnd Then
        
                'Check for the CBN_SELCHANGE flag; if present, raise the CLICK event
                If (wParam \ &H10000) = CBN_SELCHANGE Then
                    m_CurrentListIndex = SendMessage(m_ComboBoxHwnd, CB_GETCURSEL, 0, ByVal 0&)
                    cPainterBox.requestRepaint
                    RaiseEvent Click
                    bHandled = True
                End If
                
                'Check for the DROPDOWN flag; if present, we need to assign a hand cursor to the list box
                If (wParam \ &H10000) = CBN_DROPDOWN Then
                    
                    'Retrieve the hWnd of the dropdown
                    Dim cbiCombo As COMBOBOXINFO
                    cbiCombo.cbSize = LenB(cbiCombo)
                    If GetComboBoxInfo(m_ComboBoxHwnd, cbiCombo) <> 0 Then
                        
                        'Note: this does not actually set a hand cursor when compiled.  No idea why.  TODO!
                        Set cMouseEventsListBox = New pdInputMouse
                        cMouseEventsListBox.addInputTracker cbiCombo.hWndList, True, , , True, True
                        cMouseEventsListBox.setSystemCursor IDC_HAND
                        
                    End If
                    
                End If
                
                'Check for the CLOSEUP flag; if present, we need to kill our mouse handler
                If (wParam \ &H10000) = CBN_CLOSEUP Then
                    cMouseEventsListBox.setSystemCursor IDC_ARROW
                    Set cMouseEventsListBox = Nothing
                End If
                
            End If
        
        'The parent receives this message, prior to the edit box being drawn.  The parent can use this to assign text and
        ' background colors to the edit box.
        Case WM_CTLCOLOREDIT
            
            'Make sure the command is relative to *our* combo box, and not another one
            If lParam = m_ComboBoxHwnd Then
            
                'We can set the text color directly, using the API
                If g_IsProgramRunning Then
                    SetTextColor wParam, g_Themer.getThemeColor(PDTC_TEXT_EDITBOX)
                Else
                    SetTextColor wParam, RGB(0, 0, 128)
                End If
                
                'We return the background brush
                bHandled = True
                lReturn = m_ComboBoxBrush
                
            End If
        
        'Because our combo box is owner-drawn, the system will request a measurement for the drop-down entries.
        Case WM_MEASUREITEM
            
            'Check the control ID (specified by wParam) before proceeding
            If wParam = m_ComboBoxWindowID Then
            
                'Retrieve the MeasureItemStruct pointed to by lParam
                Dim MIS As MEASUREITEMSTRUCT
                CopyMemory MIS, ByVal lParam, LenB(MIS)
                
                'The control type should always be combo box, but it doesn't hurt to check
                If MIS.CtlType = ODT_COMBOBOX Then
                    
                    'If the ItemID is -1, the edit box is the source of the measure item.  Otherwise, it is the dropdown.
                    If MIS.itemID = -1 Then
                        
                        'Fill the height parameter; note that m_ItemHeight is the literal height of a string using the current font.
                        ' Any padding values must be added here.  (I've gone with 1px on either side.)
                        MIS.itemHeight = m_ItemHeight + 2
                        
                        'Copy the pointer to our modified MEASUREITEMSTRUCT back into lParam
                        CopyMemory ByVal lParam, MIS, LenB(MIS)
                        
                        'Note that we have handled the message successfully
                        bHandled = True
                        lReturn = 1
                        
                    Else
                    
                        'Fill the height parameter; note that m_ItemHeight is the literal height of a string using the current font.
                        ' Any padding values must be added here.  (I've gone with 1px on either side.)
                        MIS.itemHeight = m_ItemHeight + 2
                        
                        'Copy the pointer to our modified MEASUREITEMSTRUCT back into lParam
                        CopyMemory ByVal lParam, MIS, LenB(MIS)
                        
                        'Note that we have handled the message successfully
                        bHandled = True
                        lReturn = 1
                        
                    End If
                    
                End If
                
            End If
        
        'Because our combo box is owner-drawn, the system will forward draw requests to us.
        Case WM_DRAWITEM
        
            'Check the control ID (specified by wParam) before proceeding
            If wParam = m_ComboBoxWindowID Then
            
                'Retrieve the DrawItemStruct pointed to by lParam
                Dim DIS As DRAWITEMSTRUCT
                CopyMemory DIS, ByVal lParam, LenB(DIS)
                
                'Forward the DrawItemStruct to the dedicated draw sub
                If drawComboBoxEntry(DIS) Then
                    bHandled = True
                    lReturn = 1
                End If
                
            End If
            
        'On mouse activation, the previous VB window/control that had focus will not be redrawn to reflect its lost focus state.
        ' (Presumably, this is because VB handles focus internally, rather than using standard window messages.)  To avoid the
        ' appearance of two controls simultaneously having focus, we re-set focus to the underlying user control, which forces
        ' VB to redraw the lost focus state of whatever control previously had focus.
        Case WM_MOUSEACTIVATE
            If Not m_HasFocus Then UserControl.SetFocus
            
        'When the control receives focus, initialize a keyboard hook.  This prevents accelerators from working, but it is the
        ' only way to bypass VB's internal message translator, which will forcibly intercept dialog keys (arrows, etc).
        ' Note that focus changes also force a repaint of the control.
        Case WM_SETFOCUS
            InstallHookConditional
            cPainterBox.requestRepaint
            
        Case WM_KILLFOCUS
            If m_InHookNow Then
                tmrHookRelease.Enabled = True
            Else
                RemoveHookConditional
            End If
            cPainterBox.requestRepaint
            
        'Resize messages must be handled manually for the combo box, as we need to dynamically sync the resize state of both parent and child window
        Case WM_SIZE
                        
            'Disable VB resize handling prior to synchronizing size
            m_InternalResizeState = True
            
            'Sync the underlying user control to the combo box's dimensions
            syncUserControlSizeToComboSize
            
            'Restore VB's internal resize handler
            m_InternalResizeState = False
            
        'Other messages??
        Case Else
    
    End Select



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub



