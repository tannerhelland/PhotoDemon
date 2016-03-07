Attribute VB_Name = "Toolboxes"
'***************************************************************************
'Toolbox Manager
'Copyright 2001-2016 by Tanner Helland
'Created: 6/12/01
'Last updated: 06/March/16
'Last update: split toolbox-specific code out of the (now very large) Interface module and into this dedicated home.
'
'Miscellaneous routines related to rendering and handling PhotoDemon's toolboxes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE As Long = 0&
Private Const SW_SHOW As Long = 5&
Private Const SW_SHOWNA As Long = 8&

Private Type PD_Toolbox_Data
    ConstrainingSize As Long
    hWnd As Long
    toolRect As winRect
    
    DefaultSize As Long     'These three settings are *hard-coded*.
    MinSize As Long         'The user has no control over them.
    MaxSize As Long         '(That said, note that they are automatically DPI-corrected at run-time.)
    
    IsVisibleNow As Boolean         'In case we ever need to write these structs to file, I like to keep unaligned struct members at the end
    IsVisiblePreference As Boolean  'Why two visibility settings?  Some toolbars are auto-hidden under certain circumstances, but the user's
                                    ' *preference* for visibility remains unchanged.
End Type

Public Enum PD_Toolbox
    [_First] = 0
    PDT_LeftToolbox = 0
    PDT_BottomToolbox = 1
    PDT_RightToolbox = 2
    [_Last] = 2
    [_Count] = 3
End Enum

#If False Then
    Private Const PDT_LeftToolbox = 0, PDT_BottomToolbox = 1, PDT_RightToolbox = 2
#End If

'At present, PD tracks three toolbars in a hard-coded order: the main toolbox (left), the options toolbox (bottom), and the
' layer toolbox (right).
Private m_Toolboxes() As PD_Toolbox_Data

'Before loading any toolboxes, call this sub to populate the initial toolbox data.  Among other thing, this loads the previous toolbox
' sizes from the user's preferences file.
Public Sub LoadToolboxData()

    Dim numTools As PD_Toolbox: numTools = [_Last]
    ReDim m_Toolboxes(0 To numTools) As PD_Toolbox_Data
    
    'Set default, minimum, and maximum sizes for each toolbox.  Default sizes are only used the first time PD is run,
    ' while min and max only apply if the user click-drags the region to some new size.
    '
    'Note that we also account for DPI at run-time.
    FillDefaultToolboxValues
    
    Dim i As PD_Toolbox, newSize As Long
    For i = [_First] To [_Last]
        With m_Toolboxes(i)
            .IsVisiblePreference = g_UserPreferences.GetPref_Boolean("Toolbox", GetToolboxName(i) & "Visible", True)
            .IsVisibleNow = .IsVisiblePreference
            
            'Apply a failsafe size check to the previous session's saved value
            newSize = g_UserPreferences.GetPref_Long("Toolbox", GetToolboxName(i) & "Size", .DefaultSize)
            If newSize < .MinSize Then newSize = .MinSize
            If newSize > .MaxSize Then newSize = .MaxSize
            .ConstrainingSize = newSize
            
        End With
    Next i
    
End Sub

'Before unloading the toolboxes, call this sub to write our current toolbox data out to the user's preference file.
Public Sub SaveToolboxData()

    Dim i As PD_Toolbox
    For i = [_First] To [_Last]
        With m_Toolboxes(i)
            g_UserPreferences.SetPref_Boolean "Toolbox", GetToolboxName(i) & "Visible", .IsVisiblePreference
            g_UserPreferences.SetPref_Long "Toolbox", GetToolboxName(i) & "Size", .ConstrainingSize
        End With
    Next i

End Sub

'All hard-coded toolbox values should be handled in this sub - NOWHERE else!  (Otherwise, code maintenance becomes very unpleasant.)
Private Sub FillDefaultToolboxValues()
    
    Dim i As PD_Toolbox
    For i = [_First] To [_Last]
        With m_Toolboxes(i)
            Select Case i
            
                Case PDT_LeftToolbox
                    .DefaultSize = FixDPI(142)
                    .MinSize = FixDPI(48)
                    .MaxSize = FixDPI(200)      'ARBITRARY VALUE!  TODO: figure out a meaningful number
                
                Case PDT_BottomToolbox
                    .DefaultSize = FixDPI(100)
                    .MinSize = 0                'The bottom toolbox is unique in not being user-sizable.  It is autosized according
                    .MaxSize = FixDPI(500)      ' to the requirements of each tool, so these are basically just dummy values.
                
                Case PDT_RightToolbox
                    .DefaultSize = FixDPI(250)      'This seems large - given our new UI tools, let's see if we can shrink this a bit
                    .MinSize = FixDPI(200)
                    .MaxSize = FixDPI(400)      'ARBITRARY VALUE!  TODO: figure out a meaningful number
                    
            End Select
        End With
    Next i
    
End Sub

'Preferences are written in XML format, so we need string representations of each toolbox's title
Private Function GetToolboxName(ByVal toolID As PD_Toolbox) As String
    Select Case toolID
        Case PDT_LeftToolbox
            GetToolboxName = "LeftToolbox"
        Case PDT_BottomToolbox
            GetToolboxName = "BottomToolbox"
        Case PDT_RightToolbox
            GetToolboxName = "RightToolbox"
    End Select
End Function

'If the main form's position changes in some way, call this function to calculate new position rects for each toolbox.
' Also, return the (subsequent) position rect that's left over, which is where the primary canvas control will go!
' IMPORTANT NOTE: *both* rects passed to this function will potentially be modified, so make backups if you need them!
Public Sub CalculateNewToolboxRects(ByRef mainFormClientRect As winRect, ByRef dstCanvasRect As winRect)

    'It sounds weird, but we actually calculate the bottom toolbox's rect first, since it gets positioning preference.
    With m_Toolboxes(PDT_BottomToolbox)
        If .IsVisibleNow Then
        
            .toolRect.x1 = mainFormClientRect.x1
            .toolRect.x2 = mainFormClientRect.x2
            .toolRect.y2 = mainFormClientRect.y2
            .toolRect.y1 = mainFormClientRect.y2 - .ConstrainingSize
            
            'As each toolbar is positioned, we update the client rect we received to reflect the new positions.
            mainFormClientRect.y2 = .toolRect.y1
            
        End If
    End With
    
    'Left and right toolboxes use basically identical code, and their size algorithms should be self-explanatory.  The key thing
    ' to remember is that the *bottom* of these toolbars is determined by the *top* of the bottom toolbar.
    With m_Toolboxes(PDT_LeftToolbox)
        If .IsVisibleNow Then
            .toolRect.x1 = mainFormClientRect.x1
            .toolRect.x2 = .toolRect.x1 + .ConstrainingSize
            .toolRect.y1 = mainFormClientRect.y1
            .toolRect.y2 = mainFormClientRect.y2
            mainFormClientRect.x1 = .toolRect.x2
        End If
    End With
    
    With m_Toolboxes(PDT_RightToolbox)
        If .IsVisibleNow Then
            .toolRect.x1 = mainFormClientRect.x2 - .ConstrainingSize
            .toolRect.x2 = mainFormClientRect.x2
            .toolRect.y1 = mainFormClientRect.y1
            .toolRect.y2 = mainFormClientRect.y2
            mainFormClientRect.x2 = .toolRect.x1
        End If
    End With
    
    'Add 1-pixel's worth of padding to all affected sides of the canvas rect (e.g. the top can stay where it is,
    ' as there is no neighboring toolbox).
    With dstCanvasRect
        .x1 = mainFormClientRect.x1
        .x2 = mainFormClientRect.x2
        .y1 = mainFormClientRect.y1
        .y2 = mainFormClientRect.y2
    End With
    
End Sub

'Show and/or position a toolbox according to its current settings.  In most cases, you will want to call
' CalculateNewToolboxRects(), above, prior to invoking this function.
Public Sub PositionToolbox(ByVal toolID As PD_Toolbox, ByVal toolboxHWnd As Long, ByVal parentHwnd As Long)
    SetParent toolboxHWnd, parentHwnd
    With m_Toolboxes(toolID)
        .hWnd = toolboxHWnd
        If .IsVisibleNow Then
            MoveWindow toolboxHWnd, .toolRect.x1, .toolRect.y1, .toolRect.x2 - .toolRect.x1, .toolRect.y2 - .toolRect.y1, 1&
            ShowWindow toolboxHWnd, SW_SHOWNA
        Else
            ShowWindow toolboxHWnd, SW_HIDE
            MoveWindow toolboxHWnd, .toolRect.x1, .toolRect.y1, .toolRect.x2 - .toolRect.x1, .toolRect.y2 - .toolRect.y1, 0&
        End If
    End With
End Sub

'If a toolbox is resized by the user, you must call this function to notify other windows of the change.
' This function will return the actual size used, which may be different if the passed size is too large or too small.
Public Function SetConstrainingSize(ByVal toolID As PD_Toolbox, ByVal newSize As Long) As Long
    With m_Toolboxes(toolID)
        If newSize < .MinSize Then newSize = .MinSize
        If newSize > .MaxSize Then newSize = .MaxSize
        .ConstrainingSize = newSize
        SetConstrainingSize = newSize
    End With
End Function

'To notify us of a change to a toolbox's visibility that *doesn't* affect the user's preference (e.g. for auto-hide behavior),
' use this function.  Note that, by design, this function does not *apply* the setting.  You must separately call the
' CalculateNewToolboxRects() function for this, as other toolbox's position may be affected by this visibility change as well.
Public Sub SetToolboxVisibility(ByVal toolID As PD_Toolbox, ByVal newSetting As Boolean)
    m_Toolboxes(toolID).IsVisibleNow = newSetting
End Sub

'This function is named similarly to the one beneath it, but it's behavior is quite different.  Basically, this function just
' sets a window's visibility so that it matches the user's preference.  This is relevant for the bottom Options toolbox,
' which is forcibly hidden for some tools (e.g. the hand), but shown according to preference for other tools.
Public Sub SetToolboxVisibilityByPreference(ByVal toolID As PD_Toolbox)
    m_Toolboxes(toolID).IsVisibleNow = m_Toolboxes(toolID).IsVisiblePreference
End Sub

'If the user has changed their actual preference for toolbox visibility (e.g. by clicking the corresponding menu entry),
' call this function.  It will update the master preferences file with the new setting, but it *WILL NOT* actually show/hide
' the window.  You must manually call CalculateNewToolboxRects() first, as showing/hiding a toolbox has repercussions on
' neighboring windows that must be considered too.
'
' *IF*, however, you just want to show/hide a tool window as part of auto-hide behavior, use SetToolboxVisibility(), above.
Public Sub SetToolboxVisibilityPreference(ByVal toolID As PD_Toolbox, ByVal newSetting As Boolean)
    With m_Toolboxes(toolID)
        .IsVisiblePreference = newSetting
        .IsVisibleNow = newSetting
    End With
End Sub

'Toolbars can be dynamically shown/hidden by a variety of processes (e.g. clicking an entry in the Window menu, clicking the X in a
' toolbar's command box, etc).  All those operations should wrap this singular function.
Public Sub ToggleToolboxVisibility(ByVal whichToolbar As PD_Toolbox, Optional ByVal suppressRedraws As Boolean = False)

    Select Case whichToolbar
    
        Case PDT_LeftToolbox
            FormMain.MnuWindowToolbox(0).Checked = Not FormMain.MnuWindowToolbox(0).Checked
            SetToolboxVisibilityPreference PDT_LeftToolbox, FormMain.MnuWindowToolbox(0).Checked
            
        Case PDT_BottomToolbox
            FormMain.MnuWindow(1).Checked = Not FormMain.MnuWindow(1).Checked
            SetToolboxVisibilityPreference PDT_BottomToolbox, FormMain.MnuWindowToolbox(1).Checked
            
            'Because this toolbox's visibility is also tied to the current tool, we wrap a different function.  This function
            ' will show/hide the toolbox as necessary.
            toolbar_Toolbox.ResetToolButtonStates
            
        Case PDT_RightToolbox
            FormMain.MnuWindow(2).Checked = Not FormMain.MnuWindow(2).Checked
            SetToolboxVisibilityPreference PDT_RightToolbox, FormMain.MnuWindowToolbox(2).Checked
            
    End Select
    
    'NEW SYSTEM: the below line can stay, but we need to remove the "loaded images" check.  Even if no images are loaded,
    ' we need to reset the canvas area.
    
    'Redraw the primary image viewport, as the available client area may have changed.
    If (Not suppressRedraws) Then FormMain.UpdateMainLayout
    
End Sub
