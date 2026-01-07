Attribute VB_Name = "Toolboxes"
'***************************************************************************
'Toolbox Manager
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 06/March/16
'Last update: split toolbox-specific code out of the (now very large) Interface module and into this dedicated home.
'
'Miscellaneous routines related to rendering and handling PhotoDemon's toolboxes.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As Long) As Long
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const GWL_STYLE As Long = (-16)
Private Const SW_HIDE As Long = 0&
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
    PDT_TopToolbox = 1
    PDT_RightToolbox = 2
    [_Last] = 2
    [_Count] = 3
End Enum

#If False Then
    Private Const PDT_LeftToolbox = 0, PDT_TopToolbox = 1, PDT_RightToolbox = 2
#End If

'At present, PD tracks three toolbars in a hard-coded order: the main toolbox (left), the options toolbox (top), and the
' layer toolbox (right).
Private m_Toolboxes() As PD_Toolbox_Data

'hWnd and window bits are stored using basic key/value pairs.  We must reset these before exiting the program
' or VB will crash.
Private m_windowBits As pdDictionary

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
            .IsVisiblePreference = UserPrefs.GetPref_Boolean("Toolbox", GetToolboxName(i) & "Visible", True)
            .IsVisibleNow = .IsVisiblePreference
            
            'Retrieve the last-set size for the toolbox, then adjust it to compensate for any per-session DPI changes
            newSize = UserPrefs.GetPref_Long("Toolbox", GetToolboxName(i) & "Size", .DefaultSize)
            
            Dim lastDPI As Single
            lastDPI = Interface.GetLastSessionDPI_Ratio()
            newSize = newSize * (Interface.GetSystemDPIRatio() / lastDPI)
            
            'Apply a failsafe size check to the previous session's adjusted value
            If (newSize < .MinSize) Then newSize = .MinSize
            If (newSize > .MaxSize) Then newSize = .MaxSize
            .ConstrainingSize = newSize
            
        End With
    Next i
    
End Sub

'Before unloading the toolboxes, call this sub to write our current toolbox data out to the user's preference file.
Public Sub SaveToolboxData()

    If PDMain.WasStartupSuccessful() Then

        Dim i As PD_Toolbox
        For i = [_First] To [_Last]
            With m_Toolboxes(i)
                UserPrefs.SetPref_Boolean "Toolbox", GetToolboxName(i) & "Visible", .IsVisiblePreference
                UserPrefs.SetPref_Long "Toolbox", GetToolboxName(i) & "Size", .ConstrainingSize
            End With
        Next i
        
        UserPrefs.SetPref_Float "Toolbox", "LastSessionDPI", Interface.GetSystemDPIRatio()
        
    End If

End Sub

'This function purely exists so we can properly synchronize the main Window menu checkmarks to the saved toolbox visibility preferences
Public Function GetToolboxVisibilityPreference(ByVal toolID As PD_Toolbox) As Boolean
    GetToolboxVisibilityPreference = m_Toolboxes(toolID).IsVisiblePreference
End Function

Public Function GetToolboxVisibility(ByVal toolID As PD_Toolbox) As Boolean
    GetToolboxVisibility = m_Toolboxes(toolID).IsVisibleNow
End Function

Public Function GetToolboxMinWidth(ByVal toolID As PD_Toolbox) As Long
    GetToolboxMinWidth = m_Toolboxes(toolID).MinSize
End Function

Public Function AreAllToolboxesHidden() As Boolean
    Dim atLeastOneBoxVisible As Boolean: atLeastOneBoxVisible = False
    Dim i As PD_Toolbox
    For i = [_First] To [_Last]
        atLeastOneBoxVisible = atLeastOneBoxVisible Or Toolboxes.GetToolboxVisibility(i)
    Next i
    AreAllToolboxesHidden = (Not atLeastOneBoxVisible)
End Function

'All hard-coded toolbox values should be handled in this sub - NOWHERE else!  (Otherwise, code maintenance becomes very unpleasant.)
Private Sub FillDefaultToolboxValues()
    
    Dim i As PD_Toolbox
    For i = [_First] To [_Last]
        With m_Toolboxes(i)
            Select Case i
            
                Case PDT_LeftToolbox
                    .DefaultSize = FixDPI(98)
                    .MinSize = FixDPI(48)
                    .MaxSize = FixDPI(188)
                
                'The top toolbox is unique in not being user-sizable.  It is a fixed height with individual
                ' drop-down panels that can extend portions of the toolbox vertically.  This base height could
                ' theoretically vary by tool, but for now it is fixed to ensure a more aesthetically pleasing
                ' layout (and simpler UI design).
                Case PDT_TopToolbox
                    Const TOP_TOOLBOX_HEIGHT As Long = 59
                    .DefaultSize = FixDPI(TOP_TOOLBOX_HEIGHT)
                    .MinSize = FixDPI(TOP_TOOLBOX_HEIGHT)
                    .MaxSize = FixDPI(TOP_TOOLBOX_HEIGHT)
                
                Case PDT_RightToolbox
                    .DefaultSize = FixDPI(190)
                    .MinSize = FixDPI(174)
                    .MaxSize = FixDPI(360)
                    
            End Select
        End With
    Next i
    
End Sub

'Preferences are written in XML format, so we need string representations of each toolbox's title
Private Function GetToolboxName(ByVal toolID As PD_Toolbox) As String
    Select Case toolID
        Case PDT_LeftToolbox
            GetToolboxName = "LeftToolbox"
        Case PDT_TopToolbox
            GetToolboxName = "BottomToolbox"    'For backward compatibility, this is left as "bottom" despite
                                                ' now appearing at the top.
        Case PDT_RightToolbox
            GetToolboxName = "RightToolbox"
    End Select
End Function

'If the main form's position changes in some way, call this function to calculate new position rects for each toolbox.
' Also, return the (subsequent) position rect that's left over, which is where the primary canvas control will go!
' IMPORTANT NOTE: *both* rects passed to this function will potentially be modified, so make backups if you need them!
Public Sub CalculateNewToolboxRects(ByRef mainFormClientRect As winRect, ByRef dstCanvasRect As winRect)
    
    'Left toolbox is calculated first.  It always fills all available vertical space, regardless of layout.
    ' (This greatly simplifies rendering.)
    If m_Toolboxes(PDT_LeftToolbox).IsVisibleNow Then
        With m_Toolboxes(PDT_LeftToolbox)
            .toolRect.x1 = mainFormClientRect.x1
            .toolRect.x2 = .toolRect.x1 + .ConstrainingSize
            .toolRect.y1 = mainFormClientRect.y1
            .toolRect.y2 = mainFormClientRect.y2
            
            'As each toolbar is positioned, we update the client rect we received to reflect the new positions.
            mainFormClientRect.x1 = .toolRect.x2
        End With
    End If
    
    'Right toolbox next.  It also fills all vertical space.
    If m_Toolboxes(PDT_RightToolbox).IsVisibleNow Then
        With m_Toolboxes(PDT_RightToolbox)
            .toolRect.x1 = mainFormClientRect.x2 - .ConstrainingSize
            .toolRect.x2 = mainFormClientRect.x2
            .toolRect.y1 = mainFormClientRect.y1
            .toolRect.y2 = mainFormClientRect.y2
            mainFormClientRect.x2 = .toolRect.x1
        End With
    End If
    
    'Top toolbox goes next.  It is sandwiched between the other two toolboxes.
    If m_Toolboxes(PDT_TopToolbox).IsVisibleNow Then
        With m_Toolboxes(PDT_TopToolbox)
            .toolRect.x1 = mainFormClientRect.x1
            .toolRect.x2 = mainFormClientRect.x2
            .toolRect.y1 = mainFormClientRect.y1
            .toolRect.y2 = mainFormClientRect.y1 + .ConstrainingSize
            
            'Shift the main client area down
            mainFormClientRect.y1 = .toolRect.y2
        End With
    End If
    
    'Reflect the final calculations back to the main canvas
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
    
    'Failsafe only
    If (g_WindowManager Is Nothing) Then Exit Sub
    
    SetParent toolboxHWnd, parentHwnd
    
    'Cache default VB6 window bits (only the first time!), then set new window bits matching the
    ' parent/child relationship we just established.
    If (m_windowBits Is Nothing) Then Set m_windowBits = New pdDictionary
    If (Not m_windowBits.DoesKeyExist(toolboxHWnd)) Then m_windowBits.AddEntry toolboxHWnd, GetWindowLong(toolboxHWnd, GWL_STYLE)
    SetWindowLong toolboxHWnd, GWL_STYLE, GetWindowLong(toolboxHWnd, GWL_STYLE) Or WS_CHILD
    SetWindowLong toolboxHWnd, GWL_STYLE, GetWindowLong(toolboxHWnd, GWL_STYLE) And (Not WS_POPUP)
    
    With m_Toolboxes(toolID)
        
        If (.hWnd <> toolboxHWnd) Then
            .hWnd = toolboxHWnd
            
            If (toolID = PDT_LeftToolbox) Then
                g_WindowManager.RequestMinMaxTracking toolboxHWnd, toolID, .MinSize, , .MaxSize
            ElseIf (toolID = PDT_TopToolbox) Then
                g_WindowManager.RequestMinMaxTracking toolboxHWnd, toolID, , .MinSize, , .MaxSize
            ElseIf (toolID = PDT_RightToolbox) Then
                g_WindowManager.RequestMinMaxTracking toolboxHWnd, toolID, .MinSize, , .MaxSize
            End If
            
        End If
        
        'Prior to making any visibility changes, we need to make note of the currently focused item.  If we don't,
        ' Windows may inadvertently redirect focus somewhere bizarre.
        Dim focusHWnd As Long
        focusHWnd = g_WindowManager.GetFocusAPI()
        
        If .IsVisibleNow Then
            MoveWindow toolboxHWnd, .toolRect.x1, .toolRect.y1, .toolRect.x2 - .toolRect.x1, .toolRect.y2 - .toolRect.y1, 1&
            ShowWindow toolboxHWnd, SW_SHOWNA
        Else
            ShowWindow toolboxHWnd, SW_HIDE
            MoveWindow toolboxHWnd, .toolRect.x1, .toolRect.y1, .toolRect.x2 - .toolRect.x1, .toolRect.y2 - .toolRect.y1, 0&
        End If
        
        'Restore focus to its original window (but only if it's visible; otherwise, the main form gets focus)
        If g_WindowManager.GetVisibilityByHWnd(focusHWnd) Then g_WindowManager.SetFocusAPI focusHWnd Else g_WindowManager.SetFocusAPI FormMain.hWnd
        
    End With
    
End Sub

'If a toolbox is resized by the user, you must call this function to notify other windows of the change.
' This function will return the actual size used, which may be different if the passed size is too large or too small.
Public Function SetConstrainingSize(ByVal toolID As PD_Toolbox, ByVal newSize As Long) As Long
    With m_Toolboxes(toolID)
        If (newSize < .MinSize) Then newSize = .MinSize
        If (newSize > .MaxSize) Then newSize = .MaxSize
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

'This function is named similarly to the one beneath it, but remember this critical difference:
' this function simply sets a window's visibility *matching the current user preference*.
' This is relevant for the top "Tool Options" toolbox, which is forcibly hidden for some tools
' (e.g. the hand/pan tool), but shown according to preference for other tools.
Public Sub SetToolboxVisibilityByPreference(ByVal toolID As PD_Toolbox)
    m_Toolboxes(toolID).IsVisibleNow = m_Toolboxes(toolID).IsVisiblePreference
End Sub

'If the user has changed their actual preference for toolbox visibility (e.g. by clicking
' the corresponding menu), call this function.  It will update the central user preferences file
' with the new setting, but it *WILL NOT* actually show/hide the toolbox.  You must manually call
' CalculateNewToolboxRects() first, because changing toolbox visibility has repercussions on
' neighboring windows.
'
' *IF*, however, you just want to show/hide a tool window as part of auto-hide behavior,
' use SetToolboxVisibility(), above.
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
            FormMain.MnuWindowToolbox(0).Checked = (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            SetToolboxVisibilityPreference whichToolbar, (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            
        Case PDT_TopToolbox
            FormMain.MnuWindow(1).Checked = (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            SetToolboxVisibilityPreference PDT_TopToolbox, (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            
            'Because this toolbox's visibility is also tied to the current tool, we wrap a different function.  This function
            ' will show/hide the toolbox as necessary.
            toolbar_Toolbox.ResetToolButtonStates
            
        Case PDT_RightToolbox
            FormMain.MnuWindow(2).Checked = (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            SetToolboxVisibilityPreference PDT_RightToolbox, (Not m_Toolboxes(whichToolbar).IsVisiblePreference)
            
    End Select
    
    'Redraw the primary image viewport, as the available client area may have changed.
    If (Not suppressRedraws) Then FormMain.UpdateMainLayout
    
End Sub

Public Sub ResetAllToolboxSettings()
    
    'Reset all toolbox sizes to their initial defaults
    FillDefaultToolboxValues
    
    'Set current toolbox sizes to match their (freshly restored) default values
    Dim i As PD_Toolbox
    For i = [_First] To [_Last]
        With m_Toolboxes(i)
            .ConstrainingSize = .DefaultSize
            .IsVisibleNow = True
            .IsVisiblePreference = True
        End With
    Next i
    
    'Reset all left toolbar settings
    toolbar_Toolbox.ToggleToolCategoryLabels PD_BOOL_TRUE
    toolbar_Toolbox.UpdateButtonSize tbs_Small
    
    'The left-side toolbox is a little finicky because it auto-locks its width to match precise intervals
    ' of its current button size. To simplify the process of resetting its settings, forcibly set its
    ' width now, *before* calling the central canvas layout update function.
    If (Not g_WindowManager Is Nothing) Then
        With m_Toolboxes(PDT_LeftToolbox)
            g_WindowManager.SetVisibilityByHWnd .hWnd, True
            g_WindowManager.SetSizeByHWnd .hWnd, .DefaultSize, g_WindowManager.GetClientHeight(.hWnd), True
        End With
    End If
    
    'Sync various menus to reflect the new settings
    FormMain.MnuWindowToolbox(0).Checked = True
    FormMain.MnuWindow(1).Checked = True
    FormMain.MnuWindow(2).Checked = True
    
    'Ensure that the options toolbox is properly shown/hidden depending on the currently selected tool
    toolbar_Toolbox.ResetToolButtonStates
    
    'Reset all image strip settings
    Interface.ToggleImageTabstripVisibility 1
    Interface.ToggleImageTabstripAlignment vbAlignTop
    
    'Some toolboxes may need to perform their own internal resets (e.g. the layer panel needs to reset
    ' individual panel sizes)
    toolbar_Layers.ResetInterface
    
    'Redraw the primary image viewport to reflect our many potential changes
    FormMain.UpdateMainLayout True

End Sub

'Before unloading a toolbox, call this function to unload it.  (If you don't do this, VB will crash!)
Public Sub ReleaseToolbox(ByVal toolboxHWnd As Long)
    If (Not m_windowBits Is Nothing) Then
        If m_windowBits.DoesKeyExist(toolboxHWnd) Then
            SetWindowLong toolboxHWnd, GWL_STYLE, m_windowBits.GetEntry_Long(toolboxHWnd, GetWindowLong(toolboxHWnd, GWL_STYLE))
            m_windowBits.DeleteEntry toolboxHWnd
        End If
    End If
End Sub
