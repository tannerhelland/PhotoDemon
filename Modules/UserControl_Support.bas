Attribute VB_Name = "UserControl_Support"
'***************************************************************************
'Helper functions for various PhotoDemon UCs
'Copyright 2014-2017 by Tanner Helland
'Created: 06/February/14
'Last updated: 20/August/15
'Last update: start migrating various UC-inspecific functions here
'
'Many of PD's custom user controls share similar functionality.  Rather than duplicate that functionality across
' all controls, I've tried to collect reusable functions here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Type PD_LISTITEM
    textEn As String
    textTranslated As String
    itemTop As Long
    itemHeight As Long
    isSeparator As Boolean
End Type

'The tab key presents a particular problem when intermixing API windows and VB controls.  VB (obviously) ignores the
' API windows entirely, and PD further complicates this by sometimes mixing API windows and VB controls on the same UC.
' To avoid this disaster, we manage our own tab key presses using an automated system that sorts controls from top-left
' to bottom-right, and automatically figures out tab order from there.
'
'To make sure the automated system works correctly, some controls actually raise a TabPress event, which their parent
' UC can use to cycle control focus within the UC.  When focus is tabbed-out from the last control on the UC, the UC
' itself can then notify the master TabHandler to pass focus to an entirely new control.
Public Enum PDUC_TAB_BEHAVIOR
    TabDefaultBehavior = 0
    TabRaiseEvent = 1
End Enum

#If False Then
    Private Const TabDefaultBehavior = 0, TabRaiseEvent = 1
#End If

'At times, PD may need to post custom messages to all application windows (e.g. theme changes may eventually be implemented
' like this).  Do not call PostMessage directly, as it sends messages to the thread's message queue; instead, call the
' PostPDMessage() function below, which asynchronously relays the request to registered windows via SendNotifyMessage.
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Current list of registered windows and the custom messages they want to receive.  This spares us from having to enumerate
' all windows, or worse, blast all windows in the system with our internal messages.  (At present, these are naive lists
' because PD uses so few of them, but in the future, we could look at a hash table.  I've deliberately made the list
' interactions structure-agnostic to simplify future improvements.)
Private m_windowList() As Long, m_wMsgList() As Long
Private m_windowMsgCount As Long
Private Const INITIAL_WINDOW_MESSAGE_LIST_SIZE As Long = 16&

'Current list of shared GDI brushes; this spares us from creating unique brushes for every edit box instance
Private Type SharedGDIBrush
    brushColor As Long
    brushHandle As Long
    numOfOwners As Long
End Type

Private m_numOfSharedBrushes As Long
Private m_SharedBrushes() As SharedGDIBrush
Private Const INIT_SIZE_OF_BRUSH_CACHE As Long = 4&

'Current list of shared GDI fonts; this spares us from creating unique font for every edit box instance.
' (Note that we don't use a pdFontCollection object for this, as these fonts are immutable.  If we accidentally
'  destroy one before its matching edit box is freed, we will crash and burn, so a different handling technique
'  is requred.)
Private Type SharedGDIFont
    FontSize As Single
    fontHandle As Long
    numOfOwners As Long
End Type

Private m_numOfSharedFonts As Long
Private m_SharedFonts() As SharedGDIFont
Private Const INIT_SIZE_OF_FONT_CACHE As Long = 4&

'Sometimes controls need unique ID values.  This module will provide a non-zero unique ID via the GetUniqueControlID() function.
Private m_UniqueIDTracker As Long

'As part of broad optimization efforts in the 7.0 release, this module now tracks how many custom PD controls we're managing at
' any given time.  Use this for leak-detection and resource counting.  For example: each ucSupport-managed PD control uses two
' GDI objects: one DIB and one persistent DC for the control's backbuffer (all controls are double-buffered).  Use this to
' figure out how many of the program's GDI objects are being used by UCs, and how many are being created and used elsewhere.
Private m_PDControlCount As Long

'Dropdown boxes are problematic, because we have to play some weird window ownership games to ensure that the dropdowns
' appear "above" or "outside" VB windows, as necessary.  As such, this function is notified whenever a listbox is raised,
' and the hWnd is cached so we can kill that window as necessary.
Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hndWindow As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal ptrToRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private m_CurrentDropDownHWnd As Long, m_CurrentDropDownListHWnd As Long

'To better manage resources, we also track how many API windows we've created/destroyed during this session
Private m_APIWindowsCreated As Long, m_APIWindowsDestroyed As Long

'Same goes for timers and certain other object types
Private m_TimersCreated As Long, m_TimersDestroyed As Long
Private m_HooksCreated As Long, m_HooksDestroyed As Long

'Because there can only be one visible tooltip at a time, this support module is a great place to handle them.  Requests for new
' tooltips automatically unload old ones, although user controls still need to request tooltip hiding when they lose focus and/or
' are unloaded.
'
'Note that all tooltip sizes are in pixels at 96 DPI ("100%" in Windows).  Other DPIs will automatically be handled at run-time,
' as necessary.
Private Const PD_TT_EXTERNAL_PADDING As Long = 2
Private Const PD_TT_INTERNAL_PADDING As Long = 6
Private Const PD_TT_MAX_WIDTH As Long = 400         'Tips larger than this will be word-wrapped to fit.
Private Const PD_TT_TITLE_PADDING As Long = 4       'Pixels between the tip title (if any) and caption
Private Const AW_BLEND As Long = &H80000
Private Const AW_HIDE As Long = &H10000
Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOCOPYBITS As Long = &H100
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_TOPMOST As Long = &H8
Private m_TTActive As Boolean, m_TTOwner As Long, m_TTHwnd As Long
Private m_TTWindowStyleHasBeenSet As Boolean, m_OriginalTTWindowBits As Long, m_OriginalTTWindowBitsEx As Long
Private m_TTRectCopy As RECTL

'Tooltips can request icons, but as of 7.0, this behavior is unused.  (As such, these are not currently implemented.)
Public Enum TT_ICON_TYPE
    TTI_NONE = 0
    TTI_INFO = 1
    TTI_WARNING = 2
    TTI_ERROR = 3
End Enum

#If False Then
    Private Const TTI_NONE = 0, TTI_INFO = 1, TTI_WARNING = 2, TTI_ERROR = 3
#End If

'To optimize tooltip positioning, we determine the tooltip edge closest to the mouse hover position (with the assumption that
' the user's eyes will be pointed at or near the cursor, so the closer a tooltip is to that, the better - without obscuring
' the control underneath, obviously)
Private Enum TT_SIDE
    TTS_Top = 0
    TTS_Right = 1
    TTS_Bottom = 2
    TTS_Left = 3
End Enum

#If False Then
    Private Const TTS_Top = 0, TTS_Right = 1, TTS_Bottom = 2, TTS_Left = 3
#End If

'Tooltips are hidden based on a timer.  If a new tooltip is requested before the timer expires, we simply move the existing window
' into place, rather than animating it.
Private m_TimerEventSink As pdUCEventSink
Private m_InitialTTTimerTime As Double

'Iterate through all sibling controls in our container, and if one is capable of receiving focus, activate it.  I had *really* hoped
' to bypass this kind of manual handling by using WM_NEXTDLGCTL, but I failed to get it working reliably with all types of VB windows.
' I'm honestly not sure whether VB even uses that message, or whether it uses some internal mechanism for focus tracking; the latter
' might explain some of the bugginess involved in VB focus handling.
'
'Returns: TRUE if focus was moved; FALSE if it was left behind.
Public Function ForwardFocusToNewControl(ByRef sourceControl As Object, ByVal focusDirectionForward As Boolean) As Boolean

    'When sited on a UC, we may not be able to iterate a controls collection.  Simply exit if this occurs.
    ' (Note that this is only a temporary solution; I'm working on a better one, where UC's cascade tab handling to their parent
    '  until an acceptable handler is reached.)
    On Error GoTo ParentHasNoControls

    'If the user has deactivated tab support, or we are invisible/disabled, ignore this completely
    If sourceControl.Extender.TabStop And sourceControl.Extender.Visible And sourceControl.Enabled Then
        
        'Iterate through all controls in the container, looking for the next TabStop index
        Dim myIndex As Long
        myIndex = sourceControl.Extender.TabIndex
        
        Dim newIndex As Long
        Const MAX_INDEX As Long = 99999
        
        'Forward and back focus checks require different search strategies
        If focusDirectionForward Then
            newIndex = MAX_INDEX
        Else
            newIndex = myIndex
        End If
        
        Dim Ctl As Control, targetControl As Control
        For Each Ctl In sourceControl.Parent.Controls
        
        'Some controls may not have a TabStop property.  That's okay - just keep iterating if it happens.
        On Error GoTo 0
        On Error GoTo NextControlCheck
            
            'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
            ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
            ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
            ' is focusable.
            If IsControlFocusable(Ctl) Then
            
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
            On Error GoTo 0
            On Error GoTo NextControlCheck2
            
            'If our control is last in line for tabstops, we need to now find the LOWEST tab index to forward focus to.
            For Each Ctl In sourceControl.Parent.Controls
                
                'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
                ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
                ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
                ' is focusable.
                If IsControlFocusable(Ctl) Then
                    
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
        On Error GoTo 0
        On Error GoTo NoFocusRecipient
        
        'Ignore the second case completely, as tab should have no effect
        If newIndex <> myIndex Then
            targetControl.SetFocus
        
NoFocusRecipient:
        
        End If
        
    End If
    
    ForwardFocusToNewControl = True
    
    Exit Function
    
ParentHasNoControls:

    ForwardFocusToNewControl = False

End Function

'If an object of type Control is capable of receiving focus, this will return TRUE.  We use this to make sure focus-setting
' doesn't break by accidentally setting focus to something we shouldn't.
Public Function IsControlFocusable(ByRef Ctl As Control) As Boolean

    If Not (TypeOf Ctl Is Timer) And Not (TypeOf Ctl Is Line) And Not (TypeOf Ctl Is pdLabel) And Not (TypeOf Ctl Is Frame) And Not (TypeOf Ctl Is Shape) And Not (TypeOf Ctl Is Image) And Not (TypeOf Ctl Is pdAccelerator) And Not (TypeOf Ctl Is ShellPipe) And Not (TypeOf Ctl Is pdDownload) Then
        IsControlFocusable = True
    Else
        IsControlFocusable = False
    End If

End Function

'PD's various user controls sometimes like to share data via custom window messages.  Instead of calling PostMessage directly,
' use this wrapper function, which may perform additional maintenance.
Public Sub PostPDMessage(ByVal wMsg As Long, Optional ByVal wParam As Long = 0&, Optional ByVal lParam As Long = 0&)
    
    Dim pmReturn As Long
    pmReturn = 1
    
    'Enumerate all matching, non-zero windows, and post the requested message, without waiting for a response.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If (m_wMsgList(i) = wMsg) Then
            If (m_windowList(i) <> 0) Then
                pmReturn = pmReturn And SendNotifyMessage(m_windowList(i), wMsg, wParam, lParam)
            End If
        End If
    Next i
    
    #If DEBUGMODE = 1 Then
        If pmReturn = 0 Then
            pdDebug.LogAction "PostPDMessage was unable to post message ID #" & wMsg & " to one or more windows."
        End If
    #End If
    
End Sub

'Rather than blast all windows with manually raised messages, PD maintains a list of hWnds and registered message requests.
' Add windows and/or messages via this function, and when the messages need to be raised (via PostPDMessage(), above),
' the function will automatically notify all registered recipients.
Public Sub AddMessageRecipient(ByVal targetHwnd As Long, ByVal wMsg As Long)
    
    'Prep the storage structure, as necessary.
    If (m_windowMsgCount = 0) Then
        ReDim m_windowList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
        ReDim m_wMsgList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
    ElseIf m_windowMsgCount > UBound(m_windowList) Then
        ReDim m_windowList(0 To (UBound(m_windowList) * 2 + 1)) As Long
        ReDim m_wMsgList(0 To (UBound(m_wMsgList) * 2 + 1)) As Long
    End If
    
    m_windowList(m_windowMsgCount) = targetHwnd
    m_wMsgList(m_windowMsgCount) = wMsg
    
    m_windowMsgCount = m_windowMsgCount + 1
    
End Sub

Public Sub RemoveMessageRecipient(ByVal targetHwnd As Long)
    
    'Rather then condensing the list, we simply set all corresponding window entries to zero.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If m_windowList(i) = targetHwnd Then
            m_windowList(i) = 0
            m_wMsgList(i) = 0
        End If
    Next i
    
End Sub

Public Sub NotifyAPIWindowCreated()
    m_APIWindowsCreated = m_APIWindowsCreated + 1
End Sub

Public Sub NotifyAPIWindowDestroyed()
    m_APIWindowsDestroyed = m_APIWindowsDestroyed + 1
End Sub

Public Function GetAPIWindowCount(Optional ByRef windowsCreated As Long, Optional ByRef windowsDestroyed As Long) As Long
    windowsCreated = m_APIWindowsCreated
    windowsDestroyed = m_APIWindowsDestroyed
    GetAPIWindowCount = (windowsCreated - windowsDestroyed)
End Function

Public Sub NotifyTimerCreated()
    m_TimersCreated = m_TimersCreated + 1
End Sub

Public Sub NotifyTimerDestroyed()
    m_TimersDestroyed = m_TimersDestroyed + 1
End Sub

Public Function GetTimerCount(Optional ByRef timersCreated As Long, Optional ByRef timersDestroyed As Long) As Long
    timersCreated = m_TimersCreated
    timersDestroyed = m_TimersDestroyed
    GetTimerCount = (timersCreated - timersDestroyed)
End Function

Public Sub NotifyHookCreated()
    m_HooksCreated = m_HooksCreated + 1
End Sub

Public Sub NotifyHookDestroyed()
    m_HooksDestroyed = m_HooksDestroyed + 1
End Sub

Public Function GetHookCount(Optional ByRef hooksCreated As Long, Optional ByRef hooksDestroyed As Long) As Long
    hooksCreated = m_HooksCreated
    hooksDestroyed = m_HooksDestroyed
    GetHookCount = (hooksCreated - hooksDestroyed)
End Function

'Edit boxes can all share the same background brush (as they are all themed identically).  Call this function instead
' of creating your own brush for every text box instance.
Public Function GetSharedGDIBrush(ByVal requestedColor As Long) As Long
    
    'First things first: if this is the first brush requested by the system, initialize the shared brush list
    If (m_numOfSharedBrushes = 0) Then ReDim m_SharedBrushes(0 To INIT_SIZE_OF_BRUSH_CACHE - 1) As SharedGDIBrush
    
    'Next, look for the requested color in our current cache.  If it exists, we don't want to recreate it.
    Dim brushExists As Boolean, brushIndex As Long, i As Long
    brushExists = False
    
    If (m_numOfSharedBrushes > 0) Then
            
        For i = 0 To m_numOfSharedBrushes - 1
            If (m_SharedBrushes(i).brushColor = requestedColor) And (m_SharedBrushes(i).brushHandle <> 0) Then
            
                'As a failsafe for black brushes, make sure the owner count is valid too
                If m_SharedBrushes(i).numOfOwners > 0 Then
                    brushExists = True
                    brushIndex = i
                    Exit For
                End If
            
            End If
        Next i
            
    End If
    
    'If we found the brush in our cache, increment the owner count, and return the handle immediately
    If brushExists Then
        m_SharedBrushes(brushIndex).numOfOwners = m_SharedBrushes(brushIndex).numOfOwners + 1
        GetSharedGDIBrush = m_SharedBrushes(brushIndex).brushHandle
    
    'If the brush doesn't exist, create it anew
    Else
    
        'If the cache is too small, resize it
        If (m_numOfSharedBrushes > UBound(m_SharedBrushes)) Then ReDim Preserve m_SharedBrushes(0 To m_numOfSharedBrushes * 2 - 1) As SharedGDIBrush
        
        'Update the cache entry with new stats (including the created brush)
        m_SharedBrushes(m_numOfSharedBrushes).brushColor = requestedColor
        m_SharedBrushes(m_numOfSharedBrushes).numOfOwners = 1
        m_SharedBrushes(m_numOfSharedBrushes).brushHandle = CreateSolidBrush(requestedColor)
        
        'Return the newly created brush handle, increment the brush count, and exit
        GetSharedGDIBrush = m_SharedBrushes(m_numOfSharedBrushes).brushHandle
        m_numOfSharedBrushes = m_numOfSharedBrushes + 1
        
    End If
    
End Function

'Edit boxes can all share the same background brush (as they are all themed identically).  Call this function after
' an edit box is unloaded, so we can free the shared brush accordingly.  (If it's more convenient, you can also use
' the ReleaseSharedGDIBrushByHandle version of this function, below.)
Public Sub ReleaseSharedGDIBrushByColor(ByVal requestedColor As Long)

    'If the cache is empty, ignore this request
    If (m_numOfSharedBrushes = 0) Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIBrush() received a release request, but no shared brushes exist."
        Exit Sub
    
    'If the cache is non-empty, find the matching brush and decrement its count.
    Else
    
        Dim i As Long
        For i = 0 To m_numOfSharedBrushes - 1
            
            If (m_SharedBrushes(i).brushColor = requestedColor) Then
                m_SharedBrushes(i).numOfOwners = m_SharedBrushes(i).numOfOwners - 1
                
                'Brushes with a count of 0 are immediately killed.
                If (m_SharedBrushes(i).numOfOwners = 0) Then
                    DeleteObject m_SharedBrushes(i).brushHandle
                    m_SharedBrushes(i).brushHandle = 0
                    m_SharedBrushes(i).brushColor = 0
                End If
                
                Exit For
            End If
            
        Next i
        
    End If

End Sub

Public Sub ReleaseSharedGDIBrushByHandle(ByVal requestedHandle As Long)

    If (m_numOfSharedBrushes = 0) Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIBrushByHandle() received a release request, but no shared brushes exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedBrushes - 1
            If (m_SharedBrushes(i).brushHandle = requestedHandle) Then
                m_SharedBrushes(i).numOfOwners = m_SharedBrushes(i).numOfOwners - 1
                If (m_SharedBrushes(i).numOfOwners = 0) Then
                    DeleteObject m_SharedBrushes(i).brushHandle
                    m_SharedBrushes(i).brushHandle = 0
                    m_SharedBrushes(i).brushColor = 0
                End If
                Exit For
            End If
        Next i
    End If

End Sub

'Edit boxes can all share the same rendering font (as they are all themed identically).  Call this function instead
' of creating your own hFont for every text box instance.
Public Function GetSharedGDIFont(ByVal requestedSize As Single) As Long
    
    'First things first: if this is the first font requested by the system, initialize the shared font list
    If (m_numOfSharedFonts = 0) Then
        ReDim m_SharedFonts(0 To INIT_SIZE_OF_FONT_CACHE - 1) As SharedGDIFont
    End If
    
    'Next, look for the requested size in our current cache.  If it exists, we don't want to recreate it.
    Dim fontExists As Boolean, fontIndex As Long, i As Long
    fontExists = False
    
    If (m_numOfSharedFonts > 0) Then
            
        For i = 0 To m_numOfSharedFonts - 1
            If (m_SharedFonts(i).FontSize = requestedSize) Then
            
                'As a failsafe, make sure the owner count is valid too
                If (m_SharedFonts(i).numOfOwners > 0) And (m_SharedFonts(i).fontHandle <> 0) Then
                    fontExists = True
                    fontIndex = i
                    Exit For
                End If
            
            End If
        Next i
            
    End If
    
    'If we found the right font in our cache, increment the owner count, and return the handle immediately
    If fontExists Then
        m_SharedFonts(fontIndex).numOfOwners = m_SharedFonts(fontIndex).numOfOwners + 1
        GetSharedGDIFont = m_SharedFonts(fontIndex).fontHandle
    
    'If the font doesn't exist, create it anew
    Else
    
        'If the cache is too small, resize it
        If (m_numOfSharedFonts > UBound(m_SharedFonts)) Then ReDim Preserve m_SharedFonts(0 To m_numOfSharedFonts * 2 - 1) As SharedGDIFont
        
        'Font creation is cumbersome, but PD provides some helper functions to simplify it
        Dim tmpLogFont As LOGFONTW
        Fonts.FillLogFontW_Basic tmpLogFont, g_InterfaceFont, False, False, False, False
        Fonts.FillLogFontW_Size tmpLogFont, requestedSize, pdfu_Point
        Fonts.FillLogFontW_Quality tmpLogFont, TextRenderingHintClearTypeGridFit
        
        'Update the cache entry with new stats (including the created font)
        m_SharedFonts(m_numOfSharedFonts).FontSize = requestedSize
        m_SharedFonts(m_numOfSharedFonts).numOfOwners = 1
        If (Not Fonts.CreateGDIFont(tmpLogFont, m_SharedFonts(m_numOfSharedFonts).fontHandle)) Then
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  UserControl_Support.GetSharedGDIFont() failed to create a new UI font handle."
            #End If
        End If
        
        'Return the newly created font handle, increment the font count, and exit
        GetSharedGDIFont = m_SharedFonts(m_numOfSharedFonts).fontHandle
        m_numOfSharedFonts = m_numOfSharedFonts + 1
    
    End If
    
End Function

Public Sub ReleaseSharedGDIFontByHandle(ByVal requestedHandle As Long)

    If (m_numOfSharedFonts = 0) Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIFontByHandle() received a release request, but no shared fonts exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedFonts - 1
            If (m_SharedFonts(i).fontHandle = requestedHandle) Then
                m_SharedFonts(i).numOfOwners = m_SharedFonts(i).numOfOwners - 1
                If (m_SharedFonts(i).numOfOwners = 0) Then
                    Fonts.DeleteGDIFont m_SharedFonts(i).fontHandle
                    m_SharedFonts(i).fontHandle = 0
                    m_SharedFonts(i).FontSize = 0
                End If
                Exit For
            End If
        Next i
    End If

End Sub

'Return a unique, non-zero control ID.  Limited to the size of a VB Long (32-bytes), so don't call more than ~4 billion times.
Public Function GetUniqueControlID() As Long
    
    If (m_UniqueIDTracker = LONG_MAX) Then
        m_UniqueIDTracker = -1 * LONG_MAX
    Else
        m_UniqueIDTracker = m_UniqueIDTracker + 1
    End If
    
    GetUniqueControlID = m_UniqueIDTracker
    
End Function

'Whenever a ucSupport instance is registered by a custom PD usercontrol, this function is called, and our running UC count
' is incremented.
Public Sub IncrementPDControlCount()
    m_PDControlCount = m_PDControlCount + 1
End Sub

Public Sub DecrementPDControlCount()
    m_PDControlCount = m_PDControlCount - 1
End Sub

Public Function GetPDControlCount() As Long
    GetPDControlCount = m_PDControlCount
End Function

'Whenever a dropdown raises its list box, call this function to set some program-wide flags.  Subsequent focus events
' will also notify us, and we will kill the list box as necessary.
Public Sub NotifyDropDownChangeState(ByVal dropDownHWnd As Long, ByVal dropDownListHWnd As Long, ByVal newState As Boolean)
    
    If newState Then
        m_CurrentDropDownHWnd = dropDownHWnd
        m_CurrentDropDownListHWnd = dropDownListHWnd
    Else
        m_CurrentDropDownHWnd = 0
        m_CurrentDropDownListHWnd = 0
    End If

End Sub

'Whenever a PD control loses or receives focus, we receive a corresponding notification
Public Sub PDControlReceivedFocus(ByVal controlHWnd As Long)
    
    'If a dropdown window is still active, hide it now
    If (m_CurrentDropDownHWnd <> 0) Or (m_CurrentDropDownListHWnd <> 0) Then
    
        If (m_CurrentDropDownHWnd <> controlHWnd) And (m_CurrentDropDownListHWnd <> controlHWnd) Then
            SetParent m_CurrentDropDownListHWnd, m_CurrentDropDownHWnd
            g_WindowManager.SetVisibilityByHWnd m_CurrentDropDownListHWnd, False
            m_CurrentDropDownHWnd = 0
            m_CurrentDropDownListHWnd = 0
        End If
    
    End If

End Sub

Public Sub PDControlLostFocus(ByVal controlHWnd As Long)
    
    'If this control raised a tooltip (and said tooltip is still active), unload it now
    If (controlHWnd = m_TTOwner) Then HideUCTooltip
    
End Sub

'When an object requests a tooltip, they need to pass a number of additional parameters (like the window rect, which is used to
' ideally position the tooltip).  Logic similar to pdDropDown is used to display the tooltip.
Public Sub ShowUCTooltip(ByVal ownerHwnd As Long, ByRef srcControlRect As RECTL, ByVal mouseX As Single, ByVal mouseY As Single, ByRef ttCaption As String, ByRef ttTitle As String)
    
    On Error GoTo UnexpectedTTTrouble
    
    If (Not MainModule.IsProgramRunning()) Then Exit Sub
    
    'If a tooltip is currently active, suspend the release timer (because we're just going to "snap" the current
    ' tooltip window into place, rather than waiting for an animation).
    If ((Not m_TimerEventSink Is Nothing) And m_TTActive) Then
        m_TimerEventSink.StopTTTimer
    End If
    
    m_TTOwner = ownerHwnd
    
    'We now want to figure out the idealized coordinates for the tooltip.  The goal is to position the tooltip as
    ' close to the mouse pointer as possible, while also positioning it outside the control rectangle (so that we
    ' don't obscure the control's contents - a constant annoyance with normal tooltips).
    Dim ttRect As RECTF
    
    'Start by figuring out which edge is closest to the current mouse position.  The passed mouse x/y ratios make this simple.
    ' (Each mouse value is a value [0, 1] instead of a hard-coded coordinate.)
    Dim mouseScreenPos As POINTAPI
    mouseScreenPos.x = mouseX
    mouseScreenPos.y = mouseY
    ClientToScreen ownerHwnd, mouseScreenPos
    
    Dim ttDistance(0 To 3) As Single
    ttDistance(TTS_Right) = srcControlRect.Right - mouseScreenPos.x
    ttDistance(TTS_Left) = mouseScreenPos.x - srcControlRect.Left
    ttDistance(TTS_Bottom) = srcControlRect.Bottom - mouseScreenPos.y
    ttDistance(TTS_Top) = mouseScreenPos.y - srcControlRect.Top
    
    'Find the side of the control rect with the smallest distance from the mouse.  This is where we'll first attempt
    ' to place the tooltip.  (If it doesn't fit on that side, we'll flip it to the other side.)
    Dim ttMinPosition As TT_SIDE, minDistance As Single
    ttMinPosition = 0: minDistance = ttDistance(0)
    Dim i As Long
    For i = 1 To 3
        If (ttDistance(i) < minDistance) Then
            ttMinPosition = i
            minDistance = ttDistance(i)
        End If
    Next i
    
    'Next, figure out how big the tooltip needs to be.  This is fairly cumbersome, as we need to fit both the caption
    ' and the title (if any), with text contents auto-wrapped if the tooltip is too long.
    Dim ttCaptionWidth As Long, ttCaptionHeight As Long
    Dim ttFont As pdFont
    Set ttFont = Fonts.GetMatchingUIFont(10)
    
    Dim availableWidth As Long
    availableWidth = Interface.FixDPI(PD_TT_MAX_WIDTH)
    
    'Tooltips can include linebreaks, so all size detection needs to be multiline-aware.
    Dim dtRect As RECTL
    ttFont.GetBoundaryRectOfMultilineString ttCaption, availableWidth, dtRect
    ttCaptionWidth = dtRect.Right - dtRect.Left
    If (ttCaptionWidth > availableWidth) Then
        ttCaptionWidth = availableWidth
        ttCaptionHeight = ttFont.GetHeightOfWordwrapString(ttCaption, ttCaptionWidth)
    Else
        ttCaptionHeight = dtRect.Bottom - dtRect.Top
    End If
    
    'We now have a precise width/height measurement for the tooltip caption.  Repeat the steps for the tooltip title, if any.
    Dim ttTitleWidth As Long, ttTitleHeight As Long
    If (Len(ttTitle) > 0) Then
        Set ttFont = Fonts.GetMatchingUIFont(10, True)
        ttTitleWidth = ttFont.GetWidthOfString(ttTitle)
        
        If (ttTitleWidth > availableWidth) Then
            ttTitleWidth = availableWidth
            ttTitleHeight = ttFont.GetHeightOfWordwrapString(ttTitle, ttTitleWidth)
        Else
            ttTitleHeight = ttFont.GetHeightOfString(ttTitle)
        End If
        
    Else
        ttTitleWidth = 0
        ttTitleHeight = 0
    End If
    
    'All font calculations use PD's shared UI font cache, so we need to free our fonts when we're done with them.
    Set ttFont = Nothing
    
    'With all string sizes calculated, we can now calculate a total size for the tooltip, including padding, borders,
    ' and spacing between the caption and title (if any).
    Dim internalPadding As Long
    internalPadding = Interface.FixDPI(PD_TT_INTERNAL_PADDING)
    
    With ttRect
        .Width = internalPadding * 2 + PDMath.Max2Int(ttCaptionWidth, ttTitleWidth)
        If (ttTitleHeight > 0) Then
            .Height = internalPadding * 2 + ttCaptionHeight + ttTitleHeight + Interface.FixDPI(PD_TT_TITLE_PADDING)
        Else
            .Height = internalPadding * 2 + ttCaptionHeight
        End If
    End With
    
    'With our tooltip size correctly calculated, we now need to determine tooltip position.
    ' (For this step, we use the "closest side to mouse hover position" that we calculated earlier.)
    
    'Also, at present, external padding is not DPI-aware, by design; the only goal of external padding is to provide minimal
    ' visual separation between the tooltip and its parent control.
    Select Case ttMinPosition
        Case TTS_Top
            ttRect.Top = srcControlRect.Top - (PD_TT_EXTERNAL_PADDING + ttRect.Height)
        Case TTS_Bottom
            ttRect.Top = srcControlRect.Bottom + PD_TT_EXTERNAL_PADDING
        Case TTS_Right
            ttRect.Left = srcControlRect.Right + PD_TT_EXTERNAL_PADDING
        Case TTS_Left
            ttRect.Left = srcControlRect.Left - (PD_TT_EXTERNAL_PADDING + ttRect.Width)
    End Select
    
    'Next, make sure that the tooltip lies on-screen.  If it lies off-screen, we're going to flip it to the other side of
    ' its parent control.  (For this to work, we need to know the current screen dimensions; pdDisplays is used for this.)
    Dim hMonitor As Long
    hMonitor = g_Displays.GetHMonitorFromRectL(srcControlRect)
    
    Dim monitorRect As RECTL
    g_Displays.GetDisplayByHandle(hMonitor).GetWorkingRect monitorRect
    
    Select Case ttMinPosition
        Case TTS_Top
            If (ttRect.Top < monitorRect.Top) Then ttRect.Top = srcControlRect.Bottom + PD_TT_EXTERNAL_PADDING
        Case TTS_Bottom
            If ((ttRect.Top + ttRect.Height) > monitorRect.Bottom) Then ttRect.Top = srcControlRect.Top - (PD_TT_EXTERNAL_PADDING + ttRect.Height)
        Case TTS_Right
            If ((ttRect.Left + ttRect.Width) > monitorRect.Right) Then ttRect.Left = srcControlRect.Left - (PD_TT_EXTERNAL_PADDING + ttRect.Width)
        Case TTS_Left
            If (ttRect.Left < monitorRect.Left) Then ttRect.Left = srcControlRect.Right + PD_TT_EXTERNAL_PADDING
    End Select
    
    'The tooltip's primary dimension has been properly set.  Next, calculate its secondary dimension.
    ' (Ideally, the secondary dimension is centered relative to the mouse hover position.  If this results in an
    '  off-screen tooltip, we automatically nudge it on-screen.)
    Select Case ttMinPosition
        Case TTS_Top, TTS_Bottom
            ttRect.Left = mouseScreenPos.x - ttRect.Width \ 2
            If (ttRect.Left < monitorRect.Left) Then ttRect.Left = monitorRect.Left
            If (ttRect.Left + ttRect.Width > monitorRect.Right) Then ttRect.Left = monitorRect.Right - ttRect.Width
            
        Case TTS_Right, TTS_Left
            ttRect.Top = mouseScreenPos.y - ttRect.Height \ 2
            If (ttRect.Top < monitorRect.Top) Then ttRect.Top = monitorRect.Top
            If (ttRect.Top + ttRect.Height > monitorRect.Bottom) Then ttRect.Top = monitorRect.Bottom - ttRect.Height
            
    End Select
    
    'We have now calculated the tooltip position.  Time to display it!
    
    'The first time we raise the tooltip form, we want to cache its current window longs.  (We must restore these before
    ' unloading the form, or VB's built-in teardown functions will crash and burn.)
    If (m_TTHwnd = 0) Then
        
        Load tool_Tooltip
        m_TTHwnd = tool_Tooltip.hWnd
    
        If (Not m_TTWindowStyleHasBeenSet) Then
            m_TTWindowStyleHasBeenSet = True
            m_OriginalTTWindowBits = g_WindowManager.GetWindowLongWrapper(m_TTHwnd)
            m_OriginalTTWindowBitsEx = g_WindowManager.GetWindowLongWrapper(m_TTHwnd, True)
        End If
    
        'Overwrite VB's default window bits to ensure that the tooltip form behaves like a tooltip window.  Of particular
        ' importance is the WS_EX_NOACTIVATE option, to ensure that the tooltip does *not* receive focus.
        Const WS_POPUP As Long = &H80000000
        g_WindowManager.SetWindowLongWrapper m_TTHwnd, WS_POPUP, False, False, True
        g_WindowManager.SetWindowLongWrapper m_TTHwnd, WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW, False, True, True
            
        'Notify the window of the frame changes
        SetWindowPos m_TTHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
        
    End If
    
    'Move the tooltip window into position *but do not display it* just yet.
    With ttRect
        SetWindowPos m_TTHwnd, 0&, .Left, .Top, .Width, .Height, SWP_NOACTIVATE
    End With
    
    'Cache the tooltip's display rect.  When the tooltip disappears, we will manually invalidate windows
    ' beneath it (only on certain OS + theme combinations; Aero handles this correctly).
    With m_TTRectCopy
        .Left = ttRect.Left
        .Top = ttRect.Top
        .Right = ttRect.Left + ttRect.Width
        .Bottom = ttRect.Top + ttRect.Height
    End With
    
    'As the last step before showing the tooltip, we need to notify the tooltip form of the tooltip caption and/or title.
    ' It will cache these values and prepare internal rendering structs to match.
    tool_Tooltip.NotifyTooltipSettings ttCaption, ttTitle, PD_TT_INTERNAL_PADDING, PD_TT_TITLE_PADDING
    
    'We are finally ready to display the tooltip; we also notify the window of its changed window style bits
    AnimateWindow m_TTHwnd, 150&, AW_BLEND
    
    'If you don't want to animate the window, SetWindowPos can be used in place of AnimateWindow
    'With ttRect
    '    SetWindowPos m_TTHwnd, 0&, .Left, .Top, .Width, .Height, SWP_SHOWWINDOW Or SWP_NOACTIVATE
    'End With
    
    m_TTActive = True
    
    Exit Sub
    
UnexpectedTTTrouble:
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  UserControl_Support.ShowUCTooltip failed because of Err # " & Err.Number & ", " & Err.Description
    #End If
    
End Sub

Public Sub HideUCTooltip(Optional ByVal hideImmediately As Boolean = False, Optional ByVal useAnimation As Boolean = True)
    
    If m_TTActive Then
        
        If hideImmediately Then
            HideTTImmediately useAnimation
        Else
        
            'Note the current time, then start the tooltip hide countdown
            If (m_TimerEventSink Is Nothing) Then
                Set m_TimerEventSink = New pdUCEventSink
            Else
                m_TimerEventSink.StopTTTimer
            End If
            
            m_InitialTTTimerTime = Timer
            m_TimerEventSink.StartTTTimer 100
            
        End If
        
    End If
        
End Sub

Public Sub TTTimerFired()
    
    'If enough time has passed, hide the tooltip and release the countdown timer
    If (Abs(Timer - m_InitialTTTimerTime) >= 0.5) Then HideTTImmediately
    
End Sub

Private Sub HideTTImmediately(Optional ByVal useAnimation As Boolean = True)

    If (Not m_TimerEventSink Is Nothing) Then m_TimerEventSink.StopTTTimer
    
    If m_TTActive And (m_TTHwnd <> 0) Then
        
        'Hide (but do not unload!) the tooltip window.  The commented out code is for a non-animated approach.
        If useAnimation Then
            AnimateWindow m_TTHwnd, 150&, AW_HIDE Or AW_BLEND
        Else
            g_WindowManager.SetVisibilityByHWnd m_TTHwnd, False
        End If
        
        'If Aero theming is not active, hiding the tooltip may cause windows beneath the current one to render incorrectly.
        If (OS.IsVistaOrLater And (Not g_WindowManager.IsDWMCompositionEnabled)) Then
            InvalidateRect 0&, VarPtr(m_TTRectCopy), 0&
        End If
        
    End If
    
    m_TTOwner = 0
    m_TTActive = False
        
End Sub

Public Function IsTooltipActive(ByVal ownerHwnd As Long) As Boolean
    IsTooltipActive = CBool(m_TTOwner = ownerHwnd)
End Function

'Do not call this function until the program is going down.  VB is very unhappy about changing window longs on the fly,
' so we only do it once, when the tooltip form is first raised.  After that, we keep the form in memory as-is, and do not
' touch its window longs again until the window is released.
Public Sub FinalTooltipUnload()
    
    'If a release timer is already active, release it immediately
    If (Not m_TimerEventSink Is Nothing) Then
        m_TimerEventSink.StopTTTimer
        Set m_TimerEventSink = Nothing
    End If
    
    If (m_TTHwnd <> 0) Then
    
        'Before doing anything else, ensure the window is invisible
        g_WindowManager.SetVisibilityByHWnd m_TTHwnd, False
        
        'Restore the original VB window bits; this ensures that teardown happens correctly
        If (m_OriginalTTWindowBits <> 0) Then g_WindowManager.SetWindowLongWrapper m_TTHwnd, m_OriginalTTWindowBits, , , True
        If (m_OriginalTTWindowBitsEx <> 0) Then g_WindowManager.SetWindowLongWrapper m_TTHwnd, m_OriginalTTWindowBits, , True, True
        
        'Windows caches window longs; ensure that our changes are applied immediately
        SetWindowPos m_TTHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_HIDEWINDOW Or SWP_NOSENDCHANGING
        
        'With all original settings restored, we can safely unload the tooltip window
        Unload tool_Tooltip
        Set tool_Tooltip = Nothing
        
        m_TTHwnd = 0
        
    End If
    
End Sub
