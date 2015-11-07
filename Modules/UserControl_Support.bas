Attribute VB_Name = "UserControl_Support"
'***************************************************************************
'Helper functions for various PhotoDemon UCs
'Copyright 2014-2015 by Tanner Helland
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
    If m_windowMsgCount = 0 Then
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

