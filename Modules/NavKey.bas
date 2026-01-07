Attribute VB_Name = "NavKey"
'***************************************************************************
'Navigation Key Handler (including automated tab order handling)
'Copyright 2017-2026 by Tanner Helland
'Created: 18/August/17
'Last updated: 14/November/21
'Last update: fix nav key tracking across toolboxes (which are no longer unloaded after switching away
'             from them, but are kept in-memory for faster subsequent access)
'
'In a project as complex as PD, tab order is difficult to keep straight.  VB orders controls in the order
' they're added, and there's no easy way to modify this short of manually setting TabOrder across all forms.
' Worse still, many PD usercontrols are actually several controls condensed into one, so they need to manage
' their own internal tab order.
'
'To try and remedy this, PD now uses a homebrew tab order manager.  When a form is loaded, it notifies this
' module of the names and hWnds of all child controls.  This module manages that list internally, and when
' tab commands are raised, this module can be queried to figure out where to send focus.
'
'Similarly, this form automatically orders controls in L-R, T-B order, and because position is calculated at
' run-time, we never have to worry about order being incorrect!
'
'Finally, things like command bar "OK" and "Cancel" buttons are automatically flagged, so we can support
' "Default" and "Cancel" commands on each dialog.  Individual dialogs don't have to manage any of this.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Remember: when passing messages to PD controls, do not call PostMessage directly, as it sends
' messages to the thread's message queue.  Instead, asynchronously relay messages to target windows
' via SendNotifyMessage.
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const INIT_NUM_OF_FORMS As Long = 8
Private m_Forms() As pdObjectList
Private m_NumOfForms As Long, m_LastForm As Long

'After calling NotifyFormLoading(), above, you can proceed to notify us of all child controls.
Public Sub NotifyControlLoad(ByRef childObject As Object, Optional ByVal hostFormhWnd As Long = 0, Optional ByVal canReceiveFocus As Boolean = True)
    
    'If no parent window handle is specified, assume the last form
    If (hostFormhWnd = 0) Then
        If (Not m_Forms(m_LastForm) Is Nothing) Then m_Forms(m_LastForm).NotifyChildControl childObject, canReceiveFocus
    
    'The caller specified a parent window handle.  Find a matching object before continuing.
    Else
        
        'Failsafe checks follow
        If (m_NumOfForms > 0) And (m_LastForm < UBound(m_Forms)) Then
            If (Not m_Forms(m_LastForm) Is Nothing) Then
                
                If (m_Forms(m_LastForm).GetParentHWnd = hostFormhWnd) Then
                    m_Forms(m_LastForm).NotifyChildControl childObject, canReceiveFocus
                Else
                
                    Dim i As Long
                    For i = 0 To m_NumOfForms - 1
                        If (m_Forms(i).GetParentHWnd = hostFormhWnd) Then
                            m_Forms(i).NotifyChildControl childObject, canReceiveFocus
                            Exit For
                        End If
                    Next i
                
                End If
            
            '/failsafe check for m_Forms(m_LastForm) Is Nothing
            End If
        '/failsafe check for form index exists in form array
        End If
    
    End If
    
End Sub

'Before loading individual controls, notify this module of the parent form preceding the loop.  (This improves
' performance because we don't have to look-up the form in our table for subsequent calls.)
Public Sub NotifyFormLoading(ByRef parentForm As Form, ByVal handleAutoResize As Boolean, Optional ByVal hWndCustomAnchor As Long = 0)
    
    'At present, PD guarantees that *most* forms will not be double-loaded - e.g. only one instance
    ' is allowed for effect and adjustment dialogs.
    '
    'Some exceptions to this rule are the main form (which may be re-themed more than once if the
    ' user does something like change the active language at run-time) and various toolbox panels
    ' (which are not unloaded after switching away from them, for performance reasons).
    '
    'As such, we do need to perform a failsafe check for the specified form in our table.  If we've
    ' already loaded a given form (and not unloaded it), we don't need to initialize a new object tracker.
    If (Not parentForm Is Nothing) Then
        
        'Make sure we have room for this form (expanding the collection is harmless, even if we
        ' find a match for this form in the collection)
        If (m_NumOfForms = 0) Then
            ReDim m_Forms(0 To INIT_NUM_OF_FORMS - 1) As pdObjectList
        Else
            If (m_NumOfForms > UBound(m_Forms)) Then ReDim Preserve m_Forms(0 To m_NumOfForms * 2 - 1) As pdObjectList
        End If
        
        'Perform a quick failsafe check for the current form existing in the collection.
        Dim targetAlreadyExists As Boolean
        targetAlreadyExists = False
        
        Dim targetIndex As Long
        If (m_NumOfForms > 0) Then
            
            Dim i As Long
            For i = 0 To m_NumOfForms - 1
                If (Not m_Forms(i) Is Nothing) Then
                    If (m_Forms(i).GetParentHWnd = parentForm.hWnd) Then
                        targetIndex = i
                        targetAlreadyExists = True
                        Exit For
                    End If
                End If
            Next i
            
        Else
            targetIndex = m_NumOfForms
        End If
        
        If (Not targetAlreadyExists) Then
            targetIndex = m_NumOfForms
            Set m_Forms(targetIndex) = New pdObjectList
            m_Forms(targetIndex).SetParentHWnd parentForm.hWnd, handleAutoResize, hWndCustomAnchor, parentForm.Name
        End If
        
        m_LastForm = targetIndex
        If (Not targetAlreadyExists) Then m_NumOfForms = m_NumOfForms + 1
        
    End If

End Sub

Public Sub NotifyFormUnloading(ByRef parentForm As Form)
    
    'Find the matching form in our object list
    If (m_NumOfForms > 0) Then
        
        Dim targetHWnd As Long
        targetHWnd = parentForm.hWnd
        
        Dim i As Long, indexOfForm As Long
        indexOfForm = -1
        
        For i = 0 To m_NumOfForms - 1
            If (Not m_Forms(i) Is Nothing) Then
                If (m_Forms(i).GetParentHWnd = targetHWnd) Then
                    
                    'Want to know what this collection tracked?  Use the helpful "PrintDebugList()" function.
                    'm_Forms(i).PrintDebugList
                    
                    Set m_Forms(i) = Nothing
                    indexOfForm = i
                    Exit For
                    
                End If
            End If
        Next i
        
        'If we removed this from the middle of the list, shift subsequent entries down
        If (indexOfForm >= 0) And (indexOfForm < m_NumOfForms - 1) Then
            m_NumOfForms = m_NumOfForms - 1
            For i = indexOfForm To m_NumOfForms - 1
                Set m_Forms(i) = m_Forms(i + 1)
            Next i
            Set m_Forms(m_NumOfForms) = Nothing
        End If
        
    End If

End Sub

'When a PD control receives a "navigation" keypress (Enter, Esc, Tab), relay it to this function to activate
' automatic handling.  (For example, Enter will trigger a command bar "OK" press, if a command bar is present
' on the same dialog as the child object.)
Public Function NotifyNavKeypress(ByRef childObject As Object, ByVal navKeyCode As PD_NavigationKey, ByVal Shift As ShiftConstants) As Boolean
        
    If (Not PDMain.IsProgramRunning()) Then Exit Function
        
    NotifyNavKeypress = False
    
    Dim childHwnd As Long
    childHwnd = childObject.hWnd
    
    Dim formIndex As Long
    formIndex = GetFormIndex(childHwnd)
    
    'It should be physically impossible to *not* have a form index by now, but better safe than sorry.
    If (formIndex >= 0) Then
        
        Dim targetHWnd As Long
        
        'For Enter and Esc keypresses, we want to see if the target form contains a command bar.  If it does,
        ' we'll directly invoke the appropriate keypress.
        If (navKeyCode = pdnk_Enter) Or (navKeyCode = pdnk_Escape) Or (navKeyCode = pdnk_Space) Then
            
            'See if this form 1) is a raised dialog, and 2) contains a command bar
            If Interface.IsModalDialogActive() Then
            
                If m_Forms(formIndex).DoesTypeOfControlExist(pdct_CommandBar) Then
                
                    'It does!  Grab the hWnd and forward the relevant window message to it
                    targetHWnd = m_Forms(formIndex).GetFirstHWndForType(pdct_CommandBar)
                    SendNotifyMessage targetHWnd, WM_PD_DIALOG_NAVKEY, navKeyCode, 0&
                    NotifyNavKeypress = True
                
                'If a command bar doesn't exist, look for a "mini command bar" instead
                ElseIf m_Forms(formIndex).DoesTypeOfControlExist(pdct_CommandBarMini) Then
                    targetHWnd = m_Forms(formIndex).GetFirstHWndForType(pdct_CommandBarMini)
                    SendNotifyMessage targetHWnd, WM_PD_DIALOG_NAVKEY, navKeyCode, 0&
                    NotifyNavKeypress = True
                    
                'No command bar exists on this form, which is fine - this could be a toolpanel, for example.
                ' As such, there's nothing we need to do.
                End If
            
            'If a modal dialog is *not* active, let the caller handle Enter/Esc presses on their own
            Else
                NotifyNavKeypress = False
            End If
        
        'The only other supported key (at this point) is TAB.  Tab keypresses are handled by the object list;
        ' it's responsible for figuring out which control is next in order.
        ElseIf (navKeyCode = pdnk_Tab) Then
            m_Forms(formIndex).NotifyTabKey childHwnd, ((Shift And vbShiftMask) <> 0)
            NotifyNavKeypress = True
        End If
        
    Else
        Debug.Print "WARNING!  NavKey.NotifyNavKeypress couldn't find this control in its collection.  How is this possible?"
    End If

End Function

Private Function GetFormIndex(ByVal childHwnd As Long) As Long

    GetFormIndex = -1
    
    'First, search the LastForm object for a hit.  (In most cases, that form will be the currently active form,
    ' and it shortcuts the search process to go there first.)
    If (m_LastForm <> 0) Then
        If (Not m_Forms(m_LastForm) Is Nothing) Then
            If m_Forms(m_LastForm).DoesHWndExist(childHwnd) Then GetFormIndex = m_LastForm
        End If
    End If
    
    'If we didn't find the hWnd in our last-activated form, try other forms until we get a hit
    If (GetFormIndex < 0) Then
        
        Dim i As Long
        For i = 0 To m_NumOfForms - 1
        
            'Normally, we would never expect to encounter a null entry here, but as a failsafe against forms
            ' unloading incorrectly (especially if we ever implement plugins), check for null objects
            If (Not m_Forms(i) Is Nothing) Then
            
                'While we're here, update m_LastForm to match - it may improve performance on subsequent matches
                If m_Forms(i).DoesHWndExist(childHwnd) Then
                    GetFormIndex = i
                    m_LastForm = GetFormIndex
                    Exit For
                End If
                
            End If
            
        Next i
        
    End If
    
End Function

'Given a child hWnd, return the name of its container window.  (PD uses this to generate runtime object names
' for matching against localization text.)
Public Function GetParentName(ByVal childHwnd As Long) As String
    Dim idxParent As Long
    idxParent = GetFormIndex(childHwnd)
    If (idxParent >= 0) Then GetParentName = m_Forms(idxParent).GetParentName()
End Function
