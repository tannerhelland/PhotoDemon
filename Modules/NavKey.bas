Attribute VB_Name = "NavKey"
'***************************************************************************
'Navigation Key Handler (including automated tab order handling)
'Copyright 2017-2017 by Tanner Helland
'Created: 18/August/17
'Last updated: 22/August/17
'Last update: continuing work on initial build
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const INIT_NUM_OF_FORMS As Long = 8
Private m_Forms() As pdObjectList
Private m_NumOfForms As Long, m_LastForm As Long

'Before loading individual controls, notify this module of the parent form preceding the loop.  (This improves
' performance because we don't have to look-up the form in our table for subsequent calls.)
Public Sub NotifyFormLoading(ByRef parentForm As Form)

    'At present, PD guarantees that forms will not be double-loaded - e.g. only one instance is allowed at a time.
    ' As such, we don't have to search our table for existing entries.
    If (Not parentForm Is Nothing) Then
        
        'Make sure we have room for this form
        If (m_NumOfForms = 0) Then
            ReDim m_Forms(0 To INIT_NUM_OF_FORMS - 1) As pdObjectList
        Else
            If (m_NumOfForms > UBound(m_Forms)) Then ReDim Preserve m_Forms(0 To m_NumOfForms * 2 - 1) As pdObjectList
        End If
        
        Set m_Forms(m_NumOfForms) = New pdObjectList
        m_Forms(m_NumOfForms).SetParentHWnd parentForm.hWnd
        
        m_LastForm = m_NumOfForms
        m_NumOfForms = m_NumOfForms + 1
        
    End If

End Sub

Public Sub NotifyFormUnloading(ByRef parentForm As Form)

    'Find the matching form in our object list
    If (m_NumOfForms > 0) Then
        
        Dim targetHWnd As Long
        targetHWnd = parentForm.hWnd
        
        Dim i As Long, indexOfForm As Long
        For i = 0 To m_NumOfForms - 1
            If (Not m_Forms(i) Is Nothing) Then
                If (m_Forms(i).GetParentHWnd = targetHWnd) Then
                    
                    
                    'DEBUG ONLY!
                    m_Forms(i).PrintDebugList
                    
                    
                    Set m_Forms(i) = Nothing
                    indexOfForm = i
                    Exit For
                    
                End If
            End If
        Next i
        
        'If we removed this from the middle of the list, shift subsequent entries down
        If (indexOfForm < m_NumOfForms - 1) Then
        
            For i = indexOfForm To m_NumOfForms - 2
                Set m_Forms(i) = m_Forms(i + 1)
            Next i
            
            m_NumOfForms = m_NumOfForms - 1
        
        End If
        
    End If
    

End Sub

'After calling NotifyFormLoading(), above, you can proceed to notify us of all child controls.
Public Sub NotifyControlLoad(ByRef childObject As Control)
    m_Forms(m_LastForm).NotifyChildControl childObject
End Sub

'Most dialogs in PD are loaded in a strictly stack-like order.  Forms loaded this way are handled automatically
' (with the assumption that a just-loaded form is also the active form - in 99% of cases, this is a valid assumption.)
' The exception to this rule are toolbars and their associated panels, which are simultaneously active and frequently
' switched-between.  To improve key-matching, you can manually notify this manager of form activation...
' (IS THIS EVEN NECESSARY??)
Public Sub NotifyFormActivation(ByRef srcForm As Form)

End Sub

'When a PD control receives a "navigation" keypress (Enter, Esc, Tab), relay it to this function to activate
' automatic handling.  (For example, Enter will trigger a command bar "OK" press, if a command bar is present
' on the same dialog as the child object.)
Public Sub NotifyNavKeypress(ByRef childObject As Control, ByVal navKey As PD_NavigationKey)

    'First, search the LastForm object for a hit.  (In most cases, that form will be the currently active form,
    ' and it shortcuts the search process to go there first.)
    If (m_LastForm <> 0) Then
    
    End If

End Sub

