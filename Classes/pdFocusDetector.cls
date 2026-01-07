VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFocusDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Window Focus Detection class
'Copyright 2014-2026 by Tanner Helland
'Created: 07/May/15  (but built from many parts existing earlier)
'Last updated: 07/May/15
'Last update: abstract API focus detection bits into this class, so I don't have to keep copying the code into new UCs
'
'VB's internal Got/LostFocus events play very poorly with PD, as we generate a lot of our own API windows for
' various controls.  To receive failsafe Got/LostFocus events, simply add this class to a user control and initialize
' it with the hWnd you want tracked.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class raises correct Got/LostFocus events for any window.
' These two events mirror subclassed WM_SETFOCUS/WM_KILLFOCUS messages.
Public Event GotFocusReliable()
Public Event LostFocusReliable()

'...while these two reflect WM_NCACTIVATE (generally not needed, except for top-level windows)
Public Event AppGotFocusReliable()
Public Event AppLostFocusReliable()

'Various subclassing constants
Private Const WM_ACTIVATE As Long = &H6&
Private Const WM_NCACTIVATE As Long = &H86&
Private Const WA_INACTIVE As Long = &H0&

Private Const WM_SETFOCUS As Long = &H7&
Private Const WM_KILLFOCUS As Long = &H8&

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

'The window being subclassed
Private m_hWnd As Long

'Subclasser for intercepting window messages
Implements ISubclass

'If the window currently has focus, this will be set to TRUE.  We use this to avoid duplicate notifications.
Private m_HasFocus As Boolean

Private Sub Class_Initialize()
    m_hWnd = 0
End Sub

Private Sub Class_Terminate()
    If (m_hWnd <> 0) Then
        VBHacks.StopSubclassing m_hWnd, Me
        m_hWnd = 0
    End If
End Sub

'This function should be called in a UserControl's Initialize event.  For best results, check for the IDE and do not
' load this class.
Friend Sub StartFocusTracking(ByVal srcHWnd As Long)
    
    If PDMain.IsProgramRunning() Then
    
        'Release any existing subclassers
        If (m_hWnd <> 0) Then
            VBHacks.StopSubclassing m_hWnd, Me
            m_hWnd = 0
        End If
        
        'Subclass all necessary messages for proper focus detection
        m_hWnd = srcHWnd
        If (m_hWnd <> 0) Then VBHacks.StartSubclassing m_hWnd, Me
        
        'Query initial focus state
        If (Not g_WindowManager Is Nothing) Then
            
            'Query active focus
            Dim hWndFocus As Long
            hWndFocus = g_WindowManager.GetFocusAPI()
            m_HasFocus = (hWndFocus = srcHWnd)
            
            'If we don't have focus, it's possible one of our child windows does
            If (Not m_HasFocus) Then
                
                'Retrieve the parent (if any) of the window with focus
                Dim nextParent As Long, ctlIsParent As Boolean
                ctlIsParent = False
                nextParent = GetParent(hWndFocus)
                
                'Compare that parent to our currently subclassed window and continue iterating until we reach the
                ' desktop window (GetParent = 0)
                Do While (nextParent <> 0)
                    If (nextParent = srcHWnd) Then
                        m_HasFocus = True
                        Exit Do
                    Else
                        hWndFocus = nextParent
                        nextParent = GetParent(hWndFocus)
                    End If
                Loop
                
            End If
            
            'm_hasFocus now correctly reflects initial focus state of the subclassed control
                
        End If
        
    End If
    
End Sub

'Outside functions can use this to retrieve the current "has focus" state of the tracked control
Friend Function HasFocus() As Boolean
    HasFocus = m_HasFocus
End Function

'If an outside function received a mouse or keyboard event but the control does *not* have focus, they can call this function
' to force an immedate GotFocus event.  (Windows will process hooks prior to focus messages, so some of PD's controls may raise
' input events prior to receiving focus events; this function gives us a way to work around that.)
Friend Sub SetFocusManually()
    If (Not m_HasFocus) And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI m_hWnd
End Sub

Friend Sub ActivateManually()
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.ActivateWindowAPI m_hWnd
End Sub

Private Function GetLoWord(ByVal lParam As Long) As Integer
    If (lParam And &H8000&) Then
        GetLoWord = &H8000 Or (lParam And &H7FFF&)
    Else
        GetLoWord = lParam And &HFFFF&
    End If
End Function

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    'Note: by design, this subclasser still forwards most of its handled messages to their default wndprocs
    If (uiMsg = WM_ACTIVATE) Then
        
        'The low-order portion of wParam contains a notification of whether we are gaining or losing focus.

        'Window is losing focus
        If (GetLoWord(wParam) = WA_INACTIVE) Then
            
            If m_HasFocus Then
                m_HasFocus = False
                UserControls.PDControlLostFocus m_hWnd
                RaiseEvent LostFocusReliable
            End If

        'Window is gaining focus
        Else
            
            If (Not m_HasFocus) Then
                m_HasFocus = True
                UserControls.PDControlReceivedFocus m_hWnd
                RaiseEvent GotFocusReliable
            End If

        End If

        'Allow default focus handlers to still trigger
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)

    'This message doesn't really apply to most uses of this class, but we cover it for completeness' sake
    ElseIf (uiMsg = WM_NCACTIVATE) Then
        
        If (wParam = 1) Then
            If (Not m_HasFocus) Then
                m_HasFocus = True
                UserControls.PDControlReceivedFocus m_hWnd
            End If
            RaiseEvent AppGotFocusReliable
        ElseIf (wParam = 0) Then
            If m_HasFocus Then
                m_HasFocus = False
                UserControls.PDControlLostFocus m_hWnd
            End If
            RaiseEvent AppLostFocusReliable
        End If

        'Per MSDN (https://msdn.microsoft.com/en-us/library/windows/desktop/ms632633)...
        ' "When the wParam parameter is FALSE, an application should return TRUE to indicate that the system
        '  should proceed with the default processing, or it should return FALSE to prevent the change.
        '  When wParam is TRUE, the return value is ignored."

        'We *always* want to proceed with default handling, so we ignore this advice and forcibly allow
        ' DefSubclassProc to request the behavior it wants.
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)

    'Window is losing focus
    ElseIf (uiMsg = WM_KILLFOCUS) Then
        
        If m_HasFocus Then
            m_HasFocus = False
            UserControls.PDControlLostFocus m_hWnd
            RaiseEvent LostFocusReliable
        End If
        
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    'Window is gaining focus
    ElseIf (uiMsg = WM_SETFOCUS) Then
        
        If (Not m_HasFocus) Then
            m_HasFocus = True
            UserControls.PDControlReceivedFocus m_hWnd
            RaiseEvent GotFocusReliable
        End If
        
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    ElseIf (uiMsg = WM_NCDESTROY) Then
        
        If (hWnd <> 0) Then
            VBHacks.StopSubclassing hWnd, Me
            m_hWnd = 0
        End If
        
        'Allow VB's normal teardown to occur
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    Else
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    End If
    
End Function
