VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWindowPainter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Flicker-Free Window Painter class
'Copyright 2014-2026 by Tanner Helland
'Created: 20/October/14 (built from parts existing much earlier)
'Last updated: 06/February/17
'Last update: migrate to safer subclassing technique
'
'Now that PD provides so many of its own user-controls, we have to do a lot of manual
' window painting.  Supporting the full spectrum of Windows versions (particularly XP),
' theme settings (including accessibility themes), window settings (WS_CLIPCHILDREN,
' WS_EX_LAYERED, WS_EX_COMPOSITED all have annoying quirks) is a nightmare, and because
' each new version of Windows introduces new headaches, I've abstracted basic WM_PAINT
' message pumping into this class.  This simplifies the process of handling paint vs
' erase messages in a safe way, and this class can also be used on things like VB
' forms to simplify rendering.
'
'Note that this class does not actually manage a window buffer.  That is left to each
' underlying UI element to handle, as different controls have different needs, and it's
' not possible to use a "one shoe fits all" approach.  (For example, some controls -
' like the slider/text combo - maintain separate buffers for different control elements.)
' This class is simply used to handle paint event subclassing, and to raise a single
' "paint" event with the relevant update rect attached, and all validation/invalidation
' handled automatically.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When this class determines that a paint event is required, it will raise this event.  Note that things like empty
' update rects are automatically detected and suppressed, so the client needs to obey all PaintWindow events -
' they're always relevant!
Public Event PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

'Erase messages will only be passed if specifically requested; anything double-buffered can ignore these, but certain
' controls (like pdContainer) do not maintain a persistent buffer for performance reasons.
Public Event EraseBkgnd()

'Subclassing is used to better optimize the control's painting; this also requires manual validation of the control rect.
Private Const WM_PAINT As Long = &HF
Private Const WM_ERASEBKGND As Long = &H14
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal targetHWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal targetHWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHWnd As Long) As Long

'The window being subclassed
Private m_hWnd As Long

'The window rect to be updated.  Because this rect is passed between multiple functions, we declare it here.
Private m_UpdateRect As RECT

'As part of the painting process, we're gonna be generating a looot of paint messages.  To avoid churn, we'll declare
' a single paint struct up front.
Private m_PaintStruct As PAINTSTRUCT

'BeginPaint returns a DC for the given window; we cache this, in case the client needs to make use of it
Private m_WindowDC As Long

'Most PD controls use double-buffering, so WM_ERASEBKGND messages are not relevant.  However, some container-only
' controls (like pdContainer) don't maintain buffers for performance reasons, so they are the exception, and they
' must be notified of erase messages.
Private m_WantEraseMessages As Boolean

'Subclasser for intercepting window messages
Implements ISubclass

Private Sub Class_Initialize()
    m_hWnd = 0
End Sub

Private Sub Class_Terminate()
    EndSubclassing
End Sub

Private Sub EndSubclassing()
    If (m_hWnd <> 0) Then
        VBHacks.StopSubclassing m_hWnd, Me
        m_hWnd = 0
    End If
End Sub

'This function must be called in the UserControl's Initialize event.  For best results, check for the IDE and do not
' load this class.
Friend Sub StartPainter(ByVal srcHWnd As Long, Optional ByVal wantEraseMessages As Boolean = False)
    
    If PDMain.IsProgramRunning() Then
    
        'Release the existing subclasser, if any
        EndSubclassing
        
        m_hWnd = srcHWnd
    
        'Subclass all necessary messages for proper window painting
        If (m_hWnd <> 0) Then
            VBHacks.StartSubclassing m_hWnd, Me
            m_WantEraseMessages = wantEraseMessages
        End If
    
    End If
    
End Sub

'If a control experiences an event that requires a repaint, e.g. a Click that changes the control's appearance, it can
' manually request a repaint from this function.  Note that - by design - this event will invalidate the entire window,
' as it's assumed that a manually requested paint affects the entire client area of the window.
Friend Sub RequestRepaint(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    
    If (m_hWnd <> 0) Then
        
        'Retrieve the full client rect of the target window
        Dim tmpRect As RECT
        GetClientRect m_hWnd, tmpRect
        
        'If desired, the caller can ask us to immediately raise a paint event.  This is helpful when needing an immediate redraw,
        ' without waiting for WM_PAINT to fire, but you obviously need to be aware of the performance implications involved.
        If raiseImmediateDrawEvent Then
            InvalidateRect m_hWnd, tmpRect, 0&
            UpdateWindow m_hWnd
            
        'Invalidate the client rect, which will automatically trigger the addition of a WM_PAINT message to the window queue.
        ' This way, the window can redraw at the leisure of the system.
        Else
            InvalidateRect m_hWnd, tmpRect, 0&
        End If
        
    End If
    
End Sub

'If the client needs the hDC specified by BeginPaint, they can request it here
Friend Function GetPaintStructDC() As Long
    GetPaintStructDC = m_WindowDC
End Function

'Thin wrapper for the PaintWindow() event.  This is used to keep the actual subclassing function as small as possible.
Friend Sub PrepPaintEvent()

    If (m_hWnd <> 0) Then
    
        'Start painting
        m_WindowDC = BeginPaint(m_hWnd, m_PaintStruct)
        
        'Notify the child that it's time for painting
        With m_UpdateRect
            RaiseEvent PaintWindow(.Left, .Top, .Right - .Left, .Bottom - .Top)
        End With
        
        'End painting (note: BeginPaint automatically validated the window's contents, so we don't need to do any
        ' additional validation here)
        EndPaint m_hWnd, m_PaintStruct
        
    End If
        
End Sub

Private Sub HandleEraseEvent(ByVal wParam As Long)

    'Back up the current DC
    Dim tmpDC As Long
    tmpDC = m_WindowDC
    
    'wParam contains the new window DC
    m_WindowDC = wParam
    
    'Let the control repaint
    RaiseEvent EraseBkgnd
    
    'Reset everything to its original state
    m_WindowDC = tmpDC
    
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    If (uiMsg = WM_PAINT) Then
        
        'Ignore paint requests for empty regions
        If (GetUpdateRect(m_hWnd, m_UpdateRect, 0) <> 0) Then PrepPaintEvent
        
        'Deliberately skip DefSubclassProc, as we've handled all painting
        ISubclass_WindowMsg = 0
        
    ElseIf (uiMsg = WM_ERASEBKGND) Then
        
        'Controls that do not maintain persistent back buffers need to respond to erase messages.
        ' (Note that the HandleEraseEvent function contains very little code, but we don't want local
        '  variables to clog up the stack, so we scope them externally.)
        If m_WantEraseMessages Then HandleEraseEvent wParam
        
        'Deliberately skip DefSubclassProc, as we've handled all painting
        ISubclass_WindowMsg = 1
    
    ElseIf (uiMsg = WM_NCDESTROY) Then
        EndSubclassing
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    Else
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    End If
    
End Function
