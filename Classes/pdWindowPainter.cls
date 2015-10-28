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
'Copyright 2014-2015 by Tanner Helland
'Created: 20/October/14  (but built from many parts existing earlier)
'Last updated: 05/September/15
'Last update: add option for "raise immediate draw event" in .requestRepaint()
'
'Now that PD provides so many of its own user-controls, we have to do a lot of manual window painting.  Supporting the
' full spectrum of Windows versions (particularly XP), theme settings (including accessibility themes), window settings
' (WS_CLIPCHILDREN, WS_EX_LAYERED, WS_EX_COMPOSITED all have annoying quirks) is a nightmare, and because each new
' version of Windows introduces new headaches, I've abstracted all basic UC rendering and subclassing into this class.
' Any updates here will automatically flow to PD's many user controls, hopefully saving me headaches down the road.
'
'Note that this class does not actually manage the window buffer.  That is left to each user control to handle, as
' different controls have different needs, and it's not possible to use a "one shoe fits all" approach.  (For example,
' some controls - like the slider/text combo - maintain separate buffers for different control elements.)  This class is
' simply used to handle all paint event subclassing, and to raise a single "paint" event with the relevant update rect
' attached, and all validation/invalidation handled automatically.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'When this class determines that a paint event is required, it will raise this event.  Note that things like empty
' update rects are automatically detected and suppressed, so the client needs to obey all PaintWindow events -
' they're always relevant!
Public Event PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

'Subclassing is used to better optimize the control's painting; this also requires manual validation of the control rect.
Private Const WM_PAINT As Long = &HF
Private Const WM_ERASEBKGND As Long = &H14
Private Declare Function ValidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal targetHwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal targetHwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long

'The window being subclassed
Private m_hWnd As Long

'The window rect to be updated.  Because this rect is passed between multiple functions, we declare it here.
Private m_UpdateRect As RECT

'As part of the painting process, we're gonna be generating a looot of paint messages.  To avoid churn, we'll declare
' a single paint struct up front.
Private m_PaintStruct As PAINTSTRUCT

'BeginPaint returns a DC for the given window; we cache this, in case the client needs to make use of it
Private m_WindowDC As Long

'Subclasser for intercepting window messages
Private cSubclass As cSelfSubHookCallback

Private Sub Class_Initialize()
    m_hWnd = 0
End Sub

Private Sub Class_Terminate()
    If m_hWnd <> 0 Then
        cSubclass.ssc_UnSubclass m_hWnd
        cSubclass.ssc_Terminate
    End If
End Sub

'This function must be called in the UserControl's Initialize event.  For best results, check for the IDE and do not
' load this class.
Public Sub StartPainter(ByVal srcHwnd As Long)
    
    'Release the existing subclasser, if any
    If Not (cSubclass Is Nothing) And (m_hWnd <> 0) Then
        cSubclass.ssc_UnSubclass m_hWnd
        cSubclass.ssc_Terminate
    End If
    
    m_hWnd = srcHwnd
    
    'Subclass all necessary messages for proper window painting
    If (m_hWnd <> 0) Then
        Set cSubclass = New cSelfSubHookCallback
        cSubclass.ssc_Subclass m_hWnd, , , Me
        cSubclass.ssc_AddMsg m_hWnd, MSG_BEFORE, WM_PAINT, WM_ERASEBKGND
    End If
    
End Sub

'If a control experiences an event that requires a repaint, e.g. a Click that changes the control's appearance, it can
' manually request a repaint from this function.  Note that - by design - this event will invalidate the entire window,
' as it's assumed that a manually requested paint affects the entire client area of the window.
Public Sub RequestRepaint(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    
    If (m_hWnd <> 0) Then
        
        'Retrieve the full client rect of the target window
        Dim tmpRect As RECT
        GetClientRect m_hWnd, tmpRect
        
        'If desired, the caller can ask us to immediately raise a paint event.  This is helpful when needing an immediate redraw,
        ' without waiting for WM_PAINT to fire, but you obviously need to be aware of the performance implications involved.
        If raiseImmediateDrawEvent Then
            RaiseEvent PaintWindow(tmpRect.Left, tmpRect.Top, tmpRect.Right - tmpRect.Left, tmpRect.Bottom - tmpRect.Top)
        
        'Invalidate the client rect, which will automatically trigger the addition of a WM_PAINT message to the window queue.
        ' This way, the window can redraw at the leisure of the system.
        Else
            InvalidateRect m_hWnd, tmpRect, 0
        End If
        
    End If
    
End Sub

'If the client needs the hDC specified by BeginPaint, they can request it here
Public Function GetPaintStructDC() As Long
    GetPaintStructDC = m_WindowDC
End Function

'Thin wrapper for the PaintWindow() event.  This is used to keep the actual subclassing function as small as possible.
Public Sub PrepPaintEvent()

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


    If uMsg = WM_PAINT Then
        
        'Ignore paint requests for empty regions
        If GetUpdateRect(m_hWnd, m_UpdateRect, 0) <> 0 Then
            PrepPaintEvent
        End If
        
        'Mark the message as handled and exit
        bHandled = True
        lReturn = 0
                
    ElseIf uMsg = WM_ERASEBKGND Then
        
        bHandled = True
        lReturn = 1
        
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
