VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWindowSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon DPI-Aware Window Resize class
'Copyright 2015-2026 by Tanner Helland
'Created: 23/September/15 (but built from many parts existing earlier)
'Last updated: 26/August/20
'Last update: add optional support for handling min/max messages; this will be important when I finally
'             tackle resizeable effect dialogs.
'
'Because native VB6 is only partially DPI-aware, each unique user control has to do a bunch
' of specialized work to support resize and move events under non-standard DPI values.
'
'To reduce the amount of redundant code copied between UCs, this parent class is used.
' It provides replacement move/size functions that rely on WAPI for correct behavior.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class can optionally raise events for move/size actions, but you must enable subclassing in the main Attach() function.
Public Event WindowMove(ByVal newLeft As Long, ByVal newTop As Long)
Public Event WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
Public Event WindowResizeInitial()  'Sent just once, when interactive resizing begins; good for suspending automatic redraws
Public Event WindowResizeFinal(ByVal newWidth As Long, ByVal newHeight As Long) 'Sent just once, after interactive resizing finishes; good for performing intensive functions
Public Event WindowMaxMinRequested(ByRef minWidth As Long, ByRef minHeight As Long, ByRef maxWidth As Long, ByRef maxHeight As Long)

'Various window interaction APIs
Private Enum SWP_FLAGS
    SWP_ASYNCWINDOWPOS = &H4000
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOSENDCHANGING = &H400
    SWP_NOSIZE = &H1
    SWP_HIDEWINDOW = &H80
    SWP_SHOWWINDOW = &H40
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = &H20
    SWP_NOCOPYBITS = &H100
End Enum

Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetParent Lib "user32" (ByVal targetHWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal ptrToPointList As Long, ByVal numPoints As Long) As Long

'Optional window messages to subclass
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5
Private Const WM_GETMINMAXINFO As Long = &H24

'This is the API struct used by WM_GETMINMAXINFO
Private Type POINTL
    x As Long
    y As Long
End Type

Private Type MinMaxInfo
    Reserved As POINTL
    MaxSize As POINTL
    maxPosition As POINTL
    MinTrackSize As POINTL
    MaxTrackSize As POINTL
End Type

'Subclassing is optional; if active, this class will be non-nothing.
Implements ISubclass
Private m_SubclassActive As Boolean

'hWnd attached to this class instance; use it for all API interactions
Private m_hWnd As Long

'Rect that contains the position/dimensions of the tracked hWnd.  It will be automatically
' updated on size/move events, or you can optionally call the UpdateWinRect function to do
' it yourself.  Note that per GetWindowRect functionality, all dimensions are in *screen coords*.
Private m_WindowRect As winRect

'If (optional) max/min tracking is active, this will be set to TRUE, and the associated struct
' will be passed to the user for modification.
Private m_SubclassMaxMin As Boolean, m_tmpMinMax As MinMaxInfo

'You can always query this to see if a resize event is active; FALSE means the resize has completed
Private m_ResizeActive As Boolean

'Prior to doing anything with this class, you need to attach it to an hWnd.
' (For UCs, this will simply be UserControl.hWnd.)
'
'This class can optionally subclass API size/move events, including max/min window size requests,
' and raise corresponding events, allowing us to completely bypass VB's faulty handling.
Friend Sub AttachToHWnd(ByVal srcHWnd As Long, Optional ByVal subclassMoveSizeEvents As Boolean = False, Optional ByVal subclassMaxMinToo As Boolean = False)
    
    If PDMain.IsProgramRunning() Then
    
        EndSubclassing
        
        If (srcHWnd = 0) Then
            PDDebug.LogAction "WARNING!  pdWindowSize.AttachToHWnd received a null hWnd - was this intentional??"
        Else
        
            m_hWnd = srcHWnd
            
            'Activate the subclasser conditionally; some UCs don't need to receive API size/move events
            m_SubclassActive = subclassMoveSizeEvents
            If m_SubclassActive Then VBHacks.StartSubclassing m_hWnd, Me
            m_SubclassMaxMin = subclassMaxMinToo
            
            'Cache an initial set of window dimensions
            UpdateWinRect
            
        End If
        
    End If
    
End Sub

'FYI: all retrieval functions operate purely in the pixel domain
Friend Function GetWidth() As Long
    If (Not m_SubclassActive) Then UpdateWinRect
    GetWidth = m_WindowRect.x2 - m_WindowRect.x1
End Function

Friend Function GetHeight() As Long
    If (Not m_SubclassActive) Then UpdateWinRect
    GetHeight = m_WindowRect.y2 - m_WindowRect.y1
End Function

Friend Function GetLeft() As Long
    If (Not m_SubclassActive) Then UpdateWinRect
    GetLeft = m_WindowRect.x1
End Function

Friend Function GetTop() As Long
    If (Not m_SubclassActive) Then UpdateWinRect
    GetTop = m_WindowRect.y1
End Function

Friend Function IsResizeActive() As Boolean
    IsResizeActive = m_ResizeActive
End Function

'Move or resize a window
Friend Sub SetPosition(ByVal newLeft As Long, ByVal newTop As Long, Optional ByVal notifyWindow As Boolean = True)
    
    'Per VB convention, the new left/top coordinates should be in parent coordinates.  If you want to position a
    ' window using *screen coordinates*, you'll need to write a new function (and/or manually convert the coordinates
    ' to screen space in advance).
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    If (Not notifyWindow) Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos m_hWnd, 0&, newLeft, newTop, 0, 0, swpFlags
    If (Not notifyWindow) Then UpdateWinRect
    
End Sub

Friend Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    If (Not notifyWindow) Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos m_hWnd, 0&, 0&, 0&, newWidth, newHeight, swpFlags
    If (Not notifyWindow) Then UpdateWinRect
    
End Sub

Friend Sub SetSizeAndPosition(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)

    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    If (Not notifyWindow) Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos m_hWnd, 0&, newLeft, newTop, newWidth, newHeight, swpFlags
    If (Not notifyWindow) Then UpdateWinRect
    
End Sub

'Update the module-level copy of the attached hWnd's rect, automatically translated into our parent's coordinate space.
Private Sub UpdateWinRect()
    
    If (m_hWnd <> 0) Then
    
        'Retrieve a copy of the window rect *in screen coordinates*
        GetWindowRect m_hWnd, m_WindowRect
        
        'Find the difference between *the screen coordinate space* and *the parent's coordinate space*
        Dim tmpPoint As PointAPI
        tmpPoint.x = m_WindowRect.x1
        tmpPoint.y = m_WindowRect.y1
        MapWindowPoints 0&, GetParent(m_hWnd), VarPtr(tmpPoint), 1
        
        'Translate the window rect into our parent's coordinate space, which greatly simplifies subsequent property requests
        Dim xDiff As Long, yDiff As Long
        xDiff = tmpPoint.x - m_WindowRect.x1
        yDiff = tmpPoint.y - m_WindowRect.y1
        
        With m_WindowRect
            .x1 = .x1 + xDiff
            .x2 = .x2 + xDiff
            .y1 = .y1 + yDiff
            .y2 = .y2 + yDiff
        End With
        
    End If
    
End Sub

Private Sub Class_Initialize()
    m_hWnd = 0
End Sub

Private Sub Class_Terminate()
    EndSubclassing
End Sub

Private Sub EndSubclassing()
    If (m_hWnd <> 0) Then
        If m_SubclassActive Then VBHacks.StopSubclassing m_hWnd, Me
        m_hWnd = 0
    End If
    m_SubclassActive = False
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    If (uiMsg = WM_ENTERSIZEMOVE) Then
        m_ResizeActive = True
        RaiseEvent WindowResizeInitial
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    ElseIf (uiMsg = WM_EXITSIZEMOVE) Then
        m_ResizeActive = False
        UpdateWinRect
        RaiseEvent WindowResizeFinal(m_WindowRect.x2 - m_WindowRect.x1, m_WindowRect.y2 - m_WindowRect.y1)
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    ElseIf (uiMsg = WM_SIZE) Then
        UpdateWinRect
        RaiseEvent WindowResize(m_WindowRect.x2 - m_WindowRect.x1, m_WindowRect.y2 - m_WindowRect.y1)
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    ElseIf (uiMsg = WM_MOVE) Then
        UpdateWinRect
        RaiseEvent WindowMove(m_WindowRect.x1, m_WindowRect.y1)
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    ElseIf ((uiMsg = WM_GETMINMAXINFO) And m_SubclassMaxMin) Then
        
        'lParam contains a pointer to the window's rect location.  Retrieve it now.
        If (lParam <> 0) Then CopyMemoryStrict VarPtr(m_tmpMinMax), lParam, LenB(m_tmpMinMax)
        RaiseEvent WindowMaxMinRequested(m_tmpMinMax.MinTrackSize.x, m_tmpMinMax.MinTrackSize.y, m_tmpMinMax.MaxTrackSize.x, m_tmpMinMax.MaxTrackSize.y)
        
        'Copy the struct back into place, and return 0 *without* calling DefWindowProc
        CopyMemoryStrict lParam, VarPtr(m_tmpMinMax), LenB(m_tmpMinMax)
        ISubclass_WindowMsg = 0
    
    ElseIf (uiMsg = WM_NCDESTROY) Then
        EndSubclassing
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    Else
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    End If
    
End Function
