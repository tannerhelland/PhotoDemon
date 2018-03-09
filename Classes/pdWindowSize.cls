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
'Copyright 2015-2018 by Tanner Helland
'Created: 23/September/15 (but built from many parts existing earlier)
'Last updated: 06/February/17
'Last update: migrate to safer subclassing technique
'
'Because VB is only partially DPI-aware, each unique user control has to do a ton of specialized work to support
' resize and move events under non-standard DPI values.  To reduce the amount of redundant code copied between UCs,
' this parent class is used.  It provides replacement move/size functions that rely on WAPI for correct behavior.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class can optionally raise events for move/size actions, but you must enable subclassing in the main Attach() function.
Public Event WindowMove(ByVal newLeft As Long, ByVal newTop As Long)
Public Event WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)

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
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GW_OWNER = 4

'Optional window messages to subclass
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5

'Subclassing is optional; if active, this class will be non-nothing.
Implements ISubclass
Private m_SubclassActive As Boolean

'hWnd attached to this class instance; use it for all API interactions
Private m_hWnd As Long

'Rect that contains the position/dimensions of the tracked hWnd.  It will be automatically updated on size/move events,
' or you can optionally call the UpdateWinRect function to do it yourself.  Note that per GetWindowRect functionality,
' all dimensions are in *screen coords*.
Private m_WindowRect As winRect

'Prior to doing anything with this class, you need to attach it to an hWnd.  For UCs, this will simply be UserControl.hWnd.
' This class can optionally subclass API size/move events, and raise corresponding events, allowing us to completely
' bypass VB's faulty handling.
Friend Sub AttachToHWnd(ByVal srcHwnd As Long, Optional ByVal subclassMoveSizeEvents As Boolean = False)
    
    If pdMain.IsProgramRunning() Then
    
        EndSubclassing
        
        If (srcHwnd = 0) Then
            pdDebug.LogAction "WARNING!  pdWindowSize.AttachToHWnd received a null hWnd - was this intentional??"
        Else
        
            m_hWnd = srcHwnd
            
            'Activate the subclasser conditionally; some UCs don't need to receive API size/move events
            m_SubclassActive = subclassMoveSizeEvents
            If m_SubclassActive Then VBHacks.StartSubclassing m_hWnd, Me
            
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
        Dim tmpPoint As POINTAPI
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
    
    If (uiMsg = WM_SIZE) Then
        UpdateWinRect
        RaiseEvent WindowResize(m_WindowRect.x2 - m_WindowRect.x1, m_WindowRect.y2 - m_WindowRect.y1)
    
    ElseIf (uiMsg = WM_MOVE) Then
        UpdateWinRect
        RaiseEvent WindowMove(m_WindowRect.x1, m_WindowRect.y1)
    
    ElseIf (uiMsg = WM_NCDESTROY) Then
        EndSubclassing
    
    End If
    
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)

End Function
