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
'Copyright 2015-2015 by Tanner Helland
'Created: 23/September/15 (but built from many parts existing earlier)
'Last updated: 23/September/15
'Last update: initial build
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
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetParent Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal srcHwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal srcHwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lpPoints As POINTAPI, ByVal cPoints As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

'Optional window messages to subclass
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5

'Subclassing is optional; if active, this class will be non-nothing.
Private cSubclass As cSelfSubHookCallback
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
Public Sub AttachToHWnd(ByVal srcHwnd As Long, Optional ByVal subclassMoveSizeEvents As Boolean = False)
    
    If srcHwnd = 0 Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdWindowSize.AttachToHWnd received a null hWnd - was this intentional??"
        #End If
    Else
    
        m_hWnd = srcHwnd
        
        'Activate the subclasser conditionally; some UCs don't need to receive API size/move events
        m_SubclassActive = subclassMoveSizeEvents
        
        If m_SubclassActive Then
            Set cSubclass = New cSelfSubHookCallback
            cSubclass.ssc_Subclass m_hWnd, , , Me
            cSubclass.ssc_AddMsg m_hWnd, MSG_AFTER, WM_SIZE, WM_MOVE
        End If
        
        'Cache an initial set of window dimensions
        UpdateWinRect
        
    End If
    
End Sub

'FYI: all retrieval functions operate purely in the pixel domain
Public Function GetWidth() As Long
    If Not m_SubclassActive Then UpdateWinRect
    GetWidth = m_WindowRect.x2 - m_WindowRect.x1
End Function

Public Function GetHeight() As Long
    If Not m_SubclassActive Then UpdateWinRect
    GetHeight = m_WindowRect.y2 - m_WindowRect.y1
End Function

Public Function GetLeft() As Long
    If Not m_SubclassActive Then UpdateWinRect
    GetLeft = m_WindowRect.x1
End Function

Public Function GetTop() As Long
    If Not m_SubclassActive Then UpdateWinRect
    GetTop = m_WindowRect.y1
End Function

'Move or resize a window
Public Sub SetPosition(ByVal newLeft As Long, ByVal newTop As Long, Optional ByVal notifyWindow As Boolean = True)
    
    'Per VB convention, the new left/top coordinates should be in parent coordinates.  If you want to position a
    ' window using *screen coordinates*, you'll need to write a new function (and/or manually convert the coordinates
    ' to screen space in advance).
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    If Not notifyWindow Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos m_hWnd, 0&, newLeft, newTop, 0, 0, swpFlags
    
End Sub

Public Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE
    If Not notifyWindow Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos m_hWnd, 0&, 0, 0, newWidth, newHeight, swpFlags
    
End Sub

'Update the module-level copy of the attached hWnd's rect, automatically translated into our parent's coordinate space.
Private Sub UpdateWinRect()
    
    'Retrieve a copy of the window rect *in screen coordinates*
    GetWindowRect m_hWnd, m_WindowRect
    
    'Find the difference between *the screen coordinate space* and *the parent's coordinate space*
    Dim tmpPoint As POINTAPI
    tmpPoint.x = m_WindowRect.x1
    tmpPoint.y = m_WindowRect.y1
    MapWindowPoints 0&, GetParent(m_hWnd), tmpPoint, 1&
    
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
    
End Sub

Private Sub Class_Initialize()
    m_hWnd = 0
End Sub

Private Sub Class_Terminate()
    If Not (cSubclass Is Nothing) Then
        cSubclass.ssc_UnSubclass m_hWnd
        m_hWnd = 0
        cSubclass.ssc_Terminate
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

    Select Case uMsg
    
        Case WM_SIZE
            UpdateWinRect
            RaiseEvent WindowResize(m_WindowRect.x2 - m_WindowRect.x1, m_WindowRect.y2 - m_WindowRect.y1)
        
        Case WM_MOVE
            UpdateWinRect
            RaiseEvent WindowMove(m_WindowRect.x1, m_WindowRect.y1)
        
    End Select
    

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
