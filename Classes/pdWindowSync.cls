VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWindowSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Window Synchronization class
'Copyright 2015-2015 by Tanner Helland
'Created: 26/September/15
'Last updated: 26/September/15
'Last update: initial build
'
'The right-side panel in the main PD window supports a number of collapsible panels.  Each panel lives on its own
' form, which helps keep the code organized (as there are so many panels, each with totally different UI needs).
'
'Because it's a pain to manually move/size those forms as panels are moved and/or sized, the parent form uses an
' array of picture boxes to control subpanel layout.  This class then does the messy business of synchronizing the
' various subpanel forms with their representative picture boxes, which allows us to keep the main toolbar and various
' subpanels nicely compartmentalized, while taking advantage of built-in Windows messaging for relaying size/position
' changes to the subpanels whenever the parent window needs to make changes.
'
'Note that this class requires a 1:1 correlation between parent and child hWnds.  You cannot map multiple children
' to the same parent, by design.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'SetWindowPos flags
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

'Subclassed window messages
Private Const WM_SIZE As Long = &H5

'Currently subclassed hWnds.  Each key is a parent hWnd, while each value is the associated child hWnd.
Private m_windowDict As pdDictionary

'Each window is subclassed as it is added.
Private m_Subclass As cSelfSubHookCallback

'Temporary window calculation rect; a single declaration here results in less churn than a local one inside the subclass proc.
Private m_newRect As winRect

Private Sub Class_Initialize()
    Set m_windowDict = New pdDictionary
End Sub

Private Sub Class_Terminate()
    If Not (m_Subclass Is Nothing) Then
        m_Subclass.ssc_Terminate
    End If
End Sub

Public Sub SynchronizeWindows(ByVal parentHwnd As Long, ByVal childHwnd As Long)
    
    Dim strKey As String
    strKey = Trim$(Str(parentHwnd))
    
    'Make sure this window doesn't already exist in the collection
    If Not m_windowDict.doesKeyExist(strKey) Then
        
        'Store the pair of hWnds
        m_windowDict.AddEntry strKey, childHwnd
        
        'Subclass size and visibility messages on the parent window
        If m_Subclass Is Nothing Then Set m_Subclass = New cSelfSubHookCallback
        m_Subclass.ssc_Subclass parentHwnd, childHwnd, 1, Me
        m_Subclass.ssc_AddMsg parentHwnd, MSG_BEFORE, WM_SIZE
        
        'Make the associated window a child of the parent; this automates position and visibility handling
        SetParent childHwnd, parentHwnd
        
        'Perform an initial synchronization to position (0, 0) and the size of the parent window
        Dim parentRect As winRect
        GetClientRect parentHwnd, parentRect
        SetWindowPosition childHwnd, 0, 0
        SetWindowSize childHwnd, parentRect.x2, parentRect.y2
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  Duplicated hWnd passed to pdWindowSync.SynchronizeWindows.  Fix this!"
        #End If
    End If
    
End Sub

'Move or resize a window
Private Sub SetWindowPosition(ByVal srcHwnd As Long, ByVal newLeft As Long, ByVal newTop As Long, Optional ByVal notifyWindow As Boolean = True)
    
    'Per VB convention, the new left/top coordinates should be in parent coordinates.  If you want to position a
    ' window using *screen coordinates*, you'll need to write a new function (and/or manually convert the coordinates
    ' to screen space in advance).
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    If Not notifyWindow Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos srcHwnd, 0&, newLeft, newTop, 0&, 0&, swpFlags
    
End Sub

Private Sub SetWindowSize(ByVal srcHwnd As Long, ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)
    
    Dim swpFlags As SWP_FLAGS
    swpFlags = SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE
    If Not notifyWindow Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    
    SetWindowPos srcHwnd, 0&, 0&, 0&, newWidth, newHeight, swpFlags
    
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
            
            'Retrieve a copy of the parent window's rect (because the windows are borderless, GetClientRect works fine)
            GetClientRect lng_hWnd, m_newRect
            
            'Resize the child window accordingly.
            SetWindowSize lParamUser, m_newRect.x2, m_newRect.y2
            
    End Select
    

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
