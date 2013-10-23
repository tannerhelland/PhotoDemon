VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdVisualThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Visual Theming class
'Copyright ©2012-2013 by Tanner Helland
'Created: 23/October/13
'Last updated: 23/October/13
'Last update: initial build (still in progress)
'
'There isn't much here yet, but stay tuned - more is coming!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'In order to make themed controls transparent, we must subclass their parent containers, and manually process the
' WM_PRINTCLIENT message using the API decs below.  Many thanks to pro VB programmer LaVolpe for this workaround.
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_PAINT As Long = &HF&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'cSubclass handles all subclassing duties for us
Private cSubclass As cSelfSubHookCallback

'When a container object needs to render contained controls transparently, register it using this function
Public Sub requestContainerSubclass(ByVal containerHwnd As Long)

    If (cSubclass Is Nothing) Then
        Set cSubclass = New cSelfSubHookCallback
    End If
    
    cSubclass.ssc_Subclass containerHwnd, , , Me
    cSubclass.ssc_AddMsg containerHwnd, MSG_BEFORE, WM_PRINTCLIENT

End Sub

'While not strictly required, it is good behavior to unsubclass container objects when they are unloaded
Public Sub releaseContainerSubclass(ByVal containerHwnd As Long)
    cSubclass.ssc_UnSubclass containerHwnd
End Sub

'All events subclassed by this class are processed here.
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

    If uMsg = WM_PRINTCLIENT Then
        SendMessage lng_hWnd, WM_PAINT, wParam, ByVal 0&
        bHandled = True
    End If

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
End Sub

