VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUCSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Master User Control Support class
'Copyright 2015-2015 by Tanner Helland
'Created: 25/October/15  (but built from many parts existing earlier)
'Last updated: 25/October/15
'Last update: initial build
'
'Now that PD provides so many of its own user-controls, we have to do a lot of repeat UC tasks.  Painting, custom
' input handling, subclassing custom window messages - the list goes on for awhile.
'
'To cut down on the amount of repeat work I have to do for user controls, I'm trying to wrap as many generic
' UC capabilities as I can into this single class.  It then raises a boatload of messages for individual UCs to
' handle as they see fit.
'
'In the interim, some user controls may still implement many of these features manually.  Migrating them over is
' a WIP, and one I hope to have done by the 7.0 release.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class raises many events.  Some *must* be responded to.  Others are optional.  These are poorly documented
' at present - sorry!
Public Event CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean)

'This class subclasses a *lot* of things.
Private m_Subclass As cSelfSubHookCallback

'Persistent copy of the associated user control's hWnd (and its container, too, for good measure)
Private m_hWnd As Long, m_hWndParent As Long

'If the caller has registered one (or more) custom messages, this will be set to TRUE.  When the parent control is
' unloaded, we need to de-register the hWnd with PD's central message marshaler.
Private m_CustomInternalMessagesActive As Boolean

Public Sub RegisterControl(ByVal hWnd_Control As Long, ByVal hWnd_Container As Long)
    
    m_hWnd = hWnd_Control
    m_hWndParent = hWnd_Container
    
    #If DEBUGMODE = 1 Then
        If m_hWnd = 0 Then pdDebug.LogAction "WARNING!  pdUCSupport.RegisterControl was called with hWnd = 0."
    #End If
    
    'Subclass the user control immediately
    If m_hWnd <> 0 Then
        m_Subclass.ssc_Subclass m_hWnd, , 1, Me, True, True
    End If
        
    'TODO: subclass everything
    
End Sub

'Want to receive custom window messages?  No problem.  Register them here, then watch for the CustomMessage event.
' Note, however, that if the message is specific to PD, the caller needs to let us know, as we will need to manually
' handle message marshaling.
Public Sub SubclassCustomMessage(ByVal wMsg As Long, Optional ByVal msgIsInternalToPD As Boolean = False)
    
    m_Subclass.ssc_AddMsg m_hWnd, MSG_BEFORE, wMsg
    
    'If the message is internal to PD, register the hWnd and message with the central marshaler now
    If msgIsInternalToPD Then
        UserControl_Support.AddMessageRecipient m_hWnd, wMsg
        m_CustomInternalMessagesActive = True
    End If
    
End Sub

Private Sub Class_Initialize()
    Set m_Subclass = New cSelfSubHookCallback
End Sub

Private Sub Class_Terminate()
    
    'Release all subclassing now
    m_Subclass.ssc_Terminate
    Set m_Subclass = Nothing
    
    'If the user control was registered for internal PD messages, de-register it now.
    If m_CustomInternalMessagesActive Then UserControl_Support.RemoveMessageRecipient m_hWnd
    
End Sub

'All subclassed messages are handled here.
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
    
        'All other messages are custom requests from the window.  Use them to raise custom message events.
        Case Else
            RaiseEvent CustomMessage(uMsg, wParam, lParam, bHandled)
    
    End Select


' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
