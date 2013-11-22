Attribute VB_Name = "ProgressBar_Support_Functions"
'***************************************************************************
'Miscellaneous Functions Related to the Progress Bar
'Copyright ©2001-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
Private Declare Function TranslateMessage Lib "user32" (lpMsg As winMsg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As winMsg) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

'This function mimicks DoEvents, but instead of processing all messages for all windows on all threads (slow! error-prone!),
' it only processes messages for the supplied hWnd.
Public Sub Replacement_DoEvents(ByVal srcHwnd As Long)

    Dim tmpMsg As winMsg
    Do While PeekMessage(tmpMsg, srcHwnd, 0&, 0&, &H1&)
        TranslateMessage tmpMsg
        DispatchMessage tmpMsg
    Loop
    
End Sub

'These three routines make it easier to interact with the progress bar; note that two are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar for itself
Public Sub SetProgBarMax(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then g_ProgBar.Max = pbVal
End Sub

Public Function getProgBarMax() As Long
    getProgBarMax = g_ProgBar.Max
End Function

Public Sub SetProgBarVal(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then
        g_ProgBar.Value = pbVal
        g_ProgBar.Draw
        
        'On Windows 7 (or later), we also update the taskbar to reflect the current progress
        If g_IsWin7OrLater Then SetTaskbarProgressValue pbVal, getProgBarMax
        
    End If
End Sub

'We only want the progress bar updating when necessary, so this function finds a power of 2 closest to the progress bar
' maximum divided by 20.  This is a nice compromise between responsive progress bar updates and extremely fast rendering.
Public Function findBestProgBarValue() As Long

    'First, figure out what the range of this operation will be using the values in curLayerValues
    Dim progBarRange As Double
    progBarRange = CDbl(getProgBarMax())
    
    'Divide that value by 20.  20 is an arbitrary selection; the value can be set to any value X, where X is the number
    ' of times we want the progress bar to update during a given filter or effect.
    progBarRange = progBarRange / 20
    
    'Find the nearest power of two to that value, rounded down
    Dim nearestP2 As Long
    
    nearestP2 = Log(progBarRange) / Log(2#)
    
    findBestProgBarValue = (2 ^ nearestP2) - 1
    
End Function

