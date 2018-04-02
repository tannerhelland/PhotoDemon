Attribute VB_Name = "ProgressBars"
'***************************************************************************
'Miscellaneous Functions Related to the Progress Bar
'Copyright 2001-2018 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/February/13
'Last update: Rewrite the progress bar code against an API progress bar on the main canvas object
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'API calls for our custom DoEvents replacement
Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type

Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As winMsg) As Long
Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

'This function mimicks DoEvents, but instead of processing all messages for all windows on all threads (slow! error-prone!),
' it only processes messages for the supplied hWnd.
Public Sub Replacement_DoEvents(ByVal srcHwnd As Long)
    Dim tmpMsg As winMsg
    Do While PeekMessageW(tmpMsg, srcHwnd, 0&, 0&, &H1&)
        TranslateMessage tmpMsg
        DispatchMessageW tmpMsg
    Loop
End Sub

'These three routines make it easier to interact with the progress bar; note that two are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar.
Public Function GetProgBarMax() As Double
    GetProgBarMax = FormMain.MainCanvas(0).ProgBar_GetMax()
    If (GetProgBarMax = 0#) Then GetProgBarMax = 1#
End Function

Public Sub SetProgBarMax(ByVal pbVal As Double)
    
    If (Macros.GetMacroStatus <> MacroBATCH) And (pbVal <> 0) Then
        
        Dim prevProgBarValue As Double
        prevProgBarValue = FormMain.MainCanvas(0).ProgBar_GetValue()
        FormMain.MainCanvas(0).ProgBar_SetVisibility True
        FormMain.MainCanvas(0).ProgBar_SetMax pbVal
        
        'Attempt to retain the progress bar's previous value, if any
        If (prevProgBarValue <= FormMain.MainCanvas(0).ProgBar_GetMax()) Then
            FormMain.MainCanvas(0).ProgBar_SetValue prevProgBarValue
        Else
            FormMain.MainCanvas(0).ProgBar_SetValue 0#
        End If
        
    End If
    
End Sub

Public Sub SetProgBarVal(ByVal pbVal As Double)
    
    If (Macros.GetMacroStatus <> MacroBATCH) Then
        
        FormMain.MainCanvas(0).ProgBar_SetValue pbVal
            
        'Process some window messages on the main form, to prevent the dreaded "Not Responding" state
        ' when PD is in the midst of a long-running action.
        Replacement_DoEvents FormMain.hWnd
        
        'On Windows 7 (or later), we also update the taskbar to reflect the current progress
        If OS.IsWin7OrLater Then OS.SetTaskbarProgressValue pbVal, GetProgBarMax
        
    End If
    
End Sub

'We only want the progress bar updating when necessary, so this function finds a power of 2 closest to the progress bar
' maximum divided by 20.  This is a nice compromise between responsive progress bar updates and extremely fast rendering.
Public Function FindBestProgBarValue() As Long

    'First, figure out what the range of this operation will be, based on the current progress bar maximum
    Dim progBarRange As Double
    progBarRange = CDbl(GetProgBarMax())
    
    'Divide that value by some arbitrary number; the number is how many times we want the progress bar to update during
    ' the current process.  (e.g. a value of "10" means "try to update the progress bar ~10 times")  Larger numbers
    ' mean more visual updates, at some minor cost to performance.
    progBarRange = progBarRange / 18#
    
    'Find the nearest power of two to that value, rounded down.  (We do this so that we can simply && the result on inner
    ' pixel processing loops, which is faster than a % operation.)
    Const LOG_TWO As Double = 0.693147180559945
    
    Dim nearestP2 As Long
    If (progBarRange > 0#) Then nearestP2 = Log(progBarRange) / LOG_TWO Else nearestP2 = 1
    FindBestProgBarValue = (2 ^ nearestP2) - 1
    
End Function

'When a function is done with the progress bar, this function must be called to free up its memory and hide the associated picture box
Public Sub ReleaseProgressBar()
    
    PDDebug.LogAction "Releasing progress bar..."
    
    'Briefly display a full progress bar before exiting
    FormMain.MainCanvas(0).ProgBar_SetValue FormMain.MainCanvas(0).ProgBar_GetMax()
    
    'Release the progress bar and container picture box
    FormMain.MainCanvas(0).ProgBar_SetVisibility False
    
    'On Win 7+, also reset the taskbar progress indicator
    If OS.IsWin7OrLater Then OS.SetTaskbarProgressState TBP_NoProgress
    
End Sub
