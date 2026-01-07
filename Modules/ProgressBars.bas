Attribute VB_Name = "ProgressBars"
'***************************************************************************
'Miscellaneous Functions Related to the Progress Bar
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/February/13
'Last update: Rewrite the progress bar code against an API progress bar on the main canvas object
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'These three routines make it easier to interact with the progress bar; note that two are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar.
Public Function GetProgBarMax() As Double
    GetProgBarMax = FormMain.MainCanvas(0).ProgBar_GetMax()
    If (GetProgBarMax = 0#) Then GetProgBarMax = 1#
End Function

Public Sub SetProgBarMax(ByVal pbVal As Double)
    
    If (Macros.GetMacroStatus <> MacroBATCH) And (pbVal <> 0#) Then
        
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
        
        'On Windows 7 (or later), we also update the taskbar to reflect the current progress
        If OS.IsWin7OrLater Then OS.SetTaskbarProgressValue pbVal, GetProgBarMax, FormMain.hWnd
        
        'Process some window messages on the main form, to prevent the dreaded "Not Responding" state
        ' when PD is in the midst of a long-running action.
        VBHacks.DoEvents_PaintOnly False
        
    End If
    
End Sub

'We only want to update the UI's progress bar when necessary, so this function finds a power of two
' closest to the current progress bar maximum divided by some arbitrary number (~20 works well).
' This is a nice compromise between responsive progress bar updates and extremely fast UI rendering.
Public Function FindBestProgBarValue(Optional ByVal customMaxValue As Double = 0#) As Long

    'First, figure out what the range of this operation will be, based on the current progress bar
    ' maximum or a user-supplied max value
    Dim progBarRange As Double
    If (customMaxValue > 0#) Then
        progBarRange = customMaxValue
    Else
        progBarRange = CDbl(ProgressBars.GetProgBarMax())
    End If
    
    'Divide that value by some arbitrary number; the number is how many times we want the progress bar
    ' to update during the current process.  (e.g. a value of "10" means "try to update the progress bar
    ' ~10 times")  Larger numbers mean more visual updates, at some cost to performance.
    progBarRange = progBarRange / 18#
    
    'Find the nearest power of two to that value, rounded down.  (We do this so that we can simply && the
    ' result on inner pixel processing loops, which is much faster than a % operation.)
    Const LOG_TWO As Double = 0.693147180559945
    
    Dim nearestP2 As Long
    If (progBarRange > 0#) Then nearestP2 = Int(Log(progBarRange) / LOG_TWO + 0.5) Else nearestP2 = 1
    FindBestProgBarValue = (2 ^ nearestP2) - 1
    
End Function

Public Function IsProgressBarVisible() As Boolean
    IsProgressBarVisible = FormMain.MainCanvas(0).ProgBar_GetVisibility()
End Function

'When a function is done with the progress bar, this function must be called to free up its memory
' and hide any associated UI elements
Public Sub ReleaseProgressBar()
    
    PDDebug.LogAction "Releasing progress bar..."
    
    'Reset the progress bar before exiting
    FormMain.MainCanvas(0).ProgBar_SetValue 0
    
    'Release the progress bar and container picture box
    FormMain.MainCanvas(0).ProgBar_SetVisibility False
    
    'On Win 7+, also reset the taskbar progress indicator
    If OS.IsWin7OrLater Then OS.SetTaskbarProgressState TBP_NoProgress, FormMain.hWnd
    
End Sub
