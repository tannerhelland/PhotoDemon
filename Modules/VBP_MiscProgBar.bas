Attribute VB_Name = "Misc_ProgressBar"
'***************************************************************************
'Miscellaneous Functions Related to the Progress Bar
'Copyright ©2000-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'***************************************************************************

Option Explicit

'These three routines make it easier to interact with the progress bar; note that two are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar for itself
Public Sub SetProgBarMax(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then cProgBar.Max = pbVal
End Sub

Public Function getProgBarMax() As Long
    getProgBarMax = cProgBar.Max
End Function

Public Sub SetProgBarVal(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then
        cProgBar.Value = pbVal
        cProgBar.Draw
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

