VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Array index of the zoom array entry that corresponds to 100% zoom.  Calculated manually and treated as a constant.
Private ZOOM_100_PERCENT As Long

'Human-friendly string for each zoom value (e.g. "100%" for 1.0 zoom)
Private zoomStrings() As String

'Actual multipliers for each zoom value (e.g. 2 for 2.0 zoom, 0.5 for 50% zoom)
Private zoomValues() As Double

'When zoomed-out, images will distort when scrolled if they are not locked to multiples of the current zoom factor.
' This array stores the offset factors necessary to fix such scrolling bugs.
Private zoomOffsetFactors() As Double

'Upper bound of primary zoom array (e.g. number of unique zoom values - 1)
Private zoomCount As Long

'This set of functions are simply wrappers that external code can use to access individual zoom entries
Public Function getZoomValue(ByVal zoomIndex As Long) As Double
    getZoomValue = zoomValues(zoomIndex)
End Function

Public Function getZoomOffsetFactor(ByVal zoomIndex As Long) As Double
    getZoomOffsetFactor = zoomOffsetFactors(zoomIndex)
End Function

Public Function getZoom100Index() As Long
    getZoom100Index = ZOOM_100_PERCENT
End Function

Public Function getZoomCount() As Long
    getZoomCount = zoomCount
End Function

'Whenever one of these classes is created, remember to call this initialization function.  It will manually prepare a
' list of zoom values relevant to the program.
Public Sub initializeViewportEngine()

    'This list of zoom values is (effectively) arbitrary.  I've based this list off similar lists (Paint.NET, GIMP)
    ' while including a few extra values for convenience's sake
    
    'Total number of available zoom values
    zoomCount = 25
    
    ReDim zoomStrings(0 To zoomCount) As String
    ReDim zoomValues(0 To zoomCount) As Double
    ReDim zoomOffsetFactors(0 To zoomCount) As Double
    
    'Manually create a list of user-friendly zoom values
    zoomStrings(0) = "3200%"
        zoomValues(0) = 32
        zoomOffsetFactors(0) = 32
        
    zoomStrings(1) = "2400%"
        zoomValues(1) = 24
        zoomOffsetFactors(1) = 24
        
    zoomStrings(2) = "1600%"
        zoomValues(2) = 16
        zoomOffsetFactors(2) = 16
        
    zoomStrings(3) = "1200%"
        zoomValues(3) = 12
        zoomOffsetFactors(3) = 12
        
    zoomStrings(4) = "800%"
        zoomValues(4) = 8
        zoomOffsetFactors(4) = 8
        
    zoomStrings(5) = "700%"
        zoomValues(5) = 7
        zoomOffsetFactors(5) = 7
        
    zoomStrings(6) = "600%"
        zoomValues(6) = 6
        zoomOffsetFactors(6) = 6
        
    zoomStrings(7) = "500%"
        zoomValues(7) = 5
        zoomOffsetFactors(7) = 5
        
    zoomStrings(8) = "400%"
        zoomValues(8) = 4
        zoomOffsetFactors(8) = 4
        
    zoomStrings(9) = "300%"
        zoomValues(9) = 3
        zoomOffsetFactors(9) = 3
        
    zoomStrings(10) = "200%"
        zoomValues(10) = 2
        zoomOffsetFactors(10) = 2
        
    zoomStrings(11) = "100%"
        zoomValues(11) = 1
        zoomOffsetFactors(11) = 1
        
    zoomStrings(12) = "75%"
        zoomValues(12) = 3 / 4
        zoomOffsetFactors(12) = 4 / 3
        
    zoomStrings(13) = "67%"
        zoomValues(13) = 2 / 3
        zoomOffsetFactors(13) = 3 / 2
        
    zoomStrings(14) = "50%"
        zoomValues(14) = 0.5
        zoomOffsetFactors(14) = 2
        
    zoomStrings(15) = "33%"
        zoomValues(15) = 1 / 3
        zoomOffsetFactors(15) = 3
        
    zoomStrings(16) = "25%"
        zoomValues(16) = 0.25
        zoomOffsetFactors(16) = 4
        
    zoomStrings(17) = "20%"
        zoomValues(17) = 0.2
        zoomOffsetFactors(17) = 5
        
    zoomStrings(18) = "16%"
        zoomValues(18) = 0.16
        zoomOffsetFactors(18) = 100 / 16
        
    zoomStrings(19) = "12%"
        zoomValues(19) = 0.12
        zoomOffsetFactors(19) = 100 / 12
        
    zoomStrings(20) = "8%"
        zoomValues(20) = 0.08
        zoomOffsetFactors(20) = 100 / 8
        
    zoomStrings(21) = "6%"
        zoomValues(21) = 0.06
        zoomOffsetFactors(21) = 100 / 6
        
    zoomStrings(22) = "4%"
        zoomValues(22) = 0.04
        zoomOffsetFactors(22) = 25
        
    zoomStrings(23) = "3%"
        zoomValues(23) = 0.03
        zoomOffsetFactors(23) = 100 / 0.03
        
    zoomStrings(24) = "2%"
        zoomValues(24) = 0.02
        zoomOffsetFactors(24) = 50
        
    zoomStrings(25) = "1%"
        zoomValues(25) = 0.01
        zoomOffsetFactors(25) = 100
    
    'Note which index corresponds to 100%
    ZOOM_100_PERCENT = 11
    
End Sub

'Populate an arbitrary combo box with the current list of handled zoom values
Public Sub populateZoomComboBox(ByRef dstComboBox As ComboBox, Optional ByVal initialListIndex As Long = -1)

    dstComboBox.Clear
    
    Dim i As Long
    
    For i = 0 To zoomCount
        dstComboBox.AddItem zoomStrings(i), i
    Next i
    
    If initialListIndex = -1 Then
        dstComboBox.ListIndex = ZOOM_100_PERCENT
    Else
        dstComboBox.ListIndex = initialListIndex
    End If

End Sub
