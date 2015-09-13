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
'***************************************************************************
'PhotoDemon Zoom Handler - calculates and tracks zoom values for a given image
'Copyright 2001-2015 by Tanner Helland
'Created: 4/15/01
'Last updated: 01/October/14
'Last update: add support for "fit width", "fit height", and "fit all" zoom values, which are calculated dynamically.
'
'The main user of this class is the Viewport_Handler module.  Look there for relevant implementation details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

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
Private zoomCountFixed As Long

'Number of dynamic zoom entries, currently 3 - fit width, fit height, and fit all
Private zoomCountDynamic As Long

'This set of functions are simply wrappers that external code can use to access individual zoom entries
Public Function getZoomValue(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom value at that index
    If zoomIndex <= zoomCountFixed Then
        getZoomValue = zoomValues(zoomIndex)
        
    'If the zoom value is a dynamic entry, we need to calculate a specific zoom value at run-time
    Else
    
        'Make sure a valid image is loaded and ready
        If (g_OpenImageCount > 0) And Not (pdImages(g_CurrentImage) Is Nothing) Then
        
            'Retrieve the current image's width and height
            Dim imgWidth As Double, imgHeight As Double
            imgWidth = pdImages(g_CurrentImage).Width
            imgHeight = pdImages(g_CurrentImage).Height
            
            'Retrieve the current viewport's width and height
            Dim viewportWidth As Double, viewportHeight As Double
            viewportWidth = FormMain.mainCanvas(0).getCanvasWidth
            viewportHeight = FormMain.mainCanvas(0).getCanvasHeight
            
            'Calculate a width and height ratio in advance
            Dim horizontalRatio As Double, verticalRatio As Double
            horizontalRatio = viewportWidth / imgWidth
            verticalRatio = viewportHeight / imgHeight
            
            Select Case zoomIndex
            
                'Fit width
                Case zoomCountFixed + 1
                
                    'Check to see if the calculated zoom value will require a vertical scroll bar (since we are only fitting width).
                    ' If it will, we must subtract the scroll bar's width from our calculation.
                    If imgHeight * horizontalRatio > viewportHeight Then
                        getZoomValue = viewportWidth / imgWidth
                    Else
                        getZoomValue = horizontalRatio
                    End If
                    
                'Fit height
                Case zoomCountFixed + 2
                
                    'Check to see if the calculated zoom value will require a horizontal scroll bar (since we are only fitting height).
                    ' If it will, we must subtract the scroll bar's height from our calculation.
                    If imgWidth * verticalRatio > viewportWidth Then
                        getZoomValue = viewportHeight / imgHeight
                    Else
                        getZoomValue = verticalRatio
                    End If
                    
                'Fit everything
                Case zoomCountFixed + 3
                    If horizontalRatio < verticalRatio Then
                        getZoomValue = horizontalRatio
                    Else
                        getZoomValue = verticalRatio
                    End If
            
            End Select
            
        Else
            getZoomValue = 1
        End If
        
    
    End If
    
End Function

Public Function getZoomOffsetFactor(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom offset at that index
    If zoomIndex <= zoomCountFixed Then
        getZoomOffsetFactor = zoomOffsetFactors(zoomIndex)
    
    'If the zoom value is a dynamic entry, we need to calculate a specific zoom offset at run-time
    Else
    
        Dim curZoomValue As Double
        curZoomValue = getZoomValue(zoomIndex)
        
        If curZoomValue >= 1 Then
            getZoomOffsetFactor = curZoomValue
        Else
            getZoomOffsetFactor = 1 / curZoomValue
        End If
    
    End If
    
End Function

'To minimize the possibility of program-wide changes if I ever decide to fiddle with PD's fixed zoom values, these functions are used
' externally to retrieve specific zoom indices.
Public Function getZoom100Index() As Long
    getZoom100Index = ZOOM_100_PERCENT
End Function

Public Function getZoomFitWidthIndex() As Long
    getZoomFitWidthIndex = zoomCountFixed + 1
End Function

Public Function getZoomFitHeightIndex() As Long
    getZoomFitHeightIndex = zoomCountFixed + 2
End Function

Public Function getZoomFitAllIndex() As Long
    getZoomFitAllIndex = zoomCountFixed + 3
End Function

Public Function getZoomCount() As Long
    getZoomCount = zoomCountFixed
End Function

'Whenever one of these classes is created, remember to call this initialization function.  It will manually prepare a
' list of zoom values relevant to the program.
Public Sub initializeViewportEngine()

    'This list of zoom values is (effectively) arbitrary.  I've based this list off similar lists (Paint.NET, GIMP)
    ' while including a few extra values for convenience's sake
    
    'Total number of fixed zoom values.  Some legacy PD functions (like the old Fit to Screen code) require this so
    ' they can iterate all fixed zoom values, and find an appropriate one for their purpose.
    zoomCountFixed = 25
    
    'Total number of dynamic zoom values, e.g. values dynamically calculated on a per-image basis.  At present these include:
    ' fit width, fit height, and fit all
    zoomCountDynamic = 3
    
    'Prepare our zoom array.
    ReDim zoomStrings(0 To zoomCountFixed + zoomCountDynamic) As String
    ReDim zoomValues(0 To zoomCountFixed + zoomCountDynamic) As Double
    ReDim zoomOffsetFactors(0 To zoomCountFixed + zoomCountDynamic) As Double
    
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
    
    zoomStrings(26) = g_Language.TranslateMessage("Fit width")
        zoomValues(26) = 0
        zoomOffsetFactors(26) = 0
    
    zoomStrings(27) = g_Language.TranslateMessage("Fit height")
        zoomValues(27) = 0
        zoomOffsetFactors(27) = 0
        
    zoomStrings(28) = g_Language.TranslateMessage("Fit image")
        zoomValues(28) = 0
        zoomOffsetFactors(28) = 0
    
    'Note which index corresponds to 100%
    ZOOM_100_PERCENT = 11
    
End Sub

'Populate an arbitrary combo box with the current list of handled zoom values
Public Sub populateZoomComboBox(ByRef dstComboBox As pdComboBox, Optional ByVal initialListIndex As Long = -1)

    dstComboBox.Clear
    
    Dim i As Long
    
    For i = 0 To zoomCountFixed + zoomCountDynamic
        
        Select Case i
        
            Case 10, 11, 25
                dstComboBox.AddItem zoomStrings(i), i, True
                
            Case Else
                dstComboBox.AddItem zoomStrings(i), i
        
        End Select
        
    Next i
    
    If initialListIndex = -1 Then
        dstComboBox.ListIndex = ZOOM_100_PERCENT
    Else
        dstComboBox.ListIndex = initialListIndex
    End If

End Sub

'Given a current zoom index, find the nearest relevant "zoom in" index.  This requires special handling in the case of "fit image on screen".
Public Function getNearestZoomInIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-in index is easy.
    If curIndex <= zoomCountFixed Then
        
        getNearestZoomInIndex = curIndex - 1
        If getNearestZoomInIndex < 0 Then getNearestZoomInIndex = 0
    
    'If the current zoom index is one of the "fit" options, this is more complicated.  We want to set the first fixed index we
    ' find that is larger than the current dynamic value being used.
    Else
    
        Dim curZoomValue As Double
        curZoomValue = getZoomValue(curIndex)
        
        'Start searching the zoom array for the nearest value that is larger than the current zoom value.
        Dim i As Long
        For i = zoomCountFixed To 0 Step -1
            If zoomValues(i) > curZoomValue Then
                getNearestZoomInIndex = i
                Exit For
            End If
        Next i
    
    End If

End Function

'Given a current zoom index, find the nearest relevant "zoom out" index.  This requires special handling in the case of "fit image on screen".
Public Function getNearestZoomOutIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-out index is easy.
    If curIndex <= zoomCountFixed Then
        
        getNearestZoomOutIndex = curIndex + 1
        If getNearestZoomOutIndex > zoomCountFixed Then getNearestZoomOutIndex = zoomCountFixed
    
    'If the current zoom index is one of the "fit" options, this is more complicated.  We want to set the first fixed index we
    ' find that is smaller than the current dynamic value being used.
    Else
    
        Dim curZoomValue As Double
        curZoomValue = getZoomValue(curIndex)
        
        'Start searching the zoom array for the nearest value that is larger than the current zoom value.
        Dim i As Long
        For i = 0 To zoomCountFixed
            If zoomValues(i) < curZoomValue Then
                getNearestZoomOutIndex = i
                Exit For
            End If
        Next i
    
    End If

End Function

