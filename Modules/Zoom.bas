Attribute VB_Name = "Zoom"
'***************************************************************************
'PhotoDemon Zoom Handler - calculates and tracks zoom values for a given image
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 11/September/25
'Last update: add 150% zoom as a preset value (see https://github.com/tannerhelland/PhotoDemon/issues/666)
'
'The main user of this class is the Viewport_Handler module.  Look there for relevant implementation details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Array index of the zoom array entry that corresponds to 100% zoom.  Calculated manually and treated as a constant.
Private ZOOM_100_PERCENT As Long

'Human-friendly string for each zoom value (e.g. "100%" for 1.0 zoom)
Private m_zoomStrings() As String

'Actual multipliers for each zoom value (e.g. 2 for 2.0 zoom, 0.5 for 50% zoom)
Private m_zoomValues() As Double

'When zoomed-out, images will distort when scrolled if they are not locked to multiples of the current zoom factor.
' This array stores the offset factors necessary to fix such scrolling bugs.
Private m_zoomOffsetFactors() As Double

'Upper bound of primary zoom array (e.g. number of unique zoom values - 1)
Private m_zoomCountFixed As Long

'Number of dynamic zoom entries, currently 3 - fit width, fit height, and fit all
Private m_zoomCountDynamic As Long

'This set of functions are simply wrappers that external code can use to access individual zoom entries
Public Function GetZoomRatioFromIndex(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom value at that index
    If (zoomIndex <= m_zoomCountFixed) Then
        GetZoomRatioFromIndex = m_zoomValues(zoomIndex)
        
    'If the zoom value is a dynamic entry, we need to calculate a specific zoom value at run-time
    Else
    
        'Make sure a valid image is loaded and ready
        If PDImages.IsImageActive() Then
        
            'Retrieve the current image's width and height
            Dim imgWidth As Double, imgHeight As Double
            imgWidth = PDImages.GetActiveImage.Width
            imgHeight = PDImages.GetActiveImage.Height
            
            'Retrieve the current viewport's width and height
            Dim viewportWidth As Double, viewportHeight As Double
            viewportWidth = FormMain.MainCanvas(0).GetCanvasWidth
            viewportHeight = FormMain.MainCanvas(0).GetCanvasHeight
            
            'Calculate a width and height ratio in advance
            Dim horizontalRatio As Double, verticalRatio As Double
            If (imgHeight <> 0) And (imgWidth <> 0) Then
            
                horizontalRatio = viewportWidth / imgWidth
                verticalRatio = viewportHeight / imgHeight
                
                Select Case zoomIndex
                
                    'Fit width
                    Case m_zoomCountFixed + 1
                    
                        'Check to see if the calculated zoom value will require a vertical scroll bar (since we are only fitting width).
                        ' If it will, we must subtract the scroll bar's width from our calculation.
                        If (imgHeight * horizontalRatio > viewportHeight) Then
                            GetZoomRatioFromIndex = viewportWidth / imgWidth
                        Else
                            GetZoomRatioFromIndex = horizontalRatio
                        End If
                        
                    'Fit height
                    Case m_zoomCountFixed + 2
                    
                        'Check to see if the calculated zoom value will require a horizontal scroll bar (since we are only fitting height).
                        ' If it will, we must subtract the scroll bar's height from our calculation.
                        If (imgWidth * verticalRatio > viewportWidth) Then
                            GetZoomRatioFromIndex = viewportHeight / imgHeight
                        Else
                            GetZoomRatioFromIndex = verticalRatio
                        End If
                        
                    'Fit everything
                    Case m_zoomCountFixed + 3
                        If (horizontalRatio < verticalRatio) Then
                            GetZoomRatioFromIndex = horizontalRatio
                        Else
                            GetZoomRatioFromIndex = verticalRatio
                        End If
                
                End Select
                
            Else
                GetZoomRatioFromIndex = 1#
            End If
            
        Else
            GetZoomRatioFromIndex = 1#
        End If
        
    
    End If
    
End Function

Public Function GetZoomOffsetFactor(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom offset at that index
    If (zoomIndex <= m_zoomCountFixed) Then
        GetZoomOffsetFactor = m_zoomOffsetFactors(zoomIndex)
    
    'If the zoom value is a dynamic entry, we need to calculate a specific zoom offset at run-time
    Else
    
        Dim curZoomValue As Double
        curZoomValue = GetZoomRatioFromIndex(zoomIndex)
        
        If (curZoomValue >= 1#) Then
            GetZoomOffsetFactor = curZoomValue
        Else
            GetZoomOffsetFactor = 1# / curZoomValue
        End If
    
    End If
    
End Function

'To minimize the possibility of program-wide changes if I ever decide to fiddle with PD's fixed zoom values, these functions are used
' externally to retrieve specific zoom indices.
Public Function GetZoom100Index() As Long
    GetZoom100Index = ZOOM_100_PERCENT
End Function

Public Function GetZoomFitWidthIndex() As Long
    GetZoomFitWidthIndex = m_zoomCountFixed + 1
End Function

Public Function GetZoomFitHeightIndex() As Long
    GetZoomFitHeightIndex = m_zoomCountFixed + 2
End Function

Public Function GetZoomFitAllIndex() As Long
    GetZoomFitAllIndex = m_zoomCountFixed + 3
End Function

Public Function GetZoomCount() As Long
    GetZoomCount = m_zoomCountFixed
End Function

'Whenever one of these classes is created, remember to call this initialization function.  It will manually prepare a
' list of zoom values relevant to the program.
Public Sub InitializeZoomEngine()

    'This list of zoom values is (effectively) arbitrary.  I've based this list off similar lists (Paint.NET, GIMP)
    ' while including a few extra values for convenience's sake
    
    'Total number of fixed zoom values.  Some legacy PD functions (like the old Fit to Screen code) require this so
    ' they can iterate all fixed zoom values, and find an appropriate one for their purpose.
    m_zoomCountFixed = 26
    
    'Total number of dynamic zoom values, e.g. values dynamically calculated on a per-image basis.  At present these include:
    ' fit width, fit height, and fit all
    m_zoomCountDynamic = 3
    
    'Prepare our zoom array.
    ReDim m_zoomStrings(0 To m_zoomCountFixed + m_zoomCountDynamic) As String
    ReDim m_zoomValues(0 To m_zoomCountFixed + m_zoomCountDynamic) As Double
    ReDim m_zoomOffsetFactors(0 To m_zoomCountFixed + m_zoomCountDynamic) As Double
    
    'Manually create a list of user-friendly zoom values
    m_zoomStrings(0) = "3200%"
        m_zoomValues(0) = 32
        m_zoomOffsetFactors(0) = 32
        
    m_zoomStrings(1) = "2400%"
        m_zoomValues(1) = 24
        m_zoomOffsetFactors(1) = 24
        
    m_zoomStrings(2) = "1600%"
        m_zoomValues(2) = 16
        m_zoomOffsetFactors(2) = 16
        
    m_zoomStrings(3) = "1200%"
        m_zoomValues(3) = 12
        m_zoomOffsetFactors(3) = 12
        
    m_zoomStrings(4) = "800%"
        m_zoomValues(4) = 8
        m_zoomOffsetFactors(4) = 8
        
    m_zoomStrings(5) = "700%"
        m_zoomValues(5) = 7
        m_zoomOffsetFactors(5) = 7
        
    m_zoomStrings(6) = "600%"
        m_zoomValues(6) = 6
        m_zoomOffsetFactors(6) = 6
        
    m_zoomStrings(7) = "500%"
        m_zoomValues(7) = 5
        m_zoomOffsetFactors(7) = 5
        
    m_zoomStrings(8) = "400%"
        m_zoomValues(8) = 4
        m_zoomOffsetFactors(8) = 4
        
    m_zoomStrings(9) = "300%"
        m_zoomValues(9) = 3
        m_zoomOffsetFactors(9) = 3
        
    m_zoomStrings(10) = "200%"
        m_zoomValues(10) = 2
        m_zoomOffsetFactors(10) = 2
        
    m_zoomStrings(11) = "150%"
        m_zoomValues(11) = 1.5
        m_zoomOffsetFactors(11) = 1.5
        
    m_zoomStrings(12) = "100%"
        m_zoomValues(12) = 1
        m_zoomOffsetFactors(12) = 1
        
    m_zoomStrings(13) = "75%"
        m_zoomValues(13) = 3# / 4#
        m_zoomOffsetFactors(13) = 4# / 3#
        
    m_zoomStrings(14) = "67%"
        m_zoomValues(14) = 2# / 3#
        m_zoomOffsetFactors(14) = 3# / 2#
        
    m_zoomStrings(15) = "50%"
        m_zoomValues(15) = 0.5
        m_zoomOffsetFactors(15) = 2#
        
    m_zoomStrings(16) = "33%"
        m_zoomValues(16) = 1# / 3#
        m_zoomOffsetFactors(16) = 3
        
    m_zoomStrings(17) = "25%"
        m_zoomValues(17) = 0.25
        m_zoomOffsetFactors(17) = 4
        
    m_zoomStrings(18) = "20%"
        m_zoomValues(18) = 0.2
        m_zoomOffsetFactors(18) = 5
        
    m_zoomStrings(19) = "16%"
        m_zoomValues(19) = 0.16
        m_zoomOffsetFactors(19) = 100# / 16#
        
    m_zoomStrings(20) = "12%"
        m_zoomValues(20) = 0.12
        m_zoomOffsetFactors(20) = 100# / 12#
        
    m_zoomStrings(21) = "8%"
        m_zoomValues(21) = 0.08
        m_zoomOffsetFactors(21) = 100# / 8#
        
    m_zoomStrings(22) = "6%"
        m_zoomValues(22) = 0.06
        m_zoomOffsetFactors(22) = 100# / 6#
        
    m_zoomStrings(23) = "4%"
        m_zoomValues(23) = 0.04
        m_zoomOffsetFactors(23) = 25
        
    m_zoomStrings(24) = "3%"
        m_zoomValues(24) = 0.03
        m_zoomOffsetFactors(24) = 100# / 0.03
        
    m_zoomStrings(25) = "2%"
        m_zoomValues(25) = 0.02
        m_zoomOffsetFactors(25) = 50
        
    m_zoomStrings(26) = "1%"
        m_zoomValues(26) = 0.01
        m_zoomOffsetFactors(26) = 100
    
    m_zoomStrings(27) = g_Language.TranslateMessage("Fit width")
        m_zoomValues(27) = 0
        m_zoomOffsetFactors(27) = 0
    
    m_zoomStrings(28) = g_Language.TranslateMessage("Fit height")
        m_zoomValues(28) = 0
        m_zoomOffsetFactors(28) = 0
        
    m_zoomStrings(29) = g_Language.TranslateMessage("Fit image")
        m_zoomValues(29) = 0
        m_zoomOffsetFactors(29) = 0
    
    'Note which index corresponds to 100%
    ZOOM_100_PERCENT = 12
    
End Sub

'Populate an arbitrary combo box with the current list of handled zoom values
Public Sub PopulateZoomDropdown(ByRef dstDropDown As pdDropDown, Optional ByVal initialListIndex As Long = -1)
    
    dstDropDown.SetAutomaticRedraws False
    dstDropDown.Clear
    
    Dim i As Long
    
    For i = 0 To m_zoomCountFixed + m_zoomCountDynamic
        
        Select Case i
        
            Case 11, 12, 26
                dstDropDown.AddItem m_zoomStrings(i), i, True
                
            Case Else
                dstDropDown.AddItem m_zoomStrings(i), i
        
        End Select
        
    Next i
    
    If (initialListIndex = -1) Then
        dstDropDown.ListIndex = ZOOM_100_PERCENT
    Else
        dstDropDown.ListIndex = initialListIndex
    End If
    
    dstDropDown.SetAutomaticRedraws True, True

End Sub

'Given a current zoom index, find the nearest relevant "zoom in" index.
' (This requires special handling in the case of "fit image on screen".)
Public Function GetNearestZoomInIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-in index is easy.
    If (curIndex <= m_zoomCountFixed) Then
        GetNearestZoomInIndex = curIndex - 1
        If (GetNearestZoomInIndex < 0) Then GetNearestZoomInIndex = 0
    
    'If the current zoom index is one of the "fit" options, this is more complicated.
    ' We want to set the first fixed index we find that is larger than the current dynamic value being used.
    Else
        GetNearestZoomInIndex = GetNearestZoomInIndex_FromRatio(GetZoomRatioFromIndex(curIndex))
    End If

End Function

'Given a current zoom index, find the nearest relevant "zoom out" index.
' (This requires special handling in the case of "fit image on screen".)
Public Function GetNearestZoomOutIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-out index is easy.
    If (curIndex <= m_zoomCountFixed) Then
        GetNearestZoomOutIndex = curIndex + 1
        If (GetNearestZoomOutIndex > m_zoomCountFixed) Then GetNearestZoomOutIndex = m_zoomCountFixed
    
    'If the current zoom index is one of the "fit" options, this is more complicated.
    ' We want to set the first fixed index we find that is smaller than the current dynamic value being used.
    Else
        GetNearestZoomOutIndex = GetNearestZoomOutIndex_FromRatio(GetZoomRatioFromIndex(curIndex))
    End If

End Function

Public Function GetNearestZoomInIndex_FromRatio(ByVal srcRatio As Double) As Long

    'Search the zoom array for the nearest value that is larger than the current zoom value.
    Dim i As Long
    For i = m_zoomCountFixed To 0 Step -1
        If (m_zoomValues(i) > srcRatio) Then
            GetNearestZoomInIndex_FromRatio = i
            Exit For
        End If
    Next i

End Function

Public Function GetNearestZoomOutIndex_FromRatio(ByVal srcRatio As Double) As Long
    
    'Search the zoom array for the nearest value that is less than the current zoom value.
    Dim i As Long
    For i = 0 To m_zoomCountFixed
        If (m_zoomValues(i) < srcRatio) Then
            GetNearestZoomOutIndex_FromRatio = i
            Exit For
        End If
    Next i

End Function
