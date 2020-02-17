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
'Copyright 2001-2020 by Tanner Helland
'Created: 4/15/01
'Last updated: 28/August/17
'Last update: code clean-up and minor optimizations
'
'The main user of this class is the Viewport_Handler module.  Look there for relevant implementation details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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
Friend Function GetZoomValue(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom value at that index
    If (zoomIndex <= m_zoomCountFixed) Then
        GetZoomValue = m_zoomValues(zoomIndex)
        
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
                            GetZoomValue = viewportWidth / imgWidth
                        Else
                            GetZoomValue = horizontalRatio
                        End If
                        
                    'Fit height
                    Case m_zoomCountFixed + 2
                    
                        'Check to see if the calculated zoom value will require a horizontal scroll bar (since we are only fitting height).
                        ' If it will, we must subtract the scroll bar's height from our calculation.
                        If (imgWidth * verticalRatio > viewportWidth) Then
                            GetZoomValue = viewportHeight / imgHeight
                        Else
                            GetZoomValue = verticalRatio
                        End If
                        
                    'Fit everything
                    Case m_zoomCountFixed + 3
                        If (horizontalRatio < verticalRatio) Then
                            GetZoomValue = horizontalRatio
                        Else
                            GetZoomValue = verticalRatio
                        End If
                
                End Select
                
            Else
                GetZoomValue = 1#
            End If
            
        Else
            GetZoomValue = 1#
        End If
        
    
    End If
    
End Function

Friend Function GetZoomOffsetFactor(ByVal zoomIndex As Long) As Double
    
    'If the zoom value is a fixed entry, our work is easy - simply return the fixed zoom offset at that index
    If (zoomIndex <= m_zoomCountFixed) Then
        GetZoomOffsetFactor = m_zoomOffsetFactors(zoomIndex)
    
    'If the zoom value is a dynamic entry, we need to calculate a specific zoom offset at run-time
    Else
    
        Dim curZoomValue As Double
        curZoomValue = GetZoomValue(zoomIndex)
        
        If (curZoomValue >= 1#) Then
            GetZoomOffsetFactor = curZoomValue
        Else
            GetZoomOffsetFactor = 1# / curZoomValue
        End If
    
    End If
    
End Function

'To minimize the possibility of program-wide changes if I ever decide to fiddle with PD's fixed zoom values, these functions are used
' externally to retrieve specific zoom indices.
Friend Function GetZoom100Index() As Long
    GetZoom100Index = ZOOM_100_PERCENT
End Function

Friend Function GetZoomFitWidthIndex() As Long
    GetZoomFitWidthIndex = m_zoomCountFixed + 1
End Function

Friend Function GetZoomFitHeightIndex() As Long
    GetZoomFitHeightIndex = m_zoomCountFixed + 2
End Function

Friend Function GetZoomFitAllIndex() As Long
    GetZoomFitAllIndex = m_zoomCountFixed + 3
End Function

Friend Function GetZoomCount() As Long
    GetZoomCount = m_zoomCountFixed
End Function

'Whenever one of these classes is created, remember to call this initialization function.  It will manually prepare a
' list of zoom values relevant to the program.
Friend Sub InitializeViewportEngine()

    'This list of zoom values is (effectively) arbitrary.  I've based this list off similar lists (Paint.NET, GIMP)
    ' while including a few extra values for convenience's sake
    
    'Total number of fixed zoom values.  Some legacy PD functions (like the old Fit to Screen code) require this so
    ' they can iterate all fixed zoom values, and find an appropriate one for their purpose.
    m_zoomCountFixed = 25
    
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
        
    m_zoomStrings(11) = "100%"
        m_zoomValues(11) = 1
        m_zoomOffsetFactors(11) = 1
        
    m_zoomStrings(12) = "75%"
        m_zoomValues(12) = 3# / 4#
        m_zoomOffsetFactors(12) = 4# / 3#
        
    m_zoomStrings(13) = "67%"
        m_zoomValues(13) = 2# / 3#
        m_zoomOffsetFactors(13) = 3# / 2#
        
    m_zoomStrings(14) = "50%"
        m_zoomValues(14) = 0.5
        m_zoomOffsetFactors(14) = 2#
        
    m_zoomStrings(15) = "33%"
        m_zoomValues(15) = 1# / 3#
        m_zoomOffsetFactors(15) = 3
        
    m_zoomStrings(16) = "25%"
        m_zoomValues(16) = 0.25
        m_zoomOffsetFactors(16) = 4
        
    m_zoomStrings(17) = "20%"
        m_zoomValues(17) = 0.2
        m_zoomOffsetFactors(17) = 5
        
    m_zoomStrings(18) = "16%"
        m_zoomValues(18) = 0.16
        m_zoomOffsetFactors(18) = 100# / 16#
        
    m_zoomStrings(19) = "12%"
        m_zoomValues(19) = 0.12
        m_zoomOffsetFactors(19) = 100# / 12#
        
    m_zoomStrings(20) = "8%"
        m_zoomValues(20) = 0.08
        m_zoomOffsetFactors(20) = 100# / 8#
        
    m_zoomStrings(21) = "6%"
        m_zoomValues(21) = 0.06
        m_zoomOffsetFactors(21) = 100# / 6#
        
    m_zoomStrings(22) = "4%"
        m_zoomValues(22) = 0.04
        m_zoomOffsetFactors(22) = 25
        
    m_zoomStrings(23) = "3%"
        m_zoomValues(23) = 0.03
        m_zoomOffsetFactors(23) = 100# / 0.03
        
    m_zoomStrings(24) = "2%"
        m_zoomValues(24) = 0.02
        m_zoomOffsetFactors(24) = 50
        
    m_zoomStrings(25) = "1%"
        m_zoomValues(25) = 0.01
        m_zoomOffsetFactors(25) = 100
    
    m_zoomStrings(26) = g_Language.TranslateMessage("Fit width")
        m_zoomValues(26) = 0
        m_zoomOffsetFactors(26) = 0
    
    m_zoomStrings(27) = g_Language.TranslateMessage("Fit height")
        m_zoomValues(27) = 0
        m_zoomOffsetFactors(27) = 0
        
    m_zoomStrings(28) = g_Language.TranslateMessage("Fit image")
        m_zoomValues(28) = 0
        m_zoomOffsetFactors(28) = 0
    
    'Note which index corresponds to 100%
    ZOOM_100_PERCENT = 11
    
End Sub

'Populate an arbitrary combo box with the current list of handled zoom values
Friend Sub PopulateZoomComboBox(ByRef dstComboBox As pdDropDown, Optional ByVal initialListIndex As Long = -1)
    
    dstComboBox.SetAutomaticRedraws False
    
    dstComboBox.Clear
    
    Dim i As Long
    
    For i = 0 To m_zoomCountFixed + m_zoomCountDynamic
        
        Select Case i
        
            Case 10, 11, 25
                dstComboBox.AddItem m_zoomStrings(i), i, True
                
            Case Else
                dstComboBox.AddItem m_zoomStrings(i), i
        
        End Select
        
    Next i
    
    If (initialListIndex = -1) Then
        dstComboBox.ListIndex = ZOOM_100_PERCENT
    Else
        dstComboBox.ListIndex = initialListIndex
    End If
    
    dstComboBox.SetAutomaticRedraws True, True

End Sub

'Given a current zoom index, find the nearest relevant "zoom in" index.  This requires special handling in the case of "fit image on screen".
Friend Function GetNearestZoomInIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-in index is easy.
    If (curIndex <= m_zoomCountFixed) Then
        
        GetNearestZoomInIndex = curIndex - 1
        If (GetNearestZoomInIndex < 0) Then GetNearestZoomInIndex = 0
    
    'If the current zoom index is one of the "fit" options, this is more complicated.  We want to set the first fixed index we
    ' find that is larger than the current dynamic value being used.
    Else
    
        Dim curZoomValue As Double
        curZoomValue = GetZoomValue(curIndex)
        
        'Start searching the zoom array for the nearest value that is larger than the current zoom value.
        Dim i As Long
        For i = m_zoomCountFixed To 0 Step -1
            If (m_zoomValues(i) > curZoomValue) Then
                GetNearestZoomInIndex = i
                Exit For
            End If
        Next i
    
    End If

End Function

'Given a current zoom index, find the nearest relevant "zoom out" index.  This requires special handling in the case of "fit image on screen".
Friend Function GetNearestZoomOutIndex(ByVal curIndex As Long) As Long

    'This function is split into two cases.  If the current zoom index is a fixed value (e.g. "100%"), finding
    ' the nearest zoom-out index is easy.
    If curIndex <= m_zoomCountFixed Then
        
        GetNearestZoomOutIndex = curIndex + 1
        If GetNearestZoomOutIndex > m_zoomCountFixed Then GetNearestZoomOutIndex = m_zoomCountFixed
    
    'If the current zoom index is one of the "fit" options, this is more complicated.  We want to set the first fixed index we
    ' find that is smaller than the current dynamic value being used.
    Else
    
        Dim curZoomValue As Double
        curZoomValue = GetZoomValue(curIndex)
        
        'Start searching the zoom array for the nearest value that is larger than the current zoom value.
        Dim i As Long
        For i = 0 To m_zoomCountFixed
            If m_zoomValues(i) < curZoomValue Then
                GetNearestZoomOutIndex = i
                Exit For
            End If
        Next i
    
    End If

End Function
