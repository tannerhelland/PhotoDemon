Attribute VB_Name = "Tool_Support"
'***************************************************************************
'Helper functions for various PhotoDemon tools
'Copyright ©2013-2014 by Tanner Helland
'Created: 06/February/14
'Last updated: 06/February/14
'Last update: initial build
'
'To keep the pdCanvas user control codebase lean, much of its MouseMove events redirect here, to specialized
' functions that take mouse actions on the canvas and translate them into tool actions.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The drag-to-pan tool uses these values to store the original image offset
Private m_InitHScroll As Long, m_InitVScroll As Long

'The drag-to-pan tool uses this function to set the initial scroll bar values for a pan operation
Public Sub setInitialCanvasScrollValues(ByRef srcCanvas As pdCanvas)

    m_InitHScroll = srcCanvas.getScrollValue(PD_HORIZONTAL)
    m_InitVScroll = srcCanvas.getScrollValue(PD_VERTICAL)

End Sub

'The drag-to-pan tool uses this function to actually scroll
Public Sub panImageCanvas(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas)

    'Prevent the canvas from redrawing itself until our pan operation is complete.  (This prevents juddery movement.)
    srcCanvas.setRedrawSuspension True

    'If the canvas in question has a horizontal scrollbar, process it
    If srcCanvas.getScrollVisibility(PD_HORIZONTAL) Then
    
        'Calculate a new scroll value
        Dim hOffset As Long
        hOffset = (initX - curX)
        
        'When zoomed-in, sub-pixel scrolling is not allowed.  Compensate for that now
        If srcImage.currentZoomValue < g_Zoom.getZoom100Index Then
            hOffset = hOffset / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)
        End If
        
        'Factor in the initial scroll bar value
        hOffset = m_InitHScroll + hOffset
        
        'If that value lies within the bounds of the scroll bar, apply it
        If (hOffset < srcCanvas.getScrollMin(PD_HORIZONTAL)) Then
            srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMin(PD_HORIZONTAL)
        ElseIf (hOffset > srcCanvas.getScrollMax(PD_HORIZONTAL)) Then
            srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMax(PD_HORIZONTAL)
        Else
            srcCanvas.setScrollValue PD_HORIZONTAL, hOffset
        End If
    
    End If
    
    'If the canvas in question has a vertical scrollbar, process it
    If srcCanvas.getScrollVisibility(PD_VERTICAL) Then
    
        'Calculate a new scroll value
        Dim vOffset As Long
        vOffset = (initY - curY)
        
        'When zoomed-in, sub-pixel scrolling is not allowed.  Compensate for that now
        If srcImage.currentZoomValue < g_Zoom.getZoom100Index Then
            vOffset = vOffset / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)
        End If
        
        'Factor in the initial scroll bar value
        vOffset = m_InitVScroll + vOffset
        
        'If that value lies within the bounds of the scroll bar, apply it
        If (vOffset < srcCanvas.getScrollMin(PD_VERTICAL)) Then
            srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMin(PD_VERTICAL)
        ElseIf (vOffset > srcCanvas.getScrollMax(PD_VERTICAL)) Then
            srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMax(PD_VERTICAL)
        Else
            srcCanvas.setScrollValue PD_VERTICAL, vOffset
        End If
    
    End If
    
    'Reinstate canvas redraws
    srcCanvas.setRedrawSuspension False
    
    'Manually request a canvas redraw
    ScrollViewport srcImage, srcCanvas

End Sub
