Attribute VB_Name = "Tools_Zoom"
'***************************************************************************
'Zoom on-canvas tool interface
'Copyright 2021-2021 by Tanner Helland
'Created: 14/December/21
'Last updated: 14/December/21
'Last update: start migrating zoom bits from elsewhere into this dedicated module
'
'PD's Zoom tool is very straightforward.  It basically relays simple zoom commands from the canvas
' to PD's viewport engine.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Left-click zooms in, right-click zooms out (per convention with other software)
    If clickEventAlsoFiring Then
        
        Dim zoomIn As Boolean
        If ((Button And pdLeftButton) <> 0) Then
            zoomIn = True
        ElseIf ((Button And pdRightButton) <> 0) Then
            zoomIn = False
        Else
            Exit Sub
        End If
        
        Tools_Zoom.RelayCanvasZoom srcCanvas, srcImage, canvasX, canvasY, zoomIn
    
    'TODO
    Else
    
    End If
        
End Sub

'When a canvas event initiates zoom (mousewheel, zoom tool, etc), send the relevant canvas info here
' and this function will perform the actual zoom change.  It will return TRUE if the caller needs to
Public Sub RelayCanvasZoom(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Double, ByVal canvasY As Double, ByVal zoomIn As Boolean)

    If (Not srcCanvas.IsCanvasInteractionAllowed()) Then Exit Sub
    
    'Before doing anything else, cache the current mouse coordinates (in both Canvas and Image coordinate spaces)
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, True
    
    'Suspend automatic viewport redraws until we are done with our calculations.
    ' (Same goes for the canvas, which needs to stop handling scroll bar synchronization until we're done.)
    Viewport.DisableRendering
    srcCanvas.SetRedrawSuspension True
    
    'Calculate a new zoom value
    If srcCanvas.IsZoomEnabled() Then
        If zoomIn Then
            If (srcCanvas.GetZoomDropDownIndex > 0) Then srcCanvas.SetZoomDropDownIndex Zoom.GetNearestZoomInIndex(srcCanvas.GetZoomDropDownIndex)
        Else
            If (srcCanvas.GetZoomDropDownIndex <> Zoom.GetZoomCount) Then srcCanvas.SetZoomDropDownIndex Zoom.GetNearestZoomOutIndex(srcCanvas.GetZoomDropDownIndex)
        End If
    End If
    
    'Relay the new zoom value to the target pdImage object (pdImage objects store their current zoom value,
    ' so we can preserve it when switching between images)
    srcImage.SetZoomIndex srcCanvas.GetZoomDropDownIndex()
    
    'Re-enable automatic viewport redraws
    Viewport.EnableRendering
    srcCanvas.SetRedrawSuspension False
    
    'Request a manual redraw from Viewport.Stage1_InitializeBuffer, while supplying our x/y coordinates so that
    ' it can preserve mouse position relative to the underlying image.
    Viewport.Stage1_InitializeBuffer srcImage, srcCanvas, VSR_PreservePointPosition, canvasX, canvasY, imgX, imgY
    
    'Notify external UI elements of the change
    Viewport.NotifyEveryoneOfViewportChanges

End Sub
