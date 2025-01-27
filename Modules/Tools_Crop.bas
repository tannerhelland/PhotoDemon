Attribute VB_Name = "Tools_Crop"
'***************************************************************************
'Crop tool interface
'Copyright 2024-2025 by Tanner Helland
'Created: 12/November/24
'Last updated: 24/January/25
'Last update: initial build
'
'The crop tool performs identical operations to the rectangular selection tool + Image > Crop menu.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'TRUE if _MouseDown event was received, and _MouseUp has not yet arrived
Private m_LMBDown As Boolean

'Populated in _MouseDown
Private Const INVALID_X_COORD As Double = DOUBLE_MAX, INVALID_Y_COORD As Double = DOUBLE_MAX
Private m_InitImgX As Double, m_InitImgY As Double

'Populate in _MouseMove
Private m_LastImgX As Double, m_LastImgY As Double

'Current crop rect, always *in image coordinates*
Private m_CropRectF As RectF

'Some crop attributes can be independently locked (e.g. aspect ratio).
Private m_IsWidthLocked As Boolean, m_IsHeightLocked As Boolean, m_IsAspectLocked As Boolean
Private m_LockedWidth As Long, m_LockedHeight As Long, m_LockedAspectRatio As Double

'Index of the current mouse_over point, if any.
Private m_idxHover As PD_PointOfInterest

'On _MouseDown events, this will be set to a specific POI if the user is interacting with that POI.
' (If the user is creating a new Crop from scratch, this will be poi_Undefined.)
Private m_idxMouseDown As PD_PointOfInterest

'If the user is modifying an existing crop boundary, this list of coordinates will be filled in _MouseDown
Private m_CornerCoords() As PointFloat, m_numCornerCoords As Long

'To correctly detect clicks, we cache the current crop rect (if any) at _MouseDown
Private m_CropRectAtMouseDown As RectF

'Commit (apply) the current crop rectangle.  *All* layers will be affected by this.
' This function takes the current crop settings, builds a string parameter list from them, then calls
' PD's central processor with the resulting string.  (This allows the crop to be recorded.)
Public Sub CommitCurrentCrop()
    
    Debug.Print "(commit code here)"
    
End Sub

'Turn off the current crop rectangle.  No image modifications will be made.
Public Sub RemoveCurrentCrop()
    ResetCropRectF
    RelayCropChangesToUI
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
End Sub

Public Sub DrawCanvasUI(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage)
    
    'Update the status bar with the curent crop rectangle (if any)
    dstCanvas.SetSelectionState ValidateCropRectF
    
    'Skip if the current rect is invalid
    If (Not ValidateCropRectF()) Then Exit Sub
    
    'Clone a pair of UI pens from the main rendering module.  (Note that we clone them because we need
    ' to modify some rendering properties, and we don't want to fuck up the central cache.)
    Dim basePenInactive As pd2DPen, topPenInactive As pd2DPen
    Dim basePenActive As pd2DPen, topPenActive As pd2DPen
    Drawing.CloneCachedUIPens basePenInactive, topPenInactive, False
    Drawing.CloneCachedUIPens basePenActive, topPenActive, True
    
    'Specify rounded line edges for our pens; this looks better for this particular tool
    basePenInactive.SetPenLineCap P2_LC_Round
    topPenInactive.SetPenLineCap P2_LC_Round
    basePenActive.SetPenLineCap P2_LC_Round
    topPenActive.SetPenLineCap P2_LC_Round
    
    basePenInactive.SetPenLineJoin P2_LJ_Round
    topPenInactive.SetPenLineJoin P2_LJ_Round
    basePenActive.SetPenLineJoin P2_LJ_Round
    topPenActive.SetPenLineJoin P2_LJ_Round
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Convert the crop rectangle to canvas coordinates
    Dim cropRectCanvas As RectF
    Drawing.ConvertImageCoordsToCanvasCoords_RectF dstCanvas, srcImage, m_CropRectF, cropRectCanvas, False
    
    'We now want to add all lines to a path, which we'll render all at once
    Dim cropOutline As pd2DPath
    Set cropOutline = New pd2DPath
    cropOutline.AddRectangle_RectF cropRectCanvas
    
    Dim drawActiveOutline As Boolean
    drawActiveOutline = (m_idxHover = poi_Interior)
    If m_LMBDown Then drawActiveOutline = drawActiveOutline And (m_idxMouseDown = poi_Interior)
    
    'Stroke the outline path
    If drawActiveOutline Then
        PD2D.DrawPath cSurface, basePenActive, cropOutline
        PD2D.DrawPath cSurface, topPenActive, cropOutline
    Else
        PD2D.DrawPath cSurface, basePenInactive, cropOutline
        PD2D.DrawPath cSurface, topPenInactive, cropOutline
    End If
    
    'Next, we want to render drag outlines over the four corner vertices.
    Dim ptCorners() As PointFloat, numPoints As Long
    numPoints = GetCropCorners(ptCorners)
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, ptCorners, False
    
    Dim cornerSize As Single, halfCornerSize As Single
    cornerSize = SQUARE_CORNER_SIZE
    halfCornerSize = cornerSize * 0.5!
    
    Dim idxActive As PD_PointOfInterest
    idxActive = poi_Undefined
    If m_LMBDown Then idxActive = m_idxMouseDown Else idxActive = m_idxHover
    
    Dim i As Long
    For i = 0 To numPoints - 1
        If (i = idxActive) Then
            PD2D.DrawRectangleF cSurface, basePenActive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, topPenActive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
        Else
            PD2D.DrawRectangleF cSurface, basePenInactive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, topPenInactive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
        End If
    Next i
    
    'Free rendering objects
    Set cSurface = Nothing
    Set basePenInactive = Nothing: Set topPenInactive = Nothing
    Set basePenActive = Nothing: Set topPenActive = Nothing
    
End Sub

Public Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Cache initial x/y positions
    m_LMBDown = ((Button And pdLeftButton) = pdLeftButton)
    
    If m_LMBDown Then
        
        'Cache the current crop rect, if any.
        ' (on _Click events, we'll compare click coords against *this* rect.)
        m_CropRectAtMouseDown = m_CropRectF
        
        'Translate the canvas coordinates to image coordinates
        Dim imgX As Double, imgY As Double
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
        'Cache mouse coords so we can calculate translation amounts in future _MouseMove events
        m_InitImgX = imgX
        m_InitImgY = imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
    
        'See if the user is creating a new crop, or interacting with an existing point
        m_idxMouseDown = UpdateMousePOI(imgX, imgY)
        
        'If the user is initiating a new crop, start it now
        If (m_idxMouseDown = poi_Undefined) Then
            ResetCropRectF imgX, imgY
            
        'The user is interacting with an existing crop boundary.
        ' How we modify the crop rect depends on the point being interacted with.
        Else
            
            'Cache a full list of boundary coordinates
            m_numCornerCoords = GetCropCorners(m_CornerCoords)
            
        End If
            
    Else
        m_InitImgX = INVALID_X_COORD
        m_InitImgY = INVALID_Y_COORD
    End If
    
End Sub

Public Sub NotifyMouseMove(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
    
    'Cache current x/y positions
    If m_LMBDown Then
        
        'If the user is interacting with an existing crop, modify the crop rect accordingly.
        If (m_idxMouseDown <> poi_Undefined) Then
            
            'Calculate an offset from the original click to this new location
            Dim xOffset As Long, yOffset As Long
            xOffset = imgX - m_InitImgX
            yOffset = imgY - m_InitImgY
            
            Dim i As Long, tmpCornerCoords() As PointFloat
            If (m_numCornerCoords > 0) Then ReDim tmpCornerCoords(0 To UBound(m_CornerCoords)) As PointFloat
            
            'On a "move" event, the entire crop area is moving.  Calculate an offset and move all crop points that far.
            If (m_idxMouseDown = poi_Interior) Then
                
                'Failsafe only; _MouseDown will always set m_numCornerCoords to 4
                If (m_numCornerCoords > 0) Then
                    For i = 0 To m_numCornerCoords - 1
                        tmpCornerCoords(i).x = m_CornerCoords(i).x + xOffset
                        tmpCornerCoords(i).y = m_CornerCoords(i).y + yOffset
                    Next i
                End If
                
            'The user is click-dragging a specific point.
            Else
                
                'Failsafe only; _MouseDown will always set m_numCornerCoords to 4
                If (m_numCornerCoords > 0) Then
                    
                    For i = 0 To m_numCornerCoords - 1
                        tmpCornerCoords(i).x = m_CornerCoords(i).x
                        tmpCornerCoords(i).y = m_CornerCoords(i).y
                        If (i = m_idxMouseDown) Then
                            tmpCornerCoords(i).x = tmpCornerCoords(i).x + xOffset
                            tmpCornerCoords(i).y = tmpCornerCoords(i).y + yOffset
                        End If
                    Next i
                    
                    'Because the point-list-to-rect function operates on max/min values, we need to adjust adjoining
                    ' corners too.
                    If (m_idxMouseDown = 0) Then
                        tmpCornerCoords(1).y = tmpCornerCoords(1).y + yOffset
                        tmpCornerCoords(2).x = tmpCornerCoords(2).x + xOffset
                    ElseIf (m_idxMouseDown = 1) Then
                        tmpCornerCoords(0).y = tmpCornerCoords(0).y + yOffset
                        tmpCornerCoords(3).x = tmpCornerCoords(3).x + xOffset
                    ElseIf (m_idxMouseDown = 2) Then
                        tmpCornerCoords(0).x = tmpCornerCoords(0).x + xOffset
                        tmpCornerCoords(3).y = tmpCornerCoords(3).y + yOffset
                    Else
                        tmpCornerCoords(1).x = tmpCornerCoords(1).x + xOffset
                        tmpCornerCoords(2).y = tmpCornerCoords(2).y + yOffset
                    End If
                    
                End If
                
            End If
            
            'Update the crop rect to reflect any changes made to individual coordinates
            UpdateCropRectF_FromPtFList tmpCornerCoords, (m_idxMouseDown = poi_Interior), m_idxMouseDown
            
        '...otherwise, simply create a new crop to match the mouse movement
        Else
            
            'Move coordinates around to ensure positive width/height
            UpdateCropRectF imgX, imgY
            m_LastImgX = imgX
            m_LastImgY = imgY
            
        End If
        
    '/LMB is *not* down
    End If
    
    'Update the currently hovered crop corner, if any
    m_idxHover = UpdateMousePOI(imgX, imgY)
    
    'Finally, request a viewport redraw (to reflect any changes)
    Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Update cached button status now, before handling the actual event.  (This ensures that
    ' a viewport redraw, if any, will wipe all UI elements we may have rendered over the canvas.)
    If ((Button And pdLeftButton) = pdLeftButton) Then m_LMBDown = False
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
    'In Photoshop, double-click commits the crop.  In GIMP, it's single-click.
    '
    'I'm still debating which way to go in PD, but I'm currently leaning toward GIMP as it's also the
    ' traditional PD convention (click-outside-to-cancel, click-inside-to-apply).
    If clickEventAlsoFiring Then
        
        'Check mouse position.  If the mouse is *outside* the crop rect that existed at _MouseDown,
        ' simply clear the current rect.
        If (Not PDMath.IsPointInRectF(imgX, imgY, m_CropRectAtMouseDown)) Then
            RemoveCurrentCrop
        Else
            CommitCurrentCrop
        End If
        
    'If this is a click-drag event, we need to simply resize the crop area to match.
    ' (Commit happens in a separate step.)
    Else
        
        'If the user has modified the current crop in an unuseable way, we need to update the UI accordingly
        ' (by e.g. disabling the "commit crop" button)
        Dim cropAintGood As Boolean
        cropAintGood = False
            
        'If the user is interacting with an existing crop, modify the crop rect accordingly.
        If (m_idxMouseDown <> poi_Undefined) Then
        
        '...otherwise, simply create a new crop to match the mouse movement
        Else
                
            'Update the crop rectangle against these (final) coordinates
            UpdateCropRectF imgX, imgY
            m_LastImgX = imgX
            m_LastImgY = imgY
            
        End If
        
        'Bail if any of the final coordinates produce an unuseable crop rect
        If (m_InitImgX = INVALID_X_COORD) Or (m_InitImgY = INVALID_Y_COORD) Then cropAintGood = True
        If (Not ValidateCropRectF()) Then cropAintGood = True
        If cropAintGood Then
            RemoveCurrentCrop
        
        'Finally, request a viewport redraw (to reflect any changes)
        Else
            
            'If the UI looks OK, update the viewport to match
            Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
            
        End If
        
    End If
    
End Sub

'Only returns usable data when IsValidCropActive() is TRUE
Public Function GetCropRectF() As RectF
    GetCropRectF = m_CropRectF
End Function

'Called whenever the crop tool is selected
Public Sub InitializeCropTool()
    ResetCropRectF
    RelayCropChangesToUI
End Sub

'Check before using anything from GetCropRectF, above
Public Function IsValidCropActive() As Boolean
    IsValidCropActive = ValidateCropRectF
End Function

'Some properties can be independently locked (e.g. width or height or aspect ratio).
' When a property is locked, it cannot be changed by additional UI inputs.
Public Sub LockProperty(ByVal selProperty As PD_SelectionLockable, ByVal lockedValue As Variant)

    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = True
        m_LockedWidth = lockedValue
        m_IsHeightLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = True
        m_LockedHeight = lockedValue
        m_IsWidthLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = True
        m_LockedAspectRatio = lockedValue
        m_IsWidthLocked = False
        m_IsHeightLocked = False
    End If
    
End Sub

Public Sub UnlockProperty(ByVal selProperty As PD_SelectionLockable)
    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = False
    End If
End Sub

Public Sub ReadyForCursor(ByRef srcCanvasView As pdCanvasView)
    
    Dim idxRelevant As PD_PointOfInterest
    If m_LMBDown Then idxRelevant = m_idxMouseDown Else idxRelevant = m_idxHover
    
    Select Case idxRelevant
    
        'Mouse not over the crop
        Case poi_Undefined
            srcCanvasView.RequestCursor_System IDC_ARROW
            
        'Mouse is within the crop region, but not over a corner node
        Case poi_Interior
            If m_LMBDown Then
                If (m_idxMouseDown = poi_Interior) Then
                    srcCanvasView.RequestCursor_System IDC_SIZEALL
                Else
                    srcCanvasView.RequestCursor_System IDC_ARROW
                End If
            Else
                srcCanvasView.RequestCursor_System IDC_SIZEALL
            End If
            
        'Mouse is over one of the four crop corners
        Case 0
            srcCanvasView.RequestCursor_System IDC_SIZENWSE
        Case 1
            srcCanvasView.RequestCursor_System IDC_SIZENESW
        Case 2
            srcCanvasView.RequestCursor_System IDC_SIZENESW
        Case 3
            srcCanvasView.RequestCursor_System IDC_SIZENWSE
        
        'Should never trigger; failsafe only
        Case Else
            srcCanvasView.RequestCursor_System IDC_ARROW
        
    End Select
    
End Sub

'On _MouseMove events, update the current POI (if any)
Private Function UpdateMousePOI(ByVal imgX As Long, ByVal imgY As Long) As PD_PointOfInterest
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy.)
    ' (TODO: come up with a better solution for this.  Accuracy should *really* be handled in the canvas coordinate space,
    '        so perhaps the caller should specify an image x/y and a radius...?)
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), PDImages.GetActiveImage)
    
    'Find the smallest distance for this mouse position
    Dim minDistance As Single
    minDistance = mouseAccuracy
    
    'Retrieve current crop corners (in image space)
    Dim cropCorners() As PointFloat, numPoints As Long
    numPoints = GetCropCorners(cropCorners)
    
    'Find the nearest point (if any) to the mouse pointer
    UpdateMousePOI = poi_Undefined
    UpdateMousePOI = PDMath.FindClosestPointInFloatArray(imgX, imgY, minDistance, cropCorners)
    
    'If the mouse is not near a corner, perform an additional check for the crop rect's interior
    If (UpdateMousePOI = poi_Undefined) Then
        
        Dim tmpPath As pd2DPath
        Set tmpPath = New pd2DPath
        tmpPath.AddRectangle_RectF m_CropRectF
        
        If tmpPath.IsPointInsidePathF(imgX, imgY) Then
            UpdateMousePOI = poi_Interior
        Else
            UpdateMousePOI = poi_Undefined
        End If
        
    End If
    
End Function

'Get the current crop rect coordinates as a list of points (in image coordinates).
' Returns: number of points in the array, and a guaranteed ReDim to [0, [n] - 1].
Private Function GetCropCorners(ByRef dstCropPts() As PointFloat) As Long
    
    GetCropCorners = 4
    
    ReDim dstCropPts(0 To GetCropCorners - 1) As PointFloat
    dstCropPts(0).x = m_CropRectF.Left
    dstCropPts(0).y = m_CropRectF.Top
    dstCropPts(1).x = m_CropRectF.Left + m_CropRectF.Width
    dstCropPts(1).y = dstCropPts(0).y
    dstCropPts(2).x = dstCropPts(0).x
    dstCropPts(2).y = m_CropRectF.Top + m_CropRectF.Height
    dstCropPts(3).x = dstCropPts(1).x
    dstCropPts(3).y = dstCropPts(2).y
    
End Function

Private Sub UpdateCropRectF(ByVal newX As Single, ByVal newY As Single)
    
    Dim newHeight As Single
    If m_IsHeightLocked Then
        newHeight = m_LockedHeight
    Else
        newHeight = Abs(m_InitImgY - newY)
        If (newHeight < 1!) Then newHeight = 1!
    End If
    
    Dim newWidth As Single
    If m_IsWidthLocked Then
        newWidth = m_LockedWidth
    Else
        If m_IsAspectLocked Then
            newWidth = newHeight * m_LockedAspectRatio
        Else
            newWidth = Abs(m_InitImgX - newX)
        End If
        If (newWidth < 1!) Then newWidth = 1!
    End If
    
    With m_CropRectF
        
        .Width = newWidth
        
        If (newX < m_InitImgX) Then
            If (m_IsWidthLocked Or m_IsAspectLocked) Then
                .Left = m_InitImgX - newWidth
            Else
                .Left = newX
            End If
        Else
            .Left = m_InitImgX
        End If
        
        .Height = newHeight
        
        If (newY < m_InitImgY) Then
            If (m_IsHeightLocked Or m_IsAspectLocked) Then
                .Top = m_InitImgY - newHeight
            Else
                .Top = newY
            End If
        Else
            .Top = m_InitImgY
        End If
        
    End With
    
    'Lock position and height to their nearest integer equivalent
    PDMath.GetIntClampedRectF m_CropRectF
    
    'If the crop rect is valid, relay its values to the toolpanel
    If ValidateCropRectF() Then RelayCropChangesToUI
    
End Sub

'You can only pass a PointFloat array sized [0, 3] to this function.  It will use that array to produce an updated
' RectF of the boundary coords of the passed list.
Private Sub UpdateCropRectF_FromPtFList(ByRef srcPoints() As PointFloat, Optional ByVal okToMove As Boolean = False, Optional ByVal srcPOI As PD_PointOfInterest = poi_Undefined)
    
    'Find the min/max points in the source point list
    Dim xMin As Single, yMin As Single, xMax As Single, yMax As Single
    xMin = srcPoints(0).x
    yMin = srcPoints(0).y
    xMax = srcPoints(0).x
    yMax = srcPoints(0).y
    
    Dim i As Long
    For i = 1 To UBound(srcPoints)
        If (srcPoints(i).x < xMin) Then xMin = srcPoints(i).x
        If (srcPoints(i).x > xMax) Then xMax = srcPoints(i).x
        If (srcPoints(i).y < yMin) Then yMin = srcPoints(i).y
        If (srcPoints(i).y > yMax) Then yMax = srcPoints(i).y
    Next i
    
    'We can now use the calculated max/min values to calculate a boundary rect (but note that we must
    ' also consider any locked dimensions and/or aspect ratio).
    If m_IsAspectLocked And (Not okToMove) Then
        
        'When aspect ratio is locked, we need to manually calculate new width/height values
        Dim newWidth As Single, newHeight As Single
        newHeight = yMax - yMin
        If (newHeight < 1!) Then newHeight = 1!
        newWidth = Int(newHeight * m_LockedAspectRatio + 0.5!)
        If (newWidth < 1!) Then newWidth = 1!
        m_CropRectF.Width = newWidth
        m_CropRectF.Height = newHeight
        
        'We now need to position the selection to either the left or right, depending on the current POI
        If (srcPOI = 0) Then
            If (srcPoints(0).x < srcPoints(1).x) Then
                m_CropRectF.Left = srcPoints(1).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(0).y < srcPoints(2).y) Then
                m_CropRectF.Top = srcPoints(2).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 1) Then
            If (srcPoints(1).x < srcPoints(0).x) Then
                m_CropRectF.Left = srcPoints(0).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(0).y < srcPoints(3).y) Then
                m_CropRectF.Top = srcPoints(3).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 2) Then
            If (srcPoints(2).x < srcPoints(3).x) Then
                m_CropRectF.Left = srcPoints(3).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(2).y < srcPoints(0).y) Then
                m_CropRectF.Top = srcPoints(0).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 3) Then
            If (srcPoints(3).x < srcPoints(2).x) Then
                m_CropRectF.Left = srcPoints(2).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(3).y < srcPoints(1).y) Then
                m_CropRectF.Top = srcPoints(1).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        End If
    
    'When aspect ratio is not locked, this step is easy: just use max/min values as calculated above
    Else
        
        With m_CropRectF
            If okToMove Or (Not m_IsWidthLocked) Then .Left = xMin
            If (Not m_IsWidthLocked) Then .Width = xMax - xMin
            If okToMove Or (Not m_IsHeightLocked) Then .Top = yMin
            If (Not m_IsHeightLocked) Then .Height = yMax - yMin
        End With
        
    End If
        
    'Lock position and height to their nearest integer equivalent
    PDMath.GetIntClampedRectF m_CropRectF
    
    'If the crop rect is valid, relay its values to the toolpanel
    If ValidateCropRectF() Then RelayCropChangesToUI
    
End Sub

Private Sub ResetCropRectF(Optional ByVal initX As Single = 0!, Optional ByVal initY As Single = 0!)
    With m_CropRectF
        .Left = initX
        .Top = initY
        .Width = 0!
        .Height = 0!
    End With
End Sub

'Returns TRUE if the current crop rect is valid; if FALSE, do *not* attempt to crop
Private Function ValidateCropRectF() As Boolean
    
    ValidateCropRectF = True
    
    'Validate left/right/width/height
    With m_CropRectF
        If (.Left = INVALID_X_COORD) Or (.Top = INVALID_Y_COORD) Then ValidateCropRectF = False
        If (.Width <= 0) Or (.Height <= 0) Then ValidateCropRectF = False
    End With
    
    'Next, ensure the crop rect actually overlaps the image *somewhere*.
    If PDImages.IsImageActive() Then
        Dim dummyIntersectRectF As RectF
        If (Not GDI_Plus.IntersectRectF(dummyIntersectRectF, m_CropRectF, PDImages.GetActiveImage.GetBoundaryRectF)) Then ValidateCropRectF = False
    End If
    
End Function

Private Sub RelayCropChangesToUI()
    
    'Lock updates to prevent circular references
    Tools.SetToolBusyState True
    
    toolpanel_Crop.tudCrop(0).Value = m_CropRectF.Left
    toolpanel_Crop.tudCrop(1).Value = m_CropRectF.Top
    
    toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
    toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
    
    'Aspect ratio requires a bit of math
    Dim fracNumerator As Long, fracDenominator As Long
    If (m_CropRectF.Width > 0!) And (m_CropRectF.Height > 0!) Then
        PDMath.ConvertToFraction m_CropRectF.Width / m_CropRectF.Height, fracNumerator, fracDenominator, 0.005
    Else
        fracNumerator = 1
        fracDenominator = 1
    End If
    
    'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
    If (fracDenominator = 5) Then
        fracNumerator = fracNumerator * 2
        fracDenominator = fracDenominator * 2
    End If
    
    toolpanel_Crop.tudCrop(4).Value = fracNumerator
    toolpanel_Crop.tudCrop(5).Value = fracDenominator
    
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    'Unlock updates
    Tools.SetToolBusyState False
    
End Sub

'The crop toolpanel relays user input via this function.  Left/top/width/height as passed as integers; aspect ratio as a float.
Public Sub RelayCropChangesFromUI(ByVal changedProperty As PD_Dimension, Optional ByVal newPropI As Long = 0, Optional ByVal newPropF As Single = 0!)
    
    Select Case changedProperty
        
        Case pdd_Left
            m_CropRectF.Left = newPropI
            
        Case pdd_Top
            m_CropRectF.Top = newPropI
        
        'Changing width/height requires us to update aspect ratio to match
        Case pdd_Width, pdd_Height
            
            'If the current aspect ratio is locked, we need to update the dimension passed to its exact value,
            ' then modify the opposite dimension to match the locked aspect ratio.
            If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                
                'User is changing width
                If (changedProperty = pdd_Width) Then
                    If toolpanel_Crop.tudCrop(2).IsValid(False) Then
                        m_CropRectF.Width = newPropI
                        m_CropRectF.Height = m_CropRectF.Width * (1# / m_LockedAspectRatio)
                        toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
                    End If
                    
                'User is changing height
                Else
                    If toolpanel_Crop.tudCrop(3).IsValid(False) Then
                        m_CropRectF.Height = newPropI
                        m_CropRectF.Width = m_CropRectF.Height * m_LockedAspectRatio
                        toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
                    End If
                End If
            
            'If aspect ratio is *not* locked, freely modify either value, and cache any locked values
            ' so they can be used elsewhere.
            Else
                
                If (changedProperty = pdd_Width) Then
                    m_CropRectF.Width = newPropI
                    If m_IsWidthLocked Then m_LockedWidth = m_CropRectF.Width
                Else
                    m_CropRectF.Height = newPropI
                    If m_IsHeightLocked Then m_LockedHeight = m_CropRectF.Height
                End If
                
                'Changes to width/height (with unlocked aspect ratio) will obviously change the
                ' current aspect ratio.  Re-calculate aspect ratio and update those spinners accordingly.
                Dim fracNumerator As Long, fracDenominator As Long
                If (m_CropRectF.Width >= 1) And (m_CropRectF.Height >= 1) Then
                    
                    PDMath.ConvertToFraction m_CropRectF.Width / m_CropRectF.Height, fracNumerator, fracDenominator, 0.005
                    
                    'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
                    If (fracDenominator = 5) Then
                        fracNumerator = fracNumerator * 2
                        fracDenominator = fracDenominator * 2
                    End If
                    
                    toolpanel_Crop.tudCrop(4).Value = fracNumerator
                    toolpanel_Crop.tudCrop(5).Value = fracDenominator
                    
                End If
                    
            End If
        
        'When changing aspect ratio, the UI passes both the new aspect ratio, and the current size of
        ' the *opposite* dimension.
        Case pdd_AspectRatioW
            m_CropRectF.Width = newPropI
            toolpanel_Crop.tudCrop(2).Value = newPropI
            If m_IsAspectLocked Then m_LockedAspectRatio = newPropF
            
        Case pdd_AspectRatioH
            m_CropRectF.Height = newPropI
            toolpanel_Crop.tudCrop(3).Value = newPropI
            If m_IsAspectLocked And (newPropF <> 0!) Then m_LockedAspectRatio = 1! / newPropF
            
    End Select
    
    'After a property change, we must validate the crop rectangle to make sure no funny business occurred
    ValidateCropRectF
    
    'We don't want to relay *all* crop settings back to the UI (because this function is meant to
    ' relay values *from* the UI), but we may need to adjust the apply/commit buttons if the user's
    ' changes changed crop validity state.
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    '...then request a viewport redraw to reflect any changes
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub
