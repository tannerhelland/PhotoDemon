Attribute VB_Name = "Drawing"
'***************************************************************************
'PhotoDemon Drawing Routines
'Copyright 2001-2026 by Tanner Helland
'Created: 4/3/01
'Last updated: 28/May/24
'Last update: new "show distances" options for Move tool will render pixel distances onto the primary canvas
'
'Miscellaneous drawing routines that don't fit elsewhere.  At present, this includes rendering preview images,
' drawing the canvas background of image forms, and a gradient-rendering sub (used primarily on the histogram form).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The following Enum defines standard message box information icons, but note that PD does *not* use
' the system versions.  (Instead, to support run-time theming we paint our own copies.)
Public Enum SystemIconConstants
    IDI_HAND = 32513
    IDI_QUESTION = 32514
    IDI_EXCLAMATION = 32515
    IDI_ASTERISK = 32516
    IDI_WINDOWS = 32517
End Enum

#If False Then
    Private Const IDI_HAND = 32513, IDI_QUESTION = 32514, IDI_EXCLAMATION = 32515, IDI_ASTERISK = 32516, IDI_WINDOWS = 32517
#End If

Public Enum PD_ShowTargets
    pdst_Grid
    pdst_Guides
    pdst_LayerEdges
    pdst_Slices
    pdst_SmartGuides
End Enum

#If False Then
    Private Const pdst_Grid = 0, pdst_Guides = 0, pdst_LayerEdges = 0, pdst_Slices = 0, pdst_SmartGuides = 0
#End If

'At startup, PD caches a few different UI pens and brushes related to viewport rendering.
Private m_PenUIBase As pd2DPen, m_PenUITop As pd2DPen
Private m_PenUIBaseHighlight As pd2DPen, m_PenUITopHighlight As pd2DPen

'For performance reasons, some other recurring rendering bits are cached.
Private m_ShowLayerEdges As Boolean, m_ShowSmartGuides As Boolean

'Draw a horizontal gradient to a specified DIB from x-position xLeft to xRight.
Public Sub DrawHorizontalGradientToDIB(ByVal dstDIB As pdDIB, ByVal xLeft As Single, ByVal xRight As Single, ByVal colorLeft As Long, ByVal colorRight As Long)
    
    Dim boundsRectF As RectF
    With boundsRectF
        .Left = (xLeft - 1)
        .Width = (xRight - xLeft) + 2
        .Top = 0
        .Height = dstDIB.GetDIBHeight
    End With
    
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstDIB.GetDIBDC, False
    Drawing2D.QuickCreateTwoColorGradientBrush cBrush, boundsRectF, colorLeft, colorRight
    cBrush.SetBrushGradientWrapMode P2_WM_Clamp
    
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, boundsRectF
    
End Sub

'Given a source DIB, fill it with a 2x2 alpha checkerboard pattern matching the user's current preferences.
' (The resulting DIB size is contingent on the user's checkerboard pattern size preference, FYI.)
Public Sub CreateAlphaCheckerboardDIB(ByRef srcDIB As pdDIB)

    'Retrieve the user's preferred alpha checkerboard colors, and convert the longs into individual RGB components
    Dim chkColorOne As Long, chkColorTwo As Long
    chkColorOne = UserPrefs.GetPref_Long("Transparency", "AlphaCheckOne", RGB(255, 255, 255))
    chkColorTwo = UserPrefs.GetPref_Long("Transparency", "AlphaCheckTwo", RGB(204, 204, 204))
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = Colors.ExtractRed(chkColorOne)
    r2 = Colors.ExtractRed(chkColorTwo)
    g1 = Colors.ExtractGreen(chkColorOne)
    g2 = Colors.ExtractGreen(chkColorTwo)
    b1 = Colors.ExtractBlue(chkColorOne)
    b2 = Colors.ExtractBlue(chkColorTwo)
    
    'Determine a checkerboard block size based on the current user preference
    Dim chkSize As Long
    chkSize = UserPrefs.GetPref_Long("Transparency", "AlphaCheckSize", 1)
    
    Select Case chkSize
    
        'Small (4x4 checks)
        Case 0
            chkSize = 4
            
        'Medium (8x8 checks)
        Case 1
            chkSize = 8
        
        'Large (16x16 checks)
        Case Else
            chkSize = 16
        
    End Select
    
    'Resize the source DIB to fit a 2x2 block pattern of the requested checkerboard pattern
    srcDIB.CreateBlank chkSize * 2, chkSize * 2, 32
    
    Dim chkLookup() As Byte
    ReDim chkLookup(0 To chkSize * 2) As Byte
    Dim x As Long, y As Long
    For x = 0 To chkSize * 2
        chkLookup(x) = x \ chkSize
    Next x
    
    'Point a temporary array directly at the source DIB's bitmap bits.
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'Fill the source DIB with the checkerboard pattern
    Dim chkWidth As Long
    chkWidth = srcDIB.GetDIBWidth - 1
    
    Dim xStride As Long
    For y = 0 To chkWidth
    For x = 0 To chkWidth
    
        xStride = x * 4
        
        If (((chkLookup(x) + chkLookup(y)) And 1) = 0) Then
            srcImageData(xStride, y) = b1
            srcImageData(xStride + 1, y) = g1
            srcImageData(xStride + 2, y) = r1
            srcImageData(xStride + 3, y) = 255
        Else
            srcImageData(xStride, y) = b2
            srcImageData(xStride + 1, y) = g2
            srcImageData(xStride + 2, y) = r2
            srcImageData(xStride + 3, y) = 255
        End If
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcImageData

End Sub

'Given a source DIB, fill it with a 2x2 alpha checkerboard pattern matching the user's current preferences.
' (The resulting DIB size is contingent on the user's checkerboard pattern size preference, FYI.)
Public Sub GetArbitraryCheckerboardDIB(ByRef srcDIB As pdDIB, ByVal chkColorOne As Long, ByVal chkColorTwo As Long, ByVal chkSize As Long)
    
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = Colors.ExtractRed(chkColorOne)
    r2 = Colors.ExtractRed(chkColorTwo)
    g1 = Colors.ExtractGreen(chkColorOne)
    g2 = Colors.ExtractGreen(chkColorTwo)
    b1 = Colors.ExtractBlue(chkColorOne)
    b2 = Colors.ExtractBlue(chkColorTwo)
    
    'Resize the source DIB to fit a 2x2 block pattern of the requested checkerboard pattern
    If (srcDIB Is Nothing) Then Set srcDIB = New pdDIB
    srcDIB.CreateBlank chkSize * 2, chkSize * 2, 32, initialAlpha:=255
    
    Dim chkLookup() As Byte
    ReDim chkLookup(0 To chkSize * 2) As Byte
    Dim x As Long, y As Long
    For x = 0 To chkSize * 2
        chkLookup(x) = x \ chkSize
    Next x
    
    'Point a temporary array directly at the source DIB's bitmap bits.
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'Fill the source DIB with the checkerboard pattern
    Dim chkWidth As Long
    chkWidth = srcDIB.GetDIBWidth - 1
    
    Dim xStride As Long
    For y = 0 To chkWidth
    For x = 0 To chkWidth
    
        xStride = x * 4
        
        If (((chkLookup(x) + chkLookup(y)) And 1) = 0) Then
            srcImageData(xStride, y) = b1
            srcImageData(xStride + 1, y) = g1
            srcImageData(xStride + 2, y) = r1
        Else
            srcImageData(xStride, y) = b2
            srcImageData(xStride + 1, y) = g2
            srcImageData(xStride + 2, y) = r2
        End If
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcImageData

End Sub

'Given an (x,y) pair on the current viewport, convert the value to coordinates on the image.
Public Function ConvertCanvasCoordsToImageCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Double, ByVal canvasY As Double, ByRef imgX As Double, ByRef imgY As Double, Optional ByVal forceInBounds As Boolean = False) As Boolean

    If (Not srcImage Is Nothing) Then
    
        'Get the current zoom value from the source image, then invert it.  (We're only going to use that value in division.)
        Dim zoomVal As Double
        zoomVal = 1# / Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex)
        
        'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
        ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
        ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
        ' so we can ignore their value(s) here.
        Dim translatedImageRect As RectF
        srcImage.ImgViewport.GetImageRectTranslated translatedImageRect
        
        'Translating the canvas coordinate pair back to the image is now easy.  Subtract the top/left offset,
        ' then divide by zoom - that's all there is to it!
        imgX = (canvasX - translatedImageRect.Left) * zoomVal
        imgY = (canvasY - translatedImageRect.Top) * zoomVal
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
            If (imgX < 0#) Then imgX = 0#
            If (imgY < 0#) Then imgY = 0#
            If (imgX >= srcImage.Width - 1) Then imgX = srcImage.Width - 1
            If (imgY >= srcImage.Height - 1) Then imgY = srcImage.Height - 1
        End If
        
        ConvertCanvasCoordsToImageCoords = True
        
    Else
        ConvertCanvasCoordsToImageCoords = False
    End If
    
End Function

'Given an (x,y) pair on the current image, convert the value to coordinates on the current viewport canvas.
Public Sub ConvertImageCoordsToCanvasCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal imgX As Double, ByVal imgY As Double, ByRef canvasX As Double, ByRef canvasY As Double, Optional ByVal forceInBounds As Boolean = False)

    If Not (srcImage.ImgViewport Is Nothing) Then
    
        'Get the current zoom value from the source image
        Dim zoomVal As Double
        zoomVal = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex)
            
        'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
        ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
        ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
        ' so we can ignore their value(s) here.
        Dim translatedImageRect As RectF
        srcImage.ImgViewport.GetImageRectTranslated translatedImageRect
        
        'Translating the canvas coordinate pair back to the image is now easy.  Add the top/left offset,
        ' then multiply by zoom - that's all there is to it!
        canvasX = (imgX * zoomVal) + translatedImageRect.Left
        canvasY = (imgY * zoomVal) + translatedImageRect.Top
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
        
            'Get a copy of the current viewport intersection rect, which determines bounds of this function
            Dim vIntersectRect As RectF
            srcImage.ImgViewport.GetIntersectRectCanvas vIntersectRect
            
            If (canvasX < vIntersectRect.Left) Then canvasX = vIntersectRect.Left
            If (canvasY < vIntersectRect.Top) Then canvasY = vIntersectRect.Top
            If (canvasX >= vIntersectRect.Left + vIntersectRect.Width) Then canvasX = vIntersectRect.Left + vIntersectRect.Width - 1
            If (canvasY >= vIntersectRect.Top + vIntersectRect.Height) Then canvasY = vIntersectRect.Top + vIntersectRect.Height - 1
            
        End If
        
    End If
    
End Sub

'Given a RectF containing image-space coordinates, produce a new RectF with coordinates translated to the specified viewport canvas.
Public Sub ConvertImageCoordsToCanvasCoords_RectF(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcRectF As RectF, ByRef dstRectF As RectF, Optional ByVal forceInBounds As Boolean = False)

    If (Not srcImage.ImgViewport Is Nothing) Then
    
        'Get the current zoom value from the source image
        Dim zoomVal As Double
        zoomVal = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
            
        'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
        ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
        ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
        ' so we can ignore their value(s) here.
        Dim translatedImageRect As RectF
        srcImage.ImgViewport.GetImageRectTranslated translatedImageRect
        
        'Translating the canvas coordinate pair back to the image is now easy.  Add the top/left offset,
        ' then multiply by zoom - that's all there is to it!
        dstRectF.Left = (srcRectF.Left * zoomVal) + translatedImageRect.Left
        dstRectF.Top = (srcRectF.Top * zoomVal) + translatedImageRect.Top
        
        'Width/height are even easier - just multiply by the current zoom
        dstRectF.Width = srcRectF.Width * zoomVal
        dstRectF.Height = srcRectF.Height * zoomVal
        
        'If the caller wants the coordinates bound-checked, apply them last
        If forceInBounds Then
        
            'Get a copy of the current viewport intersection rect, which determines bounds of this function
            Dim vIntersectRect As RectF
            srcImage.ImgViewport.GetIntersectRectCanvas vIntersectRect
            
            If (dstRectF.Left < vIntersectRect.Left) Then dstRectF.Left = vIntersectRect.Left
            If (dstRectF.Top < vIntersectRect.Top) Then dstRectF.Top = vIntersectRect.Top
            If (dstRectF.Left + dstRectF.Width >= vIntersectRect.Left + vIntersectRect.Width) Then
                dstRectF.Width = (vIntersectRect.Left + vIntersectRect.Width - 1) - dstRectF.Left
                If dstRectF.Width < 0 Then dstRectF.Width = 0
            End If
            If (dstRectF.Top + dstRectF.Height >= vIntersectRect.Top + vIntersectRect.Height) Then
                dstRectF.Top = (vIntersectRect.Top + vIntersectRect.Height - 1) - dstRectF.Height
                If dstRectF.Height < 0 Then dstRectF.Height = 0
            End If
            
        End If
        
    End If
    
End Sub

'Given an (x,y) pair on the current image, convert the value to coordinates relative to the current layer.
' This is especially relevant if the layer has one or more non-destructive affine transforms active.
Public Function ConvertImageCoordsToLayerCoords(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal imgX As Single, ByVal imgY As Single, ByRef layerX As Single, ByRef layerY As Single) As Boolean

    If (srcImage Is Nothing) Then Exit Function
    If (srcLayer Is Nothing) Then Exit Function
    
    'If the layer has one or more active affine transforms, this step becomes complicated.
    If srcLayer.AffineTransformsActive(False) Then
    
        'Create a copy of either the layer's transformation matrix, or a custom matrix if passed in
        Dim tmpMatrix As pd2DTransform
        srcLayer.GetCopyOfLayerTransformationMatrix tmpMatrix
        
        'Invert the matrix
        If tmpMatrix.InvertTransform() Then
            
            'We now need to convert the image coordinates against the layer transformation matrix
            tmpMatrix.ApplyTransformToXY imgX, imgY
            
            'In order for the matrix conversion to work, it has to offset coordinates by the current layer offset.  (Rotation is
            ' particularly important in that regard, as the center-point is crucial.)  As such, we now need to undo that translation.
            ' In rare circumstances the caller can disable this behavior, for example while transforming a layer, because the original
            ' rotation matrix must be used.
            layerX = imgX + srcLayer.GetLayerOffsetX
            layerY = imgY + srcLayer.GetLayerOffsetY
            
            ConvertImageCoordsToLayerCoords = True
        
        'If we can't invert the matrix, we're in trouble.  Copy out the layer coordinates as a failsafe.
        Else
            
            layerX = imgX
            layerY = imgY
            
            Debug.Print "WARNING! Transformation matrix could not be generated."
            
            ConvertImageCoordsToLayerCoords = False
            
        End If
    
    'If the layer doesn't have affine transforms active, this step is easy.
    Else
    
        'Layer coordinates are identical to image coordinates
        layerX = imgX
        layerY = imgY
        
        ConvertImageCoordsToLayerCoords = True
    
    End If
    
End Function

'Given an (x,y) pair on the current image, convert the value to coordinates relative to the current layer.  Note that *all*
' layer transform properties are considered (including x/y offsets, scaling, rotation, and skew).  As such, you should not
' handle any of these properties specially.
Public Function ConvertImageCoordsToLayerCoords_Full(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal imgX As Single, ByVal imgY As Single, ByRef layerX As Single, ByRef layerY As Single) As Boolean

    If (srcImage Is Nothing) Then Exit Function
    If (srcLayer Is Nothing) Then Exit Function
    
    'If the layer has one or more active affine transforms, this step becomes complicated.
    If srcLayer.AffineTransformsActive(True) Then
    
        'Create a copy of either the layer's transformation matrix, or a custom matrix if passed in
        Dim tmpMatrix As pd2DTransform
        srcLayer.GetCopyOfLayerTransformationMatrix_Full tmpMatrix
        
        'Invert the matrix
        If tmpMatrix.InvertTransform() Then
            
            'Apply the matrix to the incoming image coordinates, then return them!
            tmpMatrix.ApplyTransformToXY imgX, imgY
            layerX = imgX
            layerY = imgY
            
            ConvertImageCoordsToLayerCoords_Full = True
        
        'If we can't invert the matrix, we're in trouble.  Copy out the incoming image coordinates as a failsafe.
        Else
            
            layerX = imgX
            layerY = imgY
            
            Debug.Print "WARNING! Transformation matrix could not be generated."
            
            ConvertImageCoordsToLayerCoords_Full = False
            
        End If
    
    'If the layer doesn't have affine transforms active, this step is easy.  The only "transform" we need to consider are the
    ' layer's offsets (which may be non-zero).
    Else
        layerX = imgX - srcLayer.GetLayerOffsetX
        layerY = imgY - srcLayer.GetLayerOffsetY
        ConvertImageCoordsToLayerCoords_Full = True
    End If
    
End Function

'Given an array of (x,y) pairs set in the current image's coordinate space, convert each pair to the supplied viewport canvas space.
Public Sub ConvertListOfImageCoordsToCanvasCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef listOfPoints() As PointFloat, Optional ByVal forceInBounds As Boolean = False)

    If (srcImage.ImgViewport Is Nothing) Then Exit Sub
    
    'Get the current zoom value from the source image
    Dim zoomVal As Double
    zoomVal = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
    
    'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
    ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
    ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
    ' so we can ignore their value(s) here.
    Dim translatedImageRect As RectF
    srcImage.ImgViewport.GetImageRectTranslated translatedImageRect
    
    'If the caller wants the coordinates bound-checked, we also need to grab a copy of the viewport
    ' intersection rect, which controls boundaries
    Dim vIntersectRect As RectF
    If forceInBounds Then srcImage.ImgViewport.GetIntersectRectCanvas vIntersectRect
    
    Dim canvasX As Double, canvasY As Double
    
    'Iterate through each point in turn; note that bounds are automatically detected, and there is not currently a way to override
    ' this behavior.
    Dim i As Long
    For i = LBound(listOfPoints) To UBound(listOfPoints)
        
        'Translating the canvas coordinate pair back to the image is now easy.  Add the top/left offset,
        ' then multiply by zoom - that's all there is to it!
        canvasX = (listOfPoints(i).x * zoomVal) + translatedImageRect.Left
        canvasY = (listOfPoints(i).y * zoomVal) + translatedImageRect.Top
        
        'If the caller wants the coordinates bound-checked, apply it now
        If forceInBounds Then
            If (canvasX < vIntersectRect.Left) Then canvasX = vIntersectRect.Left
            If (canvasY < vIntersectRect.Top) Then canvasY = vIntersectRect.Top
            If (canvasX >= vIntersectRect.Left) + vIntersectRect.Width Then canvasX = vIntersectRect.Left + vIntersectRect.Width - 1
            If (canvasY >= vIntersectRect.Top) + vIntersectRect.Height Then canvasY = vIntersectRect.Top + vIntersectRect.Height - 1
        End If
        
        'Store the updated coordinate pair
        listOfPoints(i).x = canvasX
        listOfPoints(i).y = canvasY
    
    Next i
        
End Sub

'Given an (x,y) pair on the current image, convert the values to coordinates in the current display coordinate space.
' (This is used for handling UI stuff as the user interacts with the canvas, and using image coordinates allows for a
' generalized solution.)
Public Sub ConvertImageCoordsToScreenCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal imgX As Double, ByVal imgY As Double, ByRef screenX As Long, ByRef screenY As Long, Optional ByVal forceInBounds As Boolean = False)
    
    'Start by converting image coordinates to canvas coordinates.
    Dim canvasX As Double, canvasY As Double
    Drawing.ConvertImageCoordsToCanvasCoords srcCanvas, srcImage, imgX, imgY, canvasX, canvasY, forceInBounds
    
    'We now need to map from canvas coordinate space to screen coordinate space
    If (Not g_WindowManager Is Nothing) Then
        
        'Map using PD's internal window manager (which wraps MapWindowPoint)
        Dim tmpPoint As PointAPI
        tmpPoint.x = Int(canvasX + 0.5)
        tmpPoint.y = Int(canvasY + 0.5)
        g_WindowManager.GetClientToScreen_Universal srcCanvas.GetCanvasViewHWnd, VarPtr(tmpPoint)
        
        'Return the final values
        screenX = tmpPoint.x
        screenY = tmpPoint.y
        
    End If
    
End Sub

'If you want to convert a position-agnostic size between image and canvas space, use these functions
Public Function ConvertCanvasSizeToImageSize(ByVal srcSize As Double, ByRef srcImage As pdImage) As Double
    ConvertCanvasSizeToImageSize = srcSize / Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
End Function

Public Function ConvertImageSizeToCanvasSize(ByVal srcSize As Double, ByRef srcImage As pdImage) As Double
    ConvertImageSizeToCanvasSize = srcSize * Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
End Function

'Return an arbitrary conversion from image space to canvas space.
' An optional image (x, y) can also passed; these will be added to the transform as source-image-space offsets.
Public Sub GetTransformFromImageToCanvas(ByRef dstTransform As pd2DTransform, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, Optional ByVal srcX As Single = 0!, Optional ByVal srcY As Single = 0!)

    If (dstTransform Is Nothing) Then Set dstTransform = New pd2DTransform

    'Get the current zoom value from the source image
    Dim zoomVal As Double
    zoomVal = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
    
    'Get a copy of the translated image rect, in canvas coordinates.  If the canvas is a window, and the zoomed
    ' image is a poster sliding around behind it, the translate image rect contains the poster coordinates,
    ' relative to the window.  What's great about this rect is that it's already accounted for scroll bars,
    ' so we can ignore their value(s) here.
    Dim translatedImageRect As RectF
    srcImage.ImgViewport.GetImageRectTranslated translatedImageRect
    
    'Apply scaling for zoom
    dstTransform.ApplyScaling zoomVal, zoomVal
    
    'Translate according to the current viewport setting, plus the original coordinates, if any
    dstTransform.ApplyTranslation (srcX * zoomVal) + translatedImageRect.Left, (srcY * zoomVal) + translatedImageRect.Top
    
End Sub

'On the current viewport, render lines around the active layer
Public Sub DrawLayerBoundaries(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer)

    'In the old days, we could get away with assuming layer boundaries form a rectangle, but as of PD 7.0, affine transforms
    ' mean this is no longer guaranteed.
    '
    'So instead of filling a rect, we must retrieve the four layer corner coordinates as floating-point pairs.
    Dim layerCorners() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    
    srcLayer.GetLayerCornerCoordinates layerCorners
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    'Pass the list of coordinates to a pd2DPath object; it will handle the actual UI rendering
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    
    'Note that we must add the layer boundary lines manually - otherwise, the top-right and bottom-left corners will connect
    ' due to the way srcLayer.getLayerCornerCoordinates returns the points!
    tmpPath.AddLine layerCorners(0).x, layerCorners(0).y, layerCorners(1).x, layerCorners(1).y
    tmpPath.AddLine layerCorners(1).x, layerCorners(1).y, layerCorners(3).x, layerCorners(3).y
    tmpPath.AddLine layerCorners(3).x, layerCorners(3).y, layerCorners(2).x, layerCorners(2).y
    tmpPath.CloseCurrentFigure
    
    'Render the final UI
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    PD2D.DrawPath cSurface, m_PenUIBase, tmpPath
    PD2D.DrawPath cSurface, m_PenUITop, tmpPath
    Set cSurface = Nothing
    
End Sub

'On the current viewport, render standard PD transformation nodes (layer corners, currently) atop the active layer.
Public Sub DrawLayerCornerNodes(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)
    
    'In the old days, we could get away with assuming layer boundaries form a rectangle,
    ' but as of PD 7.0, affine transforms mean this is no longer guaranteed.
    '
    'So instead of filling a rect, we must retrieve the four layer corner coordinates as floating-point pairs.
    Dim layerCorners() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    
    srcLayer.GetLayerCornerCoordinates layerCorners
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    Dim cornerSize As Single, halfCornerSize As Single
    cornerSize = SQUARE_CORNER_SIZE
    halfCornerSize = cornerSize * 0.5!
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    
    'Convert the POI value, if any, to an index into our list of layer coordinates
    If (curPOI <> poi_Undefined) Then
        If (curPOI = poi_CornerNW) Then
            curPOI = 0
        ElseIf (curPOI = poi_CornerNE) Then
            curPOI = 1
        ElseIf (curPOI = poi_CornerSW) Then
            curPOI = 2
        ElseIf (curPOI = poi_CornerSE) Then
            curPOI = 3
        End If
    End If
    
    'Use GDI+ to render four corner nodes
    Dim i As Long
    For i = 0 To 3
        If (i = curPOI) Then
            PD2D.DrawRectangleF cSurface, m_PenUIBaseHighlight, layerCorners(i).x - halfCornerSize, layerCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, m_PenUITopHighlight, layerCorners(i).x - halfCornerSize, layerCorners(i).y - halfCornerSize, cornerSize, cornerSize
        Else
            PD2D.DrawRectangleF cSurface, m_PenUIBase, layerCorners(i).x - halfCornerSize, layerCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, m_PenUITop, layerCorners(i).x - halfCornerSize, layerCorners(i).y - halfCornerSize, cornerSize, cornerSize
        End If
    Next i
    
End Sub

'As of May '24, PD can now render dynamic distances between the current layer and neighboring objects.
' To do this, it needs to query the canvas for mouse position (to know whether the mouse is actively moving
' the current layer, or just hovering other objects).
Public Sub DrawLayerDistances(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)
    
    'This render is handled in several passes.
    
    'First, render the interior dimensions of this layer.  (These are helpful when resizing a layer,
    ' or for doing quick math to add layer distance plus size.)
    DrawLayerDistances_Interior dstCanvas, srcImage, srcLayer, curPOI
    
    'Next, distance between layer edges and canvas edges.  The complexity of this step varies based on
    ' whether the active layer has affine transforms active.
    If srcLayer.AffineTransformsActive(False) Then
        DrawLayerDistances_EdgeAffine dstCanvas, srcImage, srcLayer, curPOI
        
    'Normal layers (with no active affine transforms except scaling) are much easier for distance rendering!
    Else
        DrawLayerDistances_EdgeNoAffine dstCanvas, srcImage, srcLayer, curPOI
    End If
    
    'How and where we draw distances depends on a few different things:
    ' 1) the active layer's size, position, and rotation
    ' 2) the current mouse state (button up/down)
    ' 2a) the current mouse position (over the active layer, over another layer, over the canvas)
    
End Sub

'Draw distances from each layer corner to each canvas edge, with proper handling for off-canvas corners
' (use only on layers with active affine transforms; otherwise, we want to draw from layer *edge* instead
'  of layer *corner, and there is a dedicated function for that).
Private Sub DrawLayerDistances_EdgeAffine(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)

    'Retrieve the all layer corners as floating-point pairs.
    Dim layerCorners() As PointFloat, layerCornersInImageCoords() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    ReDim layerCornersInImageCoords(0 To 3) As PointFloat
    srcLayer.GetLayerCornerCoordinates layerCorners
    
    'From the list of layer corner coordinates, we now want to determine the left-most, top-most,
    ' right-most, and bottom-most points.  We will place these in a new array and work exclusively
    ' with it from here out.
    Dim layerCornersInOrder() As PointFloat
    ReDim layerCornersInOrder(0 To 3) As PointFloat
    
    Dim idxMost As Long
    idxMost = 0
    
    Dim i As Long
    For i = 1 To 3
        If (layerCorners(i).x < layerCorners(idxMost).x) Then idxMost = i
    Next i
    layerCornersInOrder(0) = layerCorners(idxMost)
    
    idxMost = 0
    For i = 1 To 3
        If (layerCorners(i).y < layerCorners(idxMost).y) Then idxMost = i
    Next i
    layerCornersInOrder(1) = layerCorners(idxMost)
    
    idxMost = 0
    For i = 1 To 3
        If (layerCorners(i).x > layerCorners(idxMost).x) Then idxMost = i
    Next i
    layerCornersInOrder(2) = layerCorners(idxMost)
    
    idxMost = 0
    For i = 1 To 3
        If (layerCorners(i).y > layerCorners(idxMost).y) Then idxMost = i
    Next i
    layerCornersInOrder(3) = layerCorners(idxMost)
    
    For i = 0 To 3
        layerCornersInImageCoords(i) = layerCornersInOrder(i)
    Next i
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCornersInOrder, False
    
    'We also need to get the position of the canvas edges in viewport space.
    Dim canvasCorners() As PointFloat
    ReDim canvasCorners(0 To 3) As PointFloat
    With srcImage.GetBoundaryRectF
        canvasCorners(0).x = .Left
        canvasCorners(0).y = .Top
        canvasCorners(1).x = .Left + .Width
        canvasCorners(1).y = .Top
        canvasCorners(2).x = .Left
        canvasCorners(2).y = .Top + .Height
        canvasCorners(3).x = .Left + .Width
        canvasCorners(3).y = .Top + .Height
    End With
    
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, canvasCorners, False
    
    'All renderers will use the same drawing objects.
    
    'Surface...
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Half
    
    'Pen...
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenColor RGB(255, 0, 0)
    cPen.SetPenOpacity 100!
    cPen.SetPenWidth 1!
    cPen.SetPenStartCap P2_LC_ArrowAnchor
    cPen.SetPenEndCap P2_LC_ArrowAnchor
    
    'Brush...
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor RGB(255, 0, 0)
    
    'Text...
    Dim cText As pdFont
    Set cText = Fonts.GetMatchingUIFont(8!, True)
    
    Dim origFontColor As Long
    origFontColor = cText.GetFontColor()
    
    cText.AttachToDC dstCanvas.hDC
    cText.SetFontColor RGB(255, 255, 255)
    
    'Now we'll basically run this function four times: once per layer edge,
    ' comparing it to the matching canvas edge (and skipping if it lies *beyond* that canvas edge).
    
    '(TODO: show distance to other layers instead of canvas *if* the mouse hovers them.)
    
    'Start with the left-most point.
    
    'We want to draw a background box behind the text size, and a line behind *that* which runs from
    ' the left-most point to the nearest canvas edge.
    
    'Start by finding the endpoints for the background line on this side.
    Dim pt1 As PointFloat, pt2 As PointFloat, ptMid As PointFloat
    pt2.x = layerCornersInOrder(0).x
    pt2.y = layerCornersInOrder(0).y
    
    pt1.x = canvasCorners(0).x
    pt1.y = pt2.y
    
    'Only continue rendering if the canvas edge lies *outside* the layer edge
    If (pt1.x < pt2.x) Then
        
        'Draw the line
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        
        'Text is positioned on the midpoint of the line.
        Const FRACTION_OF_LAYER_FOR_TEXT As Single = 0.5!
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        
        'Next, calculate the relevant distance and convert it to a string
        Dim curDistanceAsString As String
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), layerCornersInImageCoords(0).x - srcImage.GetBoundaryRectF.Left, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
        
        'Render the text
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
        
    '/End "canvas edge lies beyond layer edge"
    End If
    
    'Remaining directions follow.  Comments are trimmed to save space; if you have questions,
    ' refer to the fully-commented "left line" code above.
    
    'Right-most point
    pt1.x = layerCornersInOrder(2).x
    pt1.y = layerCornersInOrder(2).y
    pt2.x = canvasCorners(1).x
    pt2.y = pt1.y
    
    If (pt1.x < pt2.x) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcImage.GetBoundaryRectF.Left + srcImage.GetBoundaryRectF.Width - layerCornersInImageCoords(2).x, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Top-most point
    pt1.x = layerCornersInOrder(1).x
    pt1.y = layerCornersInOrder(1).y
    pt2.x = pt1.x
    pt2.y = canvasCorners(0).y
    
    If (pt1.y > pt2.y) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), layerCornersInImageCoords(1).y - srcImage.GetBoundaryRectF.Top, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Bottom-most point
    pt1.x = layerCornersInOrder(3).x
    pt1.y = layerCornersInOrder(3).y
    pt2.x = pt1.x
    pt2.y = canvasCorners(2).y
    
    If (pt1.y < pt2.y) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcImage.GetBoundaryRectF.Top + srcImage.GetBoundaryRectF.Height - layerCornersInImageCoords(3).y, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Restore original DC font color (fonts are rendered using GDI), then free the font from the target DC
    cText.SetFontColor origFontColor
    cText.ReleaseFromDC
    
End Sub

'Draw distances from each layer edge to each canvas edge, with proper handling for off-canvas edges
' (and layer scaling but *NOT* other affine transforms).
Private Sub DrawLayerDistances_EdgeNoAffine(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)

    'Retrieve the all layer corners as floating-point pairs.
    Dim layerCorners() As PointFloat, layerCornersInImageCoords() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    ReDim layerCornersInImageCoords(0 To 3) As PointFloat
    srcLayer.GetLayerCornerCoordinates layerCorners
    srcLayer.GetLayerCornerCoordinates layerCornersInImageCoords
    
    'We rely on the caller to pre-check for affine transforms and call another renderer as necessary
    ' (see the sub name!).  Everything from here assumes a rectangular, non-skewed non-rotated layer.
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    'We also need to get the position of the canvas edges in viewport space.
    Dim canvasCorners() As PointFloat
    ReDim canvasCorners(0 To 3) As PointFloat
    With srcImage.GetBoundaryRectF
        canvasCorners(0).x = .Left
        canvasCorners(0).y = .Top
        canvasCorners(1).x = .Left + .Width
        canvasCorners(1).y = .Top
        canvasCorners(2).x = .Left
        canvasCorners(2).y = .Top + .Height
        canvasCorners(3).x = .Left + .Width
        canvasCorners(3).y = .Top + .Height
    End With
    
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, canvasCorners, False
    
    'All renderers will use the same drawing objects.
    
    'Surface...
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Half
    
    'Pen...
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenColor RGB(255, 0, 0)
    cPen.SetPenOpacity 100!
    cPen.SetPenWidth 1!
    cPen.SetPenStartCap P2_LC_ArrowAnchor
    cPen.SetPenEndCap P2_LC_ArrowAnchor
    
    'Brush...
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor RGB(255, 0, 0)
    
    'Text...
    Dim cText As pdFont
    Set cText = Fonts.GetMatchingUIFont(8!, True)
    
    Dim origFontColor As Long
    origFontColor = cText.GetFontColor()
    
    cText.AttachToDC dstCanvas.hDC
    cText.SetFontColor RGB(255, 255, 255)
    
    'Now we'll basically run this function four times: once per layer edge,
    ' comparing it to the matching canvas edge (and skipping if it lies *beyond* that canvas edge).
    
    '(TODO: show distance to other layers instead of canvas *if* the mouse hovers them.)
    
    'Start with the left edge.
    
    'We want to draw a background box behind the text size, and a line behind *that* which runs from
    ' the midpoint of the layer edge to the nearest canvas edge.
    
    'Start by finding the endpoints for the background line on this side.
    Dim pt1 As PointFloat, pt2 As PointFloat, ptMid As PointFloat
    pt2.x = layerCorners(0).x
    pt2.y = (layerCorners(0).y + layerCorners(2).y) * 0.5
    
    pt1.x = canvasCorners(0).x
    pt1.y = pt2.y
    
    'Only continue rendering if the canvas edge lies *outside* the layer edge
    If (pt1.x < pt2.x) Then
        
        'Draw the line
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        
        'Text is positioned on the midpoint of the line.
        Const FRACTION_OF_LAYER_FOR_TEXT As Single = 0.5!
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        
        'Next, calculate the relevant distance and convert it to a string
        Dim curDistanceAsString As String
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), layerCornersInImageCoords(0).x - srcImage.GetBoundaryRectF.Left, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
        
        'Render the text
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
        
    '/End "canvas edge lies beyond layer edge"
    End If
    
    'Remaining directions follow.  Comments are trimmed to save space; if you have questions,
    ' refer to the fully-commented "left line" code above.
    
    'Right side of layer
    pt1.x = layerCorners(1).x
    pt1.y = (layerCorners(0).y + layerCorners(2).y) * 0.5
    pt2.x = canvasCorners(1).x
    pt2.y = pt1.y
    
    If (pt1.x < pt2.x) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcImage.GetBoundaryRectF.Left + srcImage.GetBoundaryRectF.Width - layerCornersInImageCoords(1).x, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Top of layer
    pt1.x = (layerCorners(0).x + layerCorners(1).x) * 0.5
    pt1.y = layerCorners(0).y
    pt2.x = pt1.x
    pt2.y = canvasCorners(0).y
    
    If (pt1.y > pt2.y) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), layerCornersInImageCoords(0).y - srcImage.GetBoundaryRectF.Top, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Bottom of layer
    pt1.x = (layerCorners(0).x + layerCorners(1).x) * 0.5
    pt1.y = layerCorners(2).y
    pt2.x = pt1.x
    pt2.y = canvasCorners(2).y
    
    If (pt1.y < pt2.y) Then
        PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
        ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
        ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
        curDistanceAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcImage.GetBoundaryRectF.Top + srcImage.GetBoundaryRectF.Height - layerCornersInImageCoords(2).y, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
        RenderTextWithBackgroundBox curDistanceAsString, cText, ptMid, cSurface, cBrush
    End If
    
    'Restore original DC font color (fonts are rendered using GDI), then free the font from the target DC
    cText.SetFontColor origFontColor
    cText.ReleaseFromDC
    
End Sub

Private Sub RenderTextWithBackgroundBox(ByRef srcText As String, ByRef fontObject As pdFont, ByRef centerPoint As PointFloat, ByRef dstSurface As pd2DSurface, ByRef dstBrush As pd2DBrush)

    'Determine string size in pixels (as rendered on-canvas)
    Dim sizeTextX As String, sizeTextY As String
    sizeTextX = fontObject.GetWidthOfString(srcText)
    sizeTextY = fontObject.GetHeightOfString(srcText)
    
    'Calculate position of a background rectangle (behind the distance text)
    Const PADDING_AROUND_TEXT As Single = 3!
    Dim textRectL As RectL_WH
    textRectL.Width = sizeTextX + PADDING_AROUND_TEXT * 2
    textRectL.Height = sizeTextY + PADDING_AROUND_TEXT * 2
    textRectL.Left = Int(centerPoint.x - textRectL.Width * 0.5 + 0.5)
    textRectL.Top = Int(centerPoint.y - textRectL.Height * 0.5 + 0.5)
    
    'Render the background rectangle
    PD2D.FillRectangleI dstSurface, dstBrush, textRectL.Left, textRectL.Top, textRectL.Width, textRectL.Height
    
    'Position the text inside the rectangle
    Dim textRect_GDI As RECT
    With textRect_GDI
        .Left = textRectL.Left
        .Top = textRectL.Top
        .Right = textRectL.Left + textRectL.Width
        .Bottom = textRectL.Top + textRectL.Height
    End With
    
    'Draw the text
    fontObject.DrawCenteredTextToRect srcText, textRect_GDI, True

End Sub

Private Sub DrawLayerDistances_Interior(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)

    'In the old days, we could get away with assuming layer boundaries form a rectangle,
    ' but support for affine transforms mean this is no longer guaranteed.
    '
    'So instead of assuming a rect, we must retrieve the all layer corners as floating-point pairs.
    Dim layerCorners() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    srcLayer.GetLayerCornerCoordinates layerCorners
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerCorners, False
    
    'Now we'll basically run this function twice: once for layer width, then again for height.
    ' (Note that these steps need to work for normal *and* rotated layers!)
    
    'Start with rendering layer width.
    ' We want to draw a background box behind the text size, and a line behind *that* which runs from
    ' edge-to-edge on the target layer, at perhaps 20% of the way down the layer.
    Const FRACTION_OF_LAYER_FOR_LAYOUT As Single = 0.2!
    
    'Start by finding the endpoints for the horizontal line where we will define width.
    Dim pt1 As PointFloat, pt2 As PointFloat, ptMid As PointFloat
    pt1.x = (layerCorners(0).x + (layerCorners(2).x - layerCorners(0).x) * FRACTION_OF_LAYER_FOR_LAYOUT)
    pt1.y = (layerCorners(0).y + (layerCorners(2).y - layerCorners(0).y) * FRACTION_OF_LAYER_FOR_LAYOUT)
    
    pt2.x = (layerCorners(1).x + (layerCorners(3).x - layerCorners(1).x) * FRACTION_OF_LAYER_FOR_LAYOUT)
    pt2.y = (layerCorners(1).y + (layerCorners(3).y - layerCorners(1).y) * FRACTION_OF_LAYER_FOR_LAYOUT)
    
    'Text uses its own positioning along the line.
    Const FRACTION_OF_LAYER_FOR_TEXT As Single = 0.5!
    ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
    ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
    
    'Prep rendering objects
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Half
    
    'Determine pen colors
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenColor RGB(255, 0, 0)
    cPen.SetPenOpacity 100!
    cPen.SetPenWidth 1!
    cPen.SetPenStartCap P2_LC_ArrowAnchor
    cPen.SetPenEndCap P2_LC_ArrowAnchor
    
    'Background brush
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushColor RGB(255, 0, 0)
    
    'Font renderer
    Dim cText As pdFont
    Set cText = Fonts.GetMatchingUIFont(8!, True)
    
    'Back up font object color before changing it (this is a shared UI font object)
    Dim origFontColor As Long
    origFontColor = cText.GetFontColor()
    cText.AttachToDC dstCanvas.hDC
    cText.SetFontColor RGB(255, 255, 255)
    
    'Draw the line
    PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
    
    'Next, calculate all strings involved in the dimensions of the active layer.
    Dim layerSizeAsString As String
    layerSizeAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcLayer.GetLayerWidth(True), PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
    
    'Draw the text against a solid-color background box
    RenderTextWithBackgroundBox layerSizeAsString, cText, ptMid, cSurface, cBrush
    
    'Now repeat all the above steps, but for layer height!
    
    'Calculate line position
    pt1.x = (layerCorners(0).x + (layerCorners(1).x - layerCorners(0).x) * FRACTION_OF_LAYER_FOR_LAYOUT)
    pt1.y = (layerCorners(0).y + (layerCorners(1).y - layerCorners(0).y) * FRACTION_OF_LAYER_FOR_LAYOUT)
    pt2.x = (layerCorners(2).x + (layerCorners(3).x - layerCorners(2).x) * FRACTION_OF_LAYER_FOR_LAYOUT)
    pt2.y = (layerCorners(2).y + (layerCorners(3).y - layerCorners(2).y) * FRACTION_OF_LAYER_FOR_LAYOUT)
    ptMid.x = (pt1.x + pt2.x) * FRACTION_OF_LAYER_FOR_TEXT
    ptMid.y = (pt1.y + pt2.y) * FRACTION_OF_LAYER_FOR_TEXT
    
    'Draw the line
    PD2D.DrawLineF_FromPtF cSurface, cPen, pt1, pt2
    
    'Calculate height text, then draw it
    layerSizeAsString = Units.GetValueFormattedForUnit_FromPixel(FormMain.MainCanvas(0).GetRulerUnit(), srcLayer.GetLayerHeight(True), PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
    RenderTextWithBackgroundBox layerSizeAsString, cText, ptMid, cSurface, cBrush
    
    'Restore original font color, then free the font from the target DC
    cText.SetFontColor origFontColor
    cText.ReleaseFromDC
    
End Sub

'As of PD 7.0, on-canvas rotation is now supported.  Use this function to render the current rotation node.
Public Sub DrawLayerRotateNode(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)
    
    'Retrieve the layer rotate node position from the specified layer, and convert it into the canvas coordinate space
    Dim layerRotateNodes() As PointFloat
    ReDim layerRotateNodes(0 To 4) As PointFloat
    
    srcLayer.GetLayerRotationNodeCoordinates layerRotateNodes
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, layerRotateNodes, False
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    
    'Convert the POI value, if any, to an index into our list of layer coordinates
    If (curPOI <> poi_Undefined) Then
        If (curPOI = poi_EdgeE) Then
            curPOI = 1
        ElseIf (curPOI = poi_EdgeS) Then
            curPOI = 2
        ElseIf (curPOI = poi_EdgeW) Then
            curPOI = 3
        ElseIf (curPOI = poi_EdgeN) Then
            curPOI = 4
        End If
    End If
    
    'As a convenience to the user, we draw some additional UI features if a rotation node is actively hovered by the mouse.
    If (curPOI >= 1) And (curPOI <= 4) Then
        
        'First, draw a line from the center of the layer to the rotation node, to provide visual feedback on where the rotation
        ' will actually occur.
        Dim tmpPath As pd2DPath
        Set tmpPath = New pd2DPath
        tmpPath.AddLine layerRotateNodes(0).x, layerRotateNodes(0).y, layerRotateNodes(curPOI).x, layerRotateNodes(curPOI).y
        
        PD2D.DrawPath cSurface, m_PenUIBase, tmpPath
        PD2D.DrawPath cSurface, m_PenUITop, tmpPath
        
        'Next, we are going to draw an arc with arrows on the end, to display where the actual rotation will occur.
        ' (At present, we skip this step if shearing is active, as I haven't figured out how to correctly skew the arc into the
        '  proper on-screen coordinate space.)
        If (srcLayer.GetLayerShearX = 0#) And (srcLayer.GetLayerShearY = 0#) Then
            
            tmpPath.ResetPath
        
            'Start by finding the distance of the rotation line.
            Dim rRadius As Double
            rRadius = PDMath.DistanceTwoPoints(layerRotateNodes(0).x, layerRotateNodes(0).y, layerRotateNodes(curPOI).x, layerRotateNodes(curPOI).y)
            If (rRadius < 0.1) Then rRadius = 0.1
            
            'From there, bounds are easy-peasy
            Dim rotateBoundRect As RectF
            With rotateBoundRect
                .Left = layerRotateNodes(0).x - rRadius
                .Top = layerRotateNodes(0).y - rRadius
                .Width = rRadius * 2#
                .Height = rRadius * 2#
            End With
            
            'Arc sweep and arc length are inter-related.  What we ultimately want is a (roughly) equal arc size regardless of zoom or
            ' the underlying image size.  This is difficult to predict as larger images and/or higher zoom will result in larger arc widths
            ' for an identical radius.  As such, we hard-code an approximate arc length, then generate an arc sweep from it.
            '
            'In my testing, 80-ish pixels is a reasonably good size across many image dimensions.  Note that we *do* correct for DPI here.
            Dim arcLength As Double
            arcLength = Interface.FixDPIFloat(70)
            
            'Switching between arc length and sweep is easy; see https://en.wikipedia.org/wiki/Arc_%28geometry%29#Length_of_an_arc_of_a_circle
            Dim arcSweep As Double
            arcSweep = (arcLength * 180#) / (PI * rRadius)
            
            'Make sure the arc fits within a valid range (e.g. no complete circles or nearly-straight lines)
            If (arcSweep > 90#) Then arcSweep = 90#
            If (arcSweep < 30#) Then arcSweep = 30#
            
            'We need to modify the default layer angle depending on the current POI
            Dim relevantAngle As Double
            relevantAngle = srcLayer.GetLayerAngle + ((curPOI - 1) * 90#)
            tmpPath.AddArc rotateBoundRect, relevantAngle - (arcSweep * 0.5), arcSweep
            
            Dim prevLineCap As PD_2D_LineCap
            prevLineCap = m_PenUIBase.GetPenLineCap
            m_PenUIBase.SetPenLineCap P2_LC_ArrowAnchor
            m_PenUITop.SetPenLineCap P2_LC_ArrowAnchor
            
            cSurface.SetSurfacePixelOffset P2_PO_Half
            PD2D.DrawPath cSurface, m_PenUIBase, tmpPath
            PD2D.DrawPath cSurface, m_PenUITop, tmpPath
            cSurface.SetSurfacePixelOffset P2_PO_Normal
            
            m_PenUIBase.SetPenLineCap prevLineCap
            m_PenUITop.SetPenLineCap prevLineCap
            
        End If
        
    End If
    
    'Render the circles at each rotation point
    Dim circRadius As Single
    circRadius = 7!
    
    Dim i As Long
    For i = 1 To 4
        If (curPOI = i) Then
            PD2D.DrawCircleF cSurface, m_PenUIBaseHighlight, layerRotateNodes(i).x, layerRotateNodes(i).y, circRadius
            PD2D.DrawCircleF cSurface, m_PenUITopHighlight, layerRotateNodes(i).x, layerRotateNodes(i).y, circRadius
        Else
            PD2D.DrawCircleF cSurface, m_PenUIBase, layerRotateNodes(i).x, layerRotateNodes(i).y, circRadius
            PD2D.DrawCircleF cSurface, m_PenUITop, layerRotateNodes(i).x, layerRotateNodes(i).y, circRadius
        End If
    Next i
    
End Sub

Public Sub DrawSmartGuides(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage)
    
    'Drawing smart guides is *optional*
    If (Not m_ShowSmartGuides) Then Exit Sub
    If (Not Snap.GetSnap_Any()) Then Exit Sub
    
    Dim smartGuideLine() As PointFloat
    ReDim smartGuideLine(0 To 1) As PointFloat
    
    'Look for an active x-guide
    If Snap.IsSnapped_X() Then
        
        Snap.GetSnappedX_SmartGuide smartGuideLine(0), smartGuideLine(1)
        
        'Convert the smart guidelines coordinates into viewport space
        Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, smartGuideLine, False
        
        'Use pd2D to perform the render
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
        
        PD2D.DrawLineF_FromPtF cSurface, m_PenUIBaseHighlight, smartGuideLine(0), smartGuideLine(1)
        PD2D.DrawLineF_FromPtF cSurface, m_PenUITopHighlight, smartGuideLine(0), smartGuideLine(1)
        
        Set cSurface = Nothing
        
    End If
    
    'Same for y
    If Snap.IsSnapped_Y() Then
    
        Snap.GetSnappedY_SmartGuide smartGuideLine(0), smartGuideLine(1)
        Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, smartGuideLine, False
        
        'Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
        
        PD2D.DrawLineF_FromPtF cSurface, m_PenUIBaseHighlight, smartGuideLine(0), smartGuideLine(1)
        PD2D.DrawLineF_FromPtF cSurface, m_PenUITopHighlight, smartGuideLine(0), smartGuideLine(1)
        
        Set cSurface = Nothing
        
    End If
    
End Sub

Public Function Get_ShowLayerEdges() As Boolean
    Get_ShowLayerEdges = m_ShowLayerEdges
End Function

Public Sub Set_ShowLayerEdges(ByVal newState As Boolean)
    m_ShowLayerEdges = newState
End Sub

Public Function Get_ShowSmartGuides() As Boolean
    Get_ShowSmartGuides = m_ShowSmartGuides
End Function

Public Sub Set_ShowSmartGuides(ByVal newState As Boolean)
    m_ShowSmartGuides = newState
End Sub

'Toggle one of the "Show extras..." settings in the View menu.
' To forcibly set to a specific state (instead of toggling), set the forceInsteadOfToggle param to TRUE.
Public Sub ToggleShowOptions(ByVal showTarget As PD_ShowTargets, Optional ByVal forceInsteadOfToggle As Boolean = False, Optional ByVal newState As Boolean = True)
    
    'While calculating which on-screen menu to update, we also need to relay changes to two places:
    ' 1) the tools_move module (which handles actual snap calculations)
    ' 2) the user preferences file (to ensure everything is synchronized between sessions)
    Select Case showTarget
        Case pdst_LayerEdges
            If (Not forceInsteadOfToggle) Then newState = Not Drawing.Get_ShowLayerEdges()
            Drawing.Set_ShowLayerEdges newState
            UserPrefs.SetPref_Boolean "Interface", "show-layeredges", newState
            Menus.SetMenuChecked "show_layeredges", newState
            
        Case pdst_SmartGuides
            If (Not forceInsteadOfToggle) Then newState = Not Drawing.Get_ShowSmartGuides()
            Drawing.Set_ShowSmartGuides newState
            UserPrefs.SetPref_Boolean "Interface", "show-smartguides", newState
            Menus.SetMenuChecked "show_smartguides", newState
            
    End Select
            
End Sub
            
'During startup, we cache a few different UI pens and brushes; this accelerates the process of viewport rendering.
' When the UI theme changes, this cache should be regenerated against any new colors.
'
'(Also, note the corresponding "release" function below.)
Public Sub CacheUIPensAndBrushes()
    Drawing2D.QuickCreatePairOfUIPens m_PenUIBase, m_PenUITop
    Drawing2D.QuickCreatePairOfUIPens m_PenUIBaseHighlight, m_PenUITopHighlight, True
End Sub

Public Sub ReleaseUIPensAndBrushes()
    Set m_PenUIBase = Nothing
    Set m_PenUITop = Nothing
    Set m_PenUIBaseHighlight = Nothing
    Set m_PenUITopHighlight = Nothing
End Sub

Public Sub BorrowCachedUIPens(ByRef dstPenUIBase As pd2DPen, ByRef dstPenUITop As pd2DPen, Optional ByVal wantHighlightPens As Boolean = False)
    If wantHighlightPens Then
        Set dstPenUIBase = m_PenUIBaseHighlight
        Set dstPenUITop = m_PenUITopHighlight
    Else
        Set dstPenUIBase = m_PenUIBase
        Set dstPenUITop = m_PenUITop
    End If
End Sub

Public Sub CloneCachedUIPens(ByRef dstPenUIBase As pd2DPen, ByRef dstPenUITop As pd2DPen, Optional ByVal wantHighlightPens As Boolean = False)
    Set dstPenUIBase = New pd2DPen
    Set dstPenUITop = New pd2DPen
    If wantHighlightPens Then
        dstPenUIBase.ClonePen m_PenUIBaseHighlight
        dstPenUITop.ClonePen m_PenUITopHighlight
    Else
        dstPenUIBase.ClonePen m_PenUIBase
        dstPenUITop.ClonePen m_PenUITop
    End If
End Sub
