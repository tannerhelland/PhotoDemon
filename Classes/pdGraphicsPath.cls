VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGraphicsPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Path Wrapper
'Copyright 2014-2015 by Tanner Helland
'Created: 17/October/14 (though assembled from various parts written much earlier)
'Last updated: 09/May/15
'Last update: integrate with new pdGraphixMatrix class, for easier affine transformations
'Dependencies: pdGraphicsMatrix, for applying affine transformations to a path.
'
'This class is a VB6-friendly wrapper for the GDI+ GraphicsPath object.  It is not intended as a comprehensive interface;
' instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ path handle is created at class initialization, and released at class termination.  The path handle is
' persistent by design, so no functions are provided for recreating or deleting it.  If you need to start over, use
' the resetPath() function.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Creation and basic maintenance functions
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal mBrushMode As GDIFillMode, ByRef hPath As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTF, ByVal tmpTransformMatrix As Long, ByVal tmpPen As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As Long

'Path property changes (both destructive and non-destructive)
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal newBrushMode As GDIFillMode) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newSmoothingMode As SmoothingMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal pixOffsetMode As PixelOffsetMode) As Long
Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal widenFlatness As Single) As Long
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformationMatrix As Long, ByVal outlineFlatness As Single) As Long

Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal penStartCap As LineCap, ByVal penEndCap As LineCap, ByVal penDashCap As DashCap) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal penLineJoin As LineJoin) As Long

'Figure-specific path additions
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Long
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Long
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As Long
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal hPath As Long) As Long

'General matrix transformations
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal hPath As Long, ByVal srcMatrix As Long) As Long

'Pen functions (for stroking a path)
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal penColor As Long, ByVal penWidth As Single, ByVal penUnit As GpUnit, ByRef hPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As Long

'Drawing functions
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As Long
Private Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Long, ByVal y As Long) As Long

'Misc GDI+ support types and functions
Private Const FlatnessDefault As Single = 0.25      'Allowable error when flattening a path.  This value is poorly explained on MSDN, but lower values equal a more precise approximation.  See http://msdn.microsoft.com/en-us/library/ms535572%28v=vs.85%29.aspx

Private Type tmpLong
    lngResult As Long
End Type

'The handle to our GDI+ path object.  This handle is automatically created with the class, and released when the
' class is terminated.  The caller doesn't have to worry about it at all.
Private m_PathHandle As Long

'***************************************************************************
'
' BEGIN: ADDING SHAPES AND FIGURES TO THE PATH

'The following collection of subs are simply convenient wrappers for adding various shapes to the current path object.
' NOTE: all paths are contiguous by default, so the first point of any added shape will be automatically connected to
'       the last point of the previous shape.

Public Sub addCircle(ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single)
    Me.addEllipse_Absolute centerX - circleRadius, centerY - circleRadius, centerX + circleRadius, centerY + circleRadius
End Sub

Public Sub addEllipse_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    GdipAddPathEllipse m_PathHandle, x1, y1, x2 - x1, y2 - y1
End Sub

Public Sub addEllipse_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single)
    GdipAddPathEllipse m_PathHandle, x1, y1, ellipseWidth, ellipseHeight
End Sub

Friend Sub addEllipse_RectF(ByRef srcRect As RECTF)
    With srcRect
        GdipAddPathEllipse m_PathHandle, .Left, .Top, .Width, .Height
    End With
End Sub

Public Sub addLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    GdipAddPathLine m_PathHandle, x1, y1, x2, y2
End Sub

Friend Sub addArc(ByRef ellipseBoundRect As RECTF, ByVal startAngle As Single, ByVal sweepAngle As Single)
    With ellipseBoundRect
        GdipAddPathArc m_PathHandle, .Left, .Top, .Width, .Height, startAngle, sweepAngle
    End With
End Sub

Public Sub addArcCircular(ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    GdipAddPathArc m_PathHandle, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle
End Sub

Public Sub addPolygon(ByVal numOfPoints As Long, ByVal ptrToFloatArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)

    'We have a few different options for adding this shape, based on the passed parameters.
    If autoCloseShape Then
    
        If useCurveAlgorithm Then
            GdipAddPathClosedCurve2 m_PathHandle, ptrToFloatArray, numOfPoints, curvatureTension
        Else
            GdipAddPathPolygon m_PathHandle, ptrToFloatArray, numOfPoints
        End If
        
    Else
    
        If useCurveAlgorithm Then
            GdipAddPathCurve2 m_PathHandle, ptrToFloatArray, numOfPoints, curvatureTension
        Else
            GdipAddPathLine2 m_PathHandle, ptrToFloatArray, numOfPoints
        End If
    
    End If
    
End Sub

Public Sub addRectangle_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    GdipAddPathRectangle m_PathHandle, x1, y1, x2 - x1, y2 - y1
End Sub

Public Sub addRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single)
    GdipAddPathRectangle m_PathHandle, x1, y1, rectWidth, rectHeight
End Sub

Friend Sub addRectangle_RectF(ByRef srcRect As RECTF)
    With srcRect
        GdipAddPathRectangle m_PathHandle, .Left, .Top, .Width, .Height
    End With
End Sub

Public Sub addRoundedRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single)
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If xCurvature > rectWidth Then xCurvature = rectWidth
    If yCurvature > rectHeight Then yCurvature = rectHeight
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    GdipAddPathArc m_PathHandle, x1 + rectWidth - xCurvature, y1, xCurvature, yCurvature, 270, 90
    GdipAddPathArc m_PathHandle, x1 + rectWidth - xCurvature, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 0, 90
    GdipAddPathArc m_PathHandle, x1, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 90, 90
    GdipAddPathArc m_PathHandle, x1, y1, xCurvature, yCurvature, 180, 90
    GdipClosePathFigure m_PathHandle
    
End Sub

Public Sub addBezierCurve(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
    GdipAddPathBezier m_PathHandle, x1, y1, x2, y2, x3, y3, x4, y4
End Sub

'Convenience wrapper for triangles
Public Sub addTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)

    'Convert the incoming points to a float array and use the generic polygon wrapper to add 'em
    Dim tmpPoints() As POINTFLOAT
    ReDim tmpPoints(0 To 2) As POINTFLOAT
    
    tmpPoints(0).x = x1
    tmpPoints(0).y = y1
    tmpPoints(1).x = x2
    tmpPoints(1).y = y2
    tmpPoints(2).x = x3
    tmpPoints(2).y = y3
    
    Me.addPolygon 3, VarPtr(tmpPoints(0)), True, False
    
End Sub

Public Sub addPath(ByRef srcGraphicsPath As pdGraphicsPath, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0, Optional ByVal connectToLastPointOfThisPath As Boolean = False)
    
    'Make sure the glyph exists
    If srcGraphicsPath Is Nothing Then
        Debug.Print "WARNING!  pdGraphicsPath.addPath was handed a non-existent path."
        Exit Sub
    End If
    
    'If no offsets are specified, copy the path as-is
    If (offsetX = 0) And (offsetY = 0) Then
    
        GdipAddPathPath m_PathHandle, srcGraphicsPath.getCopyOfPathHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
        
    'If offsets were specified, we need to clone the path, translate it, then add it
    Else
    
        'Clone the path
        Dim tmpPath As pdGraphicsPath
        Set tmpPath = New pdGraphicsPath
        tmpPath.cloneExistingPath srcGraphicsPath
        
        'Translate the path
        tmpPath.translatePath offsetX, offsetY
        
        'Add it
        GdipAddPathPath m_PathHandle, tmpPath.getCopyOfPathHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
        
        'Free our temporary path object
        Set tmpPath = Nothing
    
    End If
    
End Sub

Public Sub startNewFigure()
    GdipStartPathFigure m_PathHandle
End Sub

Public Sub closeCurrentFigure()
    GdipClosePathFigure m_PathHandle
End Sub

'For pen previews, it's helpful to have the path class create its own little preview shape.  In the future, we could give the user a selection of
' sample paths, which would simply be constants passed to this function.
Friend Sub createSamplePathForRect(ByRef srcRect As RECTF, Optional ByVal hPadding As Single = 0, Optional ByVal vPadding As Single = 0)

    'A sample path is just a nice little curve that demonstrates a few obvious path elements for the user
    Dim samplePoints() As POINTFLOAT
    ReDim samplePoints(0 To 8) As POINTFLOAT
    
    Dim sampleRect As RECTF
    With sampleRect
        .Left = srcRect.Left + hPadding
        .Top = srcRect.Top + vPadding
        .Width = srcRect.Width - hPadding * 2
        .Height = srcRect.Height - vPadding * 2
    End With
    
    'First, we calculate x positions.  Note that these are *not* equally distributed, by design.
    With sampleRect
        samplePoints(0).x = .Left
        samplePoints(1).x = .Left + .Width * 0.04
        samplePoints(2).x = .Left + .Width * 0.1
        samplePoints(3).x = .Left + .Width * 0.18
        samplePoints(4).x = .Left + .Width * 0.3
        samplePoints(5).x = .Left + .Width * 0.46
        samplePoints(6).x = .Left + .Width * 0.64
        samplePoints(7).x = .Left + .Width * 0.9
        samplePoints(8).x = .Left + .Width
    End With
    
    'Next, we calculate y positions
    With sampleRect
        samplePoints(0).y = .Top + .Height * 0.5
        samplePoints(1).y = .Top
        samplePoints(2).y = .Top + .Height
        samplePoints(3).y = .Top + .Height * 0.1
        samplePoints(4).y = .Top + .Height * 0.8
        samplePoints(5).y = .Top + .Height * 0.3
        samplePoints(6).y = .Top + .Height * 0.7
        samplePoints(7).y = .Top + .Height * 0.5
        samplePoints(8).y = .Top + .Height * 0.5
    End With
    
    'Add the final positions to the path!
    Me.addPolygon 9, VarPtr(samplePoints(0)), False, True, 0.5

End Sub

' END: ADDING SHAPES AND FIGURES TO THE PATH
'
'***************************************************************************




'***************************************************************************
'
' BEGIN: NON-DESTRUCTIVE PATH EDITS

'Change the path's fill mode.  (This behavior is non-destructive.)
Public Sub setPathFillMode(ByVal newMode As GDIFillMode)
    GdipSetPathFillMode m_PathHandle, newMode
End Sub

'Translate the entire path by some amount in the x and/or y positions
Public Sub translatePath(ByVal xAmount As Single, ByVal yAmount As Single)

    'Retrieve a translation matrix
    Dim tmpMatrix As pdGraphicsMatrix
    Set tmpMatrix = New pdGraphicsMatrix
    tmpMatrix.TranslateMatrix xAmount, yAmount
    
    'Apply the transformation
    GdipTransformPath m_PathHandle, tmpMatrix.getMatrixHandle
    
    'Free the matrix
    Set tmpMatrix = Nothing
    
End Sub

'Transform the path by some arbitrary pdGraphicsMatrix object
Public Sub transformPathByMatrix(ByRef srcMatrix As pdGraphicsMatrix)
    
    'Apply the transformation immediately
    GdipTransformPath m_PathHandle, srcMatrix.getMatrixHandle
    
End Sub

'Rotate the path around its center point.  The center point is calculated automatically.
Public Sub rotatePathAroundItsCenter(ByRef rotateAngle As Single)
    
    'Create a transformation matrix
    Dim tmpMatrix As pdGraphicsMatrix
    Set tmpMatrix = New pdGraphicsMatrix
    
    'Figure out the path's center
    Dim pathRect As RECTF
    GdipGetPathWorldBounds m_PathHandle, pathRect, 0, 0
    
    'Update the transformation matrix with a matching rotation
    tmpMatrix.RotateMatrix rotateAngle, pathRect.Left + pathRect.Width / 2, pathRect.Top + pathRect.Height / 2, MatrixOrderAppend
    
    'Apply the transformation
    GdipTransformPath m_PathHandle, tmpMatrix.getMatrixHandle
    
End Sub

'Mirror the path around its center point.  The center point is calculated automatically.
Public Sub mirrorPathAroundItsCenter(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean)
    
    'Create a transformation matrix
    Dim tmpMatrix As pdGraphicsMatrix
    Set tmpMatrix = New pdGraphicsMatrix
    
    'Figure out the path's center
    Dim pathRect As RECTF
    GdipGetPathWorldBounds m_PathHandle, pathRect, 0, 0
    
    'Update the transformation matrix with a matching mirror operation
    tmpMatrix.MirrorMatrix mirrorHorizontal, mirrorVertical
    
    'Mirroring will reflect a path around the 0-axis, so we need to translate the path back into its original position now.
    If mirrorHorizontal Then tmpMatrix.TranslateMatrix (pathRect.Left + pathRect.Width / 2) * 2, 0, MatrixOrderAppend
    If mirrorVertical Then tmpMatrix.TranslateMatrix 0, (pathRect.Top + pathRect.Height / 2) * 2, MatrixOrderAppend
    
    'Apply the transformation
    GdipTransformPath m_PathHandle, tmpMatrix.getMatrixHandle
    
End Sub

' END: NON-DESTRUCTIVE PATH EDITS
'
'***************************************************************************





'***************************************************************************
'
' BEGIN: DESTRUCTIVE PATH EDITS

'Convert a path to border type, using a single border width parameters as the guide.
Public Sub convertPathToBorderType_Simple(ByVal borderWidth As Single, Optional ByVal useLineCap As LineCap = LineCapFlat, Optional ByVal useLineJoin As LineJoin = LineJoinMiter, Optional ByVal allowableError As Single = FlatnessDefault)

    'Widening a path requires a GDI+ pen.  All of the pen's properties (width, dashing, transforms, miters, etc)
    ' are considered when applying the widening transform.
    Dim hPen As Long
    GdipCreatePen1 0, borderWidth, gdipUnitPixel, hPen
    
    'Apply the requested line cap and miter
    GdipSetPenLineCap hPen, useLineCap, useLineCap, DashCapFlat
    GdipSetPenLineJoin hPen, useLineJoin
    
    'Apply the widen command
    GdipWidenPath m_PathHandle, hPen, 0, allowableError
    
    'Release the temporary pen
    GdipDeletePen hPen

End Sub

'Inflate the current path by some arbitrary amount.  Note that additional parameters are available to control the way junctions and endpoints
' are affected by the widening.
'
'Unlike other subs, this actually returns a success/failure result.  GDI+ is prone to unpredictable failures, so the caller may need to take
' precautions if the request fails.  (On failure, this sub will automatically try to restore the original path if it can.)
Public Function inflatePath(ByVal inflateWidth As Single, Optional ByVal useLineCap As LineCap = LineCapFlat, Optional ByVal useLineJoin As LineJoin = LineJoinMiter, Optional ByVal allowableError As Single = FlatnessDefault) As Boolean

    'Widening a path requires a GDI+ pen.  All of the pen's properties (width, dashing, transforms, miters, etc)
    ' are considered when applying the widening transform.
    Dim hPen As Long
    GdipCreatePen1 0, inflateWidth, gdipUnitPixel, hPen
    
    'Because GDI+ widening transforms are prone to failure, make a backup of the current path (as we may need to restore it).
    Dim pathBackup As Long
    GdipClonePath m_PathHandle, pathBackup
    
    'Apply the requested line cap and miter
    GdipSetPenLineCap hPen, useLineCap, useLineCap, DashCapFlat
    GdipSetPenLineJoin hPen, useLineJoin
        
    'Apply the widen command
    If GdipWidenPath(m_PathHandle, hPen, 0, allowableError) = 0 Then
        
        'Re-clone the path in its current state
        GdipDeletePath pathBackup
        GdipClonePath m_PathHandle, pathBackup
        
        'Convert the path to an outline-only representation
        If GdipWindingModeOutline(m_PathHandle, 0, allowableError) = 0 Then
            GdipDeletePath pathBackup
            inflatePath = True
        Else
            GdipClonePath pathBackup, m_PathHandle
            GdipDeletePath pathBackup
            inflatePath = False
        End If
        
    Else
        GdipClonePath pathBackup, m_PathHandle
        GdipDeletePath pathBackup
        inflatePath = False
    End If
    
    'Regardless of outcome, we can now safely release our temporary pen
    GdipDeletePen hPen

End Function

' END: DESTRUCTIVE PATH EDITS
'
'***************************************************************************



'***************************************************************************
'
' BEGIN: EXTERNAL PATH ACCESS

Public Function getCopyOfPathHandle() As Long
    getCopyOfPathHandle = m_PathHandle
End Function

Public Sub cloneExistingPath(ByRef srcPath As pdGraphicsPath)
    
    'Make sure the source path exists!
    If srcPath Is Nothing Then
        Debug.Print "WARNING!  cloneExistingPath was passed a non-existent path - FIX THIS!"
        Exit Sub
    End If
    
    'If we already have a path, delete it
    If m_PathHandle <> 0 Then GdipDeletePath m_PathHandle
    
    'Clone the new path
    GdipClonePath srcPath.getCopyOfPathHandle, m_PathHandle
    
End Sub

Friend Sub getPathBounds(ByRef dstRect As RECTF)
    GdipGetPathWorldBounds m_PathHandle, dstRect, 0, 0
End Sub

' END: EXTERNAL PATH ACCESS
'
'***************************************************************************




'***************************************************************************
'
' BEGIN: BASIC MAINTENANCE FUNCTIONS

Private Sub Class_Initialize()

    'Create a blank graphics path
    GdipCreatePath FillModeWinding, m_PathHandle

End Sub

Private Sub Class_Terminate()

    'Release our path handle
    If m_PathHandle <> 0 Then GdipDeletePath m_PathHandle

End Sub

'Reset the path to a blank state.  Note that this also resets the fill mode to ALTERNATE; we override this with PD's
' preferred default behavior (winding).
Public Sub resetPath()
    GdipResetPath m_PathHandle
    setPathFillMode FillModeWinding
End Sub

' END: BASIC MAINTENANCE FUNCTIONS
'
'***************************************************************************



'***************************************************************************
'
' BEGIN: HIT-DETECTION FUNCTIONS

Public Function isPointInsidePath(ByVal x As Single, ByVal y As Single) As Boolean
    
    Dim gdipReturn As Long, ptInside As Long
    gdipReturn = GdipIsVisiblePathPoint(m_PathHandle, x, y, 0&, ptInside)
    
    If gdipReturn = 0 Then
        isPointInsidePath = CBool(ptInside <> 0)
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdGraphicsPath.isPointInsidePath() failed to call GdipIsVisiblePathPoint(). Failure code " & gdipReturn & "."
        #End If
    End If
    
End Function

' END: HIT-DETECTION FUNCTIONS
'
'***************************************************************************


'***************************************************************************
'
' BEGIN: DRAWING FUNCTIONS

'Stroke the current path, using the default arrangement for PD's current theme
Public Sub strokePathToDIB_UIStyle(ByRef dstDIB As pdDIB, Optional ByVal useThisDCInstead As Long = 0, Optional ByVal useHighlightColor As Boolean = False, Optional ByVal useHQOffsets As Boolean = True, Optional ByVal customLinejoin As LineJoin = LineJoinRound, Optional ByVal useCustomStartCap As LineCap = LineCapRound, Optional ByVal useCustomEndCap As LineCap = LineCapRound)

    'PD's default behavior for UI-related paths is to stroke a path twice; once with a thickness of 2px
    ' (which actually ends up being more like 1px along the path, and 0.5 px along either side)
    ' using a black pen at 75% opacity, then again at a thickness of 1px at 75% opacity.  This provides
    ' a sharp line regardless of the underlying image.
    
    'Antialiasing is always used.  Note that this is relevant even for shapes like rectangles, in order
    ' to split the 2px line across 3 pixels.
    
    'Obtain a GDI+ handle to the image
    Dim hGraphics As Long
    
    If useThisDCInstead <> 0 Then
        GdipCreateFromHDC useThisDCInstead, hGraphics
    Else
        GdipCreateFromHDC dstDIB.getDIBDC, hGraphics
    End If
    
    'Activate AA
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    'Activate half-pixel offsets
    If useHQOffsets Then
        GdipSetPixelOffsetMode hGraphics, PixelOffsetModeHighQuality
    Else
        GdipSetPixelOffsetMode hGraphics, PixelOffsetModeHighSpeed
    End If
    
    'Create two pens, each of which will be used to stroke the line
    Dim hBlackPen As Long, hWhitePen As Long
    GdipCreatePen1 fillQuadWithVBRGB(RGB(0, 0, 0), 192), 3#, gdipUnitPixel, hBlackPen
    
    If useHighlightColor Then
        GdipCreatePen1 fillQuadWithVBRGB(g_Themer.getThemeColor(PDTC_ACCENT_HIGHLIGHT), 220), 1.6, gdipUnitPixel, hWhitePen
    Else
        GdipCreatePen1 fillQuadWithVBRGB(RGB(255, 255, 255), 220), 1.6, gdipUnitPixel, hWhitePen
    End If
    
    'Populate line end and join parameters
    GdipSetPenLineCap hBlackPen, useCustomStartCap, useCustomEndCap, DashCapRound
    GdipSetPenLineCap hWhitePen, useCustomStartCap, useCustomEndCap, DashCapRound
    
    GdipSetPenLineJoin hBlackPen, customLinejoin
    GdipSetPenLineJoin hWhitePen, customLinejoin
    
    'Stroke the paths in turn
    GdipDrawPath hGraphics, hBlackPen, m_PathHandle
    GdipDrawPath hGraphics, hWhitePen, m_PathHandle
        
    'Release the two pens and DIB handle
    GdipDeletePen hBlackPen
    GdipDeletePen hWhitePen
    GdipDeleteGraphics hGraphics
    
End Sub

'Stroke the current path, using a bare pen handle for the rendering.
Public Sub strokePathToDIB_BarePen(ByRef srcPen As Long, Optional ByRef dstDIB As pdDIB = Nothing, Optional ByVal useThisDCInstead As Long = 0, Optional ByVal useAA As Boolean = True)
    
    'Obtain a GDI+ handle to the image
    Dim hGraphics As Long
    
    If useThisDCInstead <> 0 Then
        GdipCreateFromHDC useThisDCInstead, hGraphics
    Else
        GdipCreateFromHDC dstDIB.getDIBDC, hGraphics
    End If
    
    'Activate AA
    If useAA Then
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    Else
        GdipSetSmoothingMode hGraphics, SmoothingModeNone
    End If
        
    'Stroke the path
    GdipDrawPath hGraphics, srcPen, m_PathHandle
        
    'Release the DIB handle.  The caller is responsible for freeing the pen.
    GdipDeleteGraphics hGraphics
    
End Sub

'Fill the current path, using a bare brush handle for the rendering.
Public Sub fillPathToDIB_BareBrush(ByVal srcBrush As Long, Optional ByRef dstDIB As pdDIB = Nothing, Optional ByVal useThisDCInstead As Long = 0, Optional ByVal useAA As Boolean = True, Optional ByVal customRenderingOriginX As Long = 0, Optional ByVal customRenderingOriginY As Long = 0)
    
    'Obtain a GDI+ handle to the image
    Dim hGraphics As Long
    
    If useThisDCInstead <> 0 Then
        GdipCreateFromHDC useThisDCInstead, hGraphics
    Else
        GdipCreateFromHDC dstDIB.getDIBDC, hGraphics
    End If
    
    'Activate AA
    If useAA Then
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    Else
        GdipSetSmoothingMode hGraphics, SmoothingModeNone
    End If
    
    'Set rendering origin
    GdipSetRenderingOrigin hGraphics, customRenderingOriginX, customRenderingOriginY
    
    'Fill the path
    GdipFillPath hGraphics, srcBrush, m_PathHandle
        
    'Release the DIB handle.  The caller is responsible for freeing the brush.
    GdipDeleteGraphics hGraphics
    
End Sub

' END: DRAWING FUNCTIONS
'
'***************************************************************************




'***************************************************************************
'
' BEGIN: MISC SUPPORT FUNCTIONS

'GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from a standard RGB()
' long and supplied alpha.  It's not a very efficient conversion, but I need it so infrequently that I don't really care.
Private Function fillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
    
    Dim dstQuad As RGBQUAD
    With dstQuad
        .Red = ExtractR(vbRGB)
        .Green = ExtractG(vbRGB)
        .Blue = ExtractB(vbRGB)
        .Alpha = alphaValue
    End With
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    fillQuadWithVBRGB = placeHolder.lngResult
    
End Function

' END: MISC SUPPORT FUNCTIONS
'
'***************************************************************************
