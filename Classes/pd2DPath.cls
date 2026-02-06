VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Path Wrapper
'Copyright 2014-2026 by Tanner Helland
'Created: 17/October/14 (though assembled from various parts written much earlier)
'Last updated: 06/May/22
'Last update: new GetPathPoints() function for retrieving this path (including all subpaths!) as a list of discrete lines
'Dependencies: pd2DTransform, for applying affine transformations to a path.
'
'This class is a VB6-friendly wrapper for the GDI+ GraphicsPath object.  It is not intended as a
' comprehensive interface; instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ path handle is created at class initialization, and released at class termination.
' The path handle is persistent by design, so no functions are provided for recreating or deleting it.
' If you need to start over, use the ResetPath() function.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When enumerating path points, you need to know what type of point each point is.
Private Enum GP_PathPointType
    gp_PathPointTypeStart = 0
    gp_PathPointTypeLine = 1
    gp_PathPointTypeBezier = 3
    gp_PathPointTypePathTypeMask = &H7&
    gp_PathPointTypeDashMode = &H10&
    gp_PathPointTypePathMarker = &H20&
    gp_PathPointTypeCloseSubpath = &H80&
End Enum

#If False Then
    Private Const gp_PathPointTypeStart = 0, gp_PathPointTypeLine = 1, gp_PathPointTypeBezier = 3, gp_PathPointTypePathTypeMask = &H7&, gp_PathPointTypeDashMode = &H10&, gp_PathPointTypePathMarker = &H20&, gp_PathPointTypeCloseSubpath = &H80&
#End If

Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As GP_Result
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GP_Result
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathCurve2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GP_Result
Private Declare Function GdipAddPathBeziers Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToIntArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As GP_Result
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal pathFillMode As GP_FillMode, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipFlattenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As GP_Result
Private Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByRef dstFillRule As GP_FillMode) As GP_Result
Private Declare Function GdipGetPointCount Lib "gdiplus" (ByVal hPath As Long, ByRef dstCount As Long) As GP_Result
Private Declare Function GdipGetPathPoints Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipGetPathTypes Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToByteArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RectF, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RectL, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal pathFillMode As GP_FillMode) As GP_Result
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal hPath As Long, ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As GP_Result
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformationMatrix As Long, ByVal allowableError As Single) As GP_Result

'Path iterators (not currently used)
'Private Declare Function GdipCreatePathIter Lib "gdiplus" (ByRef hIterator As Long, ByVal hPath As Long) As GP_Result
'Private Declare Function GdipDeletePathIter Lib "gdiplus" (ByVal hIterator As Long) As GP_Result
'Private Declare Function GdipPathIterNextSubpath Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByRef dstStartIndex As Long, ByRef dstEndIndex As Long, ByRef isClosedBool As Long) As GP_Result
'Private Declare Function GdipPathIterNextSubpathPath Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByRef dstHPath As Long, ByRef isClosedBool As Long) As GP_Result
'Private Declare Function GdipPathIterNextPathType Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByRef dstPathType As Byte, ByRef dstStartIndex As Long, ByRef dstEndIndex As Long) As GP_Result
'Private Declare Function GdipPathIterNextMarker Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByRef dstStartIndex As Long, ByRef dstEndIndex As Long) As GP_Result
'Private Declare Function GdipPathIterNextMarkerPath Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByRef dstHPath As Long) As GP_Result
'Private Declare Function GdipPathIterGetCount Lib "gdiplus" (ByVal hIterator As Long, ByRef dstCount As Long) As GP_Result
'Private Declare Function GdipPathIterGetSubpathCount Lib "gdiplus" (ByVal hIterator As Long, ByRef dstCount As Long) As GP_Result
'Private Declare Function GdipPathIterIsValid Lib "gdiplus" (ByVal hIterator As Long, ByRef dstIsValidBool As Long) As GP_Result
'Private Declare Function GdipPathIterHasCurve Lib "gdiplus" (ByVal hIterator As Long, ByRef dstHasCurveBool As Long) As GP_Result
'Private Declare Function GdipPathIterRewind Lib "gdiplus" (ByVal hIterator As Long) As GP_Result
'Private Declare Function GdipPathIterEnumerate Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByVal ptrToDstPointFs As Long, ByVal ptrToDstTypeBytes As Long, ByVal numDstPointCount As Long) As GP_Result
'Private Declare Function GdipPathIterCopyData Lib "gdiplus" (ByVal hIterator As Long, ByRef dstResultCount As Long, ByVal ptrToDstPointFs As Long, ByVal ptrToDstTypeBytes As Long, ByVal startIndex As Long, ByVal endIndex As Long) As GP_Result

'Unlike GDI+ and DirectX, pd2DPath follows the Cairo convention and defaults to winding mode.
' You can see a visualization of what this means at this MSDN page:
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd368110(v=vs.85).aspx
' If you want the old behavior, you can change it here.
' (Note: this is treated as a constant, set manually at Class_Initialize because VB doesn't let you declare constants
' of custom types publicly declared elsewhere.)
Private DEFAULT_FILL_RULE As PD_2D_FillRule

'Allowable error (in pixels) when flattening a path.  This value is poorly explained on MSDN,
' but lower values equal a more precise approximation.
' See http://msdn.microsoft.com/en-us/library/ms535572%28v=vs.85%29.aspx
' The default value of 0.25 is identical to the default value in GdiPlusEnums.h.
Private Const DEFAULT_APPROXIMATION_ERROR As Single = 0.25!

'The handle to our GDI+ path object.  This handle is automatically created with the class,
' and released when the class is terminated.  The caller shouldn't worry about it.
Private m_PathHandle As Long

'As a convenience, this class exposes some basic "transform" functions that don't require
' the caller to supply their own pd2DTransform object.  To apply said transformations,
' we use an internal instance of that class.
Private m_Transform As pd2DTransform

Friend Function GetHandle() As Long
    GetHandle = m_PathHandle
End Function

Friend Function HasPath() As Boolean
    HasPath = (m_PathHandle <> 0)
End Function

'Create an actual transform handle using the current backend and the current transform settings.
' NOTE: by design, this function is not exposed externally, because the caller *never* needs to call this directly.
'       If GetTransformHandle is called and transform doesn't yet exist, it will be auto-created (using an
'       identity matrix).  Similarly, if a matrix operation is applied to this class but the base matrix doesn't
'       exist, it will also be auto-created.
Private Function CreatePath() As Boolean

    If (m_PathHandle <> 0) Then Me.ReleasePath
    CreatePath = (GdipCreatePath(DEFAULT_FILL_RULE, m_PathHandle) = GP_OK)
    
    'When debug mode is active, all object creations are reported back to the central Drawing2D module
    CreatePath = CreatePath And (m_PathHandle <> 0)
    If (CreatePath And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyPathCountChange True
    
End Function

'Free the current path handle.
Friend Function ReleasePath() As Boolean
    
    ReleasePath = True
    
    If (m_PathHandle <> 0) Then
        
        ReleasePath = (GdipDeletePath(m_PathHandle) = GP_OK)
        
        'After a successful release, we must always reset the class-level handle to match,
        ' and during debug mode, the central Drawing2D module also needs to be notified.
        If ReleasePath Then
            m_PathHandle = 0
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyPathCountChange False
        End If
        
    End If
    
End Function

'Reset the path to a blank state.  Note that this may also reset the fill mode to a backend-specific value; for consistency,
' we override this with PD's currently specified default fill rule (which is WINDING, although the coder can change it using
' the constant at the top of this class).
Friend Function ResetPath() As Boolean
    ResetPath = True
    If (m_PathHandle <> 0) Then
        ResetPath = (GdipResetPath(m_PathHandle) = GP_OK)
        Me.SetFillRule DEFAULT_FILL_RULE
    End If
End Function

'Circles are just a subset of ellipses, so we don't include any circle-specific backend code here
Friend Function AddCircle(ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddCircle = Me.AddEllipse_Absolute(centerX - circleRadius, centerY - circleRadius, centerX + circleRadius, centerY + circleRadius)
End Function

Friend Function AddEllipse_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddEllipse_Absolute = (GdipAddPathEllipse(m_PathHandle, x1, y1, x2 - x1, y2 - y1) = GP_OK)
    If (Not AddEllipse_Absolute) Then InternalError "AddEllipse_Absolute", "GDI+ failure"
End Function

Friend Function AddEllipse_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddEllipse_Relative = (GdipAddPathEllipse(m_PathHandle, x1, y1, ellipseWidth, ellipseHeight) = GP_OK)
    If (Not AddEllipse_Relative) Then InternalError "AddEllipse_Relative", "GDI+ failure"
End Function

Friend Function AddEllipse_RectF(ByRef srcRect As RectF) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddEllipse_RectF = (GdipAddPathEllipse(m_PathHandle, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height) = GP_OK)
    If (Not AddEllipse_RectF) Then InternalError "AddEllipse_RectF", "GDI+ failure"
End Function

Friend Function AddLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddLine = (GdipAddPathLine(m_PathHandle, x1, y1, x2, y2) = GP_OK)
    If (Not AddLine) Then InternalError "AddLine", "GDI+ failure"
End Function

Friend Function AddLineInt(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddLineInt = (GdipAddPathLineI(m_PathHandle, x1, y1, x2, y2) = GP_OK)
    If (Not AddLineInt) Then InternalError "AddLineInt", "GDI+ failure"
End Function

Friend Function AddLines(ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddLines = (GdipAddPathLine2(m_PathHandle, ptrToPtFArray, numOfPoints) = GP_OK)
    If (Not AddLines) Then InternalError "AddLines", "GDI+ failure"
End Function

Friend Function AddLinesInt(ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddLinesInt = (GdipAddPathLine2I(m_PathHandle, ptrToPtLArray, numOfPoints) = GP_OK)
    If (Not AddLinesInt) Then InternalError "AddLinesInt", "GDI+ failure"
End Function

Friend Function AddArc(ByRef ellipseBoundRect As RectF, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddArc = (GdipAddPathArc(m_PathHandle, ellipseBoundRect.Left, ellipseBoundRect.Top, ellipseBoundRect.Width, ellipseBoundRect.Height, startAngle, sweepAngle) = GP_OK)
    If (Not AddArc) Then InternalError "AddArc", "GDI+ failure"
End Function

Friend Function AddArc_Absolute(ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddArc_Absolute = (GdipAddPathArc(m_PathHandle, x, y, arcWidth, arcHeight, startAngle, sweepAngle) = GP_OK)
    If (Not AddArc_Absolute) Then InternalError "AddArc_Absolute", "GDI+ failure"
End Function

Friend Function AddArcCircular(ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddArcCircular = (GdipAddPathArc(m_PathHandle, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
    If (Not AddArcCircular) Then InternalError "AddArcCircular", "GDI+ failure"
End Function

Friend Function AddPolygon(ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curveTension As Single = 0.5!) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'We have a few different options for adding this shape, based on the passed parameters.
    If autoCloseShape Then
        If useCurveAlgorithm Then
            AddPolygon = (GdipAddPathClosedCurve2(m_PathHandle, ptrToPtFArray, numOfPoints, curveTension) = GP_OK)
        Else
            AddPolygon = (GdipAddPathPolygon(m_PathHandle, ptrToPtFArray, numOfPoints) = GP_OK)
        End If
    Else
        If useCurveAlgorithm Then
            AddPolygon = (GdipAddPathCurve2(m_PathHandle, ptrToPtFArray, numOfPoints, curveTension) = GP_OK)
        Else
            AddPolygon = (GdipAddPathLine2(m_PathHandle, ptrToPtFArray, numOfPoints) = GP_OK)
        End If
    End If
    
    If (Not AddPolygon) Then InternalError "AddPolygon", "GDI+ failure"
    
End Function

Friend Function AddPolygonInt(ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curveTension As Single = 0.5!) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'We have a few different options for adding this shape, based on the passed parameters.
    If autoCloseShape Then
        If useCurveAlgorithm Then
            AddPolygonInt = (GdipAddPathClosedCurve2I(m_PathHandle, ptrToPtLArray, numOfPoints, curveTension) = GP_OK)
        Else
            AddPolygonInt = (GdipAddPathPolygonI(m_PathHandle, ptrToPtLArray, numOfPoints) = GP_OK)
        End If
    Else
        If useCurveAlgorithm Then
            AddPolygonInt = (GdipAddPathCurve2I(m_PathHandle, ptrToPtLArray, numOfPoints, curveTension) = GP_OK)
        Else
            AddPolygonInt = (GdipAddPathLine2I(m_PathHandle, ptrToPtLArray, numOfPoints) = GP_OK)
        End If
    End If
    
    If (Not AddPolygonInt) Then InternalError "AddPolygonInt", "GDI+ failure"
    
End Function

'Technically this should be called "AddPolygon_RegularConvex", but for brevity, we use just "regular".  Regular polygons
' (https://en.wikipedia.org/wiki/Regular_polygon) are the ones you learned in grade school: square, pentagon, hexagon, etc.
' Regular polygons are defined by a center point, radius, and number of sides.  This function also supports curvature.
Friend Function AddPolygon_Regular(ByVal numOfSides As Long, ByVal pRadius As Single, Optional ByVal centerX As Single = 0!, Optional ByVal centerY As Single = 0!, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curveTension As Single = 0.5!) As Boolean
    
    If (numOfSides > 0) Then
        
        If (m_PathHandle = 0) Then CreatePath
        
        'Regardless of backend, we start by establishing the polygon's points using basic geometry.
        Dim polyPoints() As PointFloat
        ReDim polyPoints(0 To numOfSides - 1) As PointFloat
        
        Dim i As Long, invNumSidesF As Single
        invNumSidesF = 1! / numOfSides
        For i = 0 To numOfSides - 1
            polyPoints(i).x = centerX + pRadius * Cos(PI_DOUBLE * CSng(i) * invNumSidesF)
            polyPoints(i).y = centerY + pRadius * Sin(PI_DOUBLE * CSng(i) * invNumSidesF)
        Next i
        
        'I don't know if there is a performance difference between adding closed curves with tension = 0 vs
        ' adding plain polygons, but just in case there is, we split handling of these two types.
        If useCurveAlgorithm Then
            AddPolygon_Regular = (GdipAddPathClosedCurve2(m_PathHandle, VarPtr(polyPoints(0)), numOfSides, curveTension) = GP_OK)
        Else
            AddPolygon_Regular = (GdipAddPathPolygon(m_PathHandle, VarPtr(polyPoints(0)), numOfSides) = GP_OK)
        End If
        
    End If
    
    If (Not AddPolygon_Regular) Then InternalError "AddPolygon_Regular", "GDI+ failure"
    
End Function

Friend Function AddRectangle_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddRectangle_Absolute = (GdipAddPathRectangle(m_PathHandle, x1, y1, x2 - x1, y2 - y1) = GP_OK)
    If (Not AddRectangle_Absolute) Then InternalError "AddRectangle_Absolute", "GDI+ failure"
End Function

Friend Function AddRectangle_AbsoluteI(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddRectangle_AbsoluteI = (GdipAddPathRectangleI(m_PathHandle, x1, y1, x2 - x1, y2 - y1) = GP_OK)
    If (Not AddRectangle_AbsoluteI) Then InternalError "AddRectangle_AbsoluteI", "GDI+ failure"
End Function

Friend Function AddRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddRectangle_Relative = (GdipAddPathRectangle(m_PathHandle, x1, y1, rectWidth, rectHeight) = GP_OK)
    If (Not AddRectangle_Relative) Then InternalError "AddRectangle_Relative", "GDI+ failure"
End Function

Friend Function AddRectangle_RectF(ByRef srcRect As RectF) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddRectangle_RectF = (GdipAddPathRectangle(m_PathHandle, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height) = GP_OK)
    If (Not AddRectangle_RectF) Then InternalError "AddRectangle_RectF", "GDI+ failure"
End Function

Friend Function AddRoundedRectangle_RectF(ByRef srcRect As RectF, ByVal cornerRadius As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If (xCurvature > srcRect.Width) Then xCurvature = srcRect.Width
    If (yCurvature > srcRect.Height) Then yCurvature = srcRect.Height
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    AddRoundedRectangle_RectF = (GdipAddPathArc(m_PathHandle, srcRect.Left + srcRect.Width - xCurvature, srcRect.Top, xCurvature, yCurvature, 270!, 90!) = GP_OK)
    AddRoundedRectangle_RectF = AddRoundedRectangle_RectF And (GdipAddPathArc(m_PathHandle, srcRect.Left + srcRect.Width - xCurvature, srcRect.Top + srcRect.Height - yCurvature, xCurvature, yCurvature, 0!, 90!) = GP_OK)
    AddRoundedRectangle_RectF = AddRoundedRectangle_RectF And (GdipAddPathArc(m_PathHandle, srcRect.Left, srcRect.Top + srcRect.Height - yCurvature, xCurvature, yCurvature, 90!, 90!) = GP_OK)
    AddRoundedRectangle_RectF = AddRoundedRectangle_RectF And (GdipAddPathArc(m_PathHandle, srcRect.Left, srcRect.Top, xCurvature, yCurvature, 180!, 90!) = GP_OK)
    AddRoundedRectangle_RectF = AddRoundedRectangle_RectF And (GdipClosePathFigure(m_PathHandle) = GP_OK)
    
    If (Not AddRoundedRectangle_RectF) Then InternalError "AddRoundedRectangle_RectF", "GDI+ failure"
    
End Function

Friend Function AddRoundedRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If (xCurvature > rectWidth) Then xCurvature = rectWidth
    If (yCurvature > rectHeight) Then yCurvature = rectHeight
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    AddRoundedRectangle_Relative = (GdipAddPathArc(m_PathHandle, x1 + rectWidth - xCurvature, y1, xCurvature, yCurvature, 270!, 90!) = GP_OK)
    AddRoundedRectangle_Relative = AddRoundedRectangle_Relative And (GdipAddPathArc(m_PathHandle, x1 + rectWidth - xCurvature, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 0!, 90!) = GP_OK)
    AddRoundedRectangle_Relative = AddRoundedRectangle_Relative And (GdipAddPathArc(m_PathHandle, x1, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 90!, 90!) = GP_OK)
    AddRoundedRectangle_Relative = AddRoundedRectangle_Relative And (GdipAddPathArc(m_PathHandle, x1, y1, xCurvature, yCurvature, 180!, 90!) = GP_OK)
    AddRoundedRectangle_Relative = AddRoundedRectangle_Relative And (GdipClosePathFigure(m_PathHandle) = GP_OK)
    
    If (Not AddRoundedRectangle_Relative) Then InternalError "AddRoundedRectangle_Relative", "GDI+ failure"
    
End Function

Friend Function AddBezierCurve(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddBezierCurve = (GdipAddPathBezier(m_PathHandle, x1, y1, x2, y2, x3, y3, x4, y4) = GP_OK)
    If (Not AddBezierCurve) Then InternalError "AddBezierCurve", "GDI+ failure"
End Function

Friend Function AddBezierCurveList(ByVal ptrToPointFs As Long, ByVal numPoints As Long) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    AddBezierCurveList = (GdipAddPathBeziers(m_PathHandle, ptrToPointFs, numPoints) = GP_OK)
    If (Not AddBezierCurveList) Then InternalError "AddBezierCurveList", "GDI+ failure"
End Function

'Convenience wrapper for triangles
Friend Function AddTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Convert the incoming points to a float array and use the generic polygon wrapper to add 'em
    Dim tmpPoints() As PointFloat
    ReDim tmpPoints(0 To 2) As PointFloat
    
    tmpPoints(0).x = x1
    tmpPoints(0).y = y1
    tmpPoints(1).x = x2
    tmpPoints(1).y = y2
    tmpPoints(2).x = x3
    tmpPoints(2).y = y3
    
    AddTriangle = Me.AddPolygon(3, VarPtr(tmpPoints(0)), True, False)
    
End Function

Friend Function AddPath(ByRef srcGraphicsPath As pd2DPath, Optional ByVal offsetX As Single = 0!, Optional ByVal offsetY As Single = 0!, Optional ByVal connectToLastPointOfThisPath As Boolean = False) As Boolean
    
    'Make sure the source glyph exists
    If (Not srcGraphicsPath Is Nothing) Then
        
        'Create a null path
        If (m_PathHandle = 0) Then CreatePath
        
        'If no offsets are specified, copy the path as-is
        If (offsetX = 0!) And (offsetY = 0!) Then
            AddPath = (GdipAddPathPath(m_PathHandle, srcGraphicsPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)) = GP_OK)
            
        'If offsets were specified, we need to clone the path, translate it, then add it
        Else
        
            'Clone the path
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.CloneExistingPath srcGraphicsPath
            
            'Translate the path
            tmpPath.TranslatePath offsetX, offsetY
            
            'Add it
            AddPath = (GdipAddPathPath(m_PathHandle, tmpPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)) = GP_OK)
            
        End If
        
        If (Not AddPath) Then InternalError "AddPath", "GDI+ failure"
        
    Else
        InternalError "AddPath", "null source"
    End If
    
End Function

Friend Function StartNewFigure() As Boolean
    If (m_PathHandle = 0) Then CreatePath
    StartNewFigure = (GdipStartPathFigure(m_PathHandle) = GP_OK)
    If (Not StartNewFigure) Then InternalError "StartNewFigure", "GDI+ failure"
End Function

Friend Function CloseCurrentFigure() As Boolean
    If (m_PathHandle = 0) Then CreatePath
    CloseCurrentFigure = (GdipClosePathFigure(m_PathHandle) = GP_OK)
    If (Not CloseCurrentFigure) Then InternalError "CloseCurrentFigure", "GDI+ failure"
End Function

'This is a little silly, but this function creates a squiggly line inside the target rectangle.
' It provides a nice shorthand method for previewing something like a pen against an arbitrary path.
Friend Function CreateSamplePathForRect(ByRef srcRect As RectF, Optional ByVal hPadding As Single = 0!, Optional ByVal vPadding As Single = 0!) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    
    'A sample path is just a nice little curve that demonstrates a few obvious path elements for the user
    Dim samplePoints() As PointFloat
    ReDim samplePoints(0 To 8) As PointFloat
    
    Dim sampleRect As RectF
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
    CreateSamplePathForRect = Me.AddPolygon(9, VarPtr(samplePoints(0)), False, True, 0.5!)

End Function

'Get/Set the path's current fill rule.  (Setting a new fill rule is non-destructive, FYI.)
Friend Function GetFillRule() As PD_2D_FillRule
    If (m_PathHandle = 0) Then
        GetFillRule = DEFAULT_FILL_RULE
    Else
        If (GdipGetPathFillMode(m_PathHandle, GetFillRule) <> GP_OK) Then InternalError "GetFillRule", "GDI+ failure"
    End If
End Function

Friend Function SetFillRule(ByVal newRule As PD_2D_FillRule) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    SetFillRule = (GdipSetPathFillMode(m_PathHandle, newRule) = GP_OK)
    If (Not SetFillRule) Then InternalError "SetFillRule", "GDI+ failure"
End Function

'Flatten the current path (e.g. reduce all arcs, curves, circles, etc to line approximations)
Friend Function FlattenPath(Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    If (m_PathHandle <> 0) Then
        FlattenPath = (GdipFlattenPath(m_PathHandle, 0&, allowableError) = GP_OK)
        If (Not FlattenPath) Then InternalError "FlattenPath", "GDI+ failure"
    Else
        InternalError "FlattenPath", "null path"
    End If
End Function

'Translate the entire path by some amount in the x and/or y positions
Friend Function TranslatePath(ByVal xAmount As Single, ByVal yAmount As Single) As Boolean
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    m_Transform.ApplyTranslation xAmount, yAmount
    TranslatePath = (GdipTransformPath(m_PathHandle, m_Transform.GetHandle) = GP_OK)
    If (Not TranslatePath) Then InternalError "TranslatePath", "GDI+ failure"
End Function

'Translate the entire path by some amount in polar coordinates (angle + radius).  Radius is in degrees.
Friend Function TranslatePath_Polar(ByVal translateAngle As Single, ByVal translateRadius As Single, Optional ByVal angleIsInDegrees As Boolean = True) As Boolean
    If angleIsInDegrees Then translateAngle = PDMath.DegreesToRadians(translateAngle)
    TranslatePath_Polar = Me.TranslatePath(translateRadius * Cos(translateAngle), translateRadius * Sin(translateAngle))
End Function

'Transform the path by some arbitrary pd2DTransform object
Friend Function ApplyTransformation(ByRef srcTransform As pd2DTransform) As Boolean
    If (Not srcTransform Is Nothing) Then
        If (m_PathHandle = 0) Then CreatePath
        ApplyTransformation = (GdipTransformPath(m_PathHandle, srcTransform.GetHandle) = GP_OK)
        If (Not ApplyTransformation) Then InternalError "ApplyTransformation", "GDI+ failure"
    Else
        InternalError "ApplyTransformation", "null transform"
    End If
End Function

'Rotate the path around its center point.  The center point is calculated automatically.
Friend Function RotatePathAroundItsCenter(ByVal rotateAngle As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    
    'Figure out the path's center
    Dim pathRect As RectF
    pathRect = Me.GetPathBoundariesF()
    
    'Update the transformation matrix with a matching rotation
    m_Transform.ApplyRotation rotateAngle, pathRect.Left + pathRect.Width / 2, pathRect.Top + pathRect.Height / 2, P2_TO_Append
    
    'Apply the transformation
    RotatePathAroundItsCenter = (GdipTransformPath(m_PathHandle, m_Transform.GetHandle) = GP_OK)
    If (Not RotatePathAroundItsCenter) Then InternalError "RotatePathAroundItsCenter ", "GDI+ failure"
    
End Function

'Mirror the path around its center point.  The center point is calculated automatically.
Friend Function MirrorPathAroundItsCenter(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    
    'Figure out the path's center
    Dim pathRect As RectF
    pathRect = Me.GetPathBoundariesF()
    
    'Update the transformation matrix with a matching mirror operation
    m_Transform.ApplyMirror mirrorHorizontal, mirrorVertical
    
    'Mirroring will reflect a path around the 0-axis, so we need to translate the path back into its original position now.
    If mirrorHorizontal Then m_Transform.ApplyTranslation (pathRect.Left + pathRect.Width / 2) * 2, 0, P2_TO_Append
    If mirrorVertical Then m_Transform.ApplyTranslation 0, (pathRect.Top + pathRect.Height / 2) * 2, P2_TO_Append
    
    'Apply the transformation
    MirrorPathAroundItsCenter = (GdipTransformPath(m_PathHandle, m_Transform.GetHandle) = GP_OK)
    If (Not MirrorPathAroundItsCenter) Then InternalError "MirrorPathAroundItsCenter ", "GDI+ failure"
    
End Function

'Use this function to strip any interior paths and leave just the outline of the shape.  Winding mode rules are used,
' so discrete subpaths will be "outlined" individually.
'
'Also, this function requires an allowable error parameter because any curves in the image will be flattened
' (e.g. converted to a series of straight lines that approximate the original curve).
Friend Function ConvertPath_OutlineOnly(Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    If (m_PathHandle <> 0) Then
        ConvertPath_OutlineOnly = (GdipWindingModeOutline(m_PathHandle, 0&, allowableError) = GP_OK)
        If (Not ConvertPath_OutlineOnly) Then InternalError "ConvertPath_OutlineOnly", "GDI+ failure"
    Else
        InternalError "ConvertPath_OutlineOnly", "null hPath"
    End If
End Function

'This (poorly named?) function converts the current path to a new path, as if the path were traced by the supplied pen.
' All of the pen's settings (including things like line ends, joins, dash patterns, and more!) are considered when creating
' the new path, so this transform is potentially very powerful.
'
'Prior to conversion, any curves in the path will be converted to a series of lines.  The "allowable error" value determines
' how closely those lines approximate the original curve; a higher error means fewer lines and a faster transform, while a
' lower error means a more perfect approximation (at some cost to performance).
Friend Function ConvertPath_PenTrace(ByRef srcPen As pd2DPen, Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    
    If (m_PathHandle <> 0) Then
        If (Not srcPen Is Nothing) Then
            ConvertPath_PenTrace = (GdipWidenPath(m_PathHandle, srcPen.GetHandle, 0&, allowableError) = GP_OK)
            If (Not ConvertPath_PenTrace) Then InternalError "ConvertPath_PenTrace", "GDI+ failure"
        Else
            ConvertPath_PenTrace = False
        End If
    Else
        InternalError "ConvertPath_PenTrace", "null hPath"
    End If
    
End Function

'Inflate the current path by some arbitrary amount.  Note that additional parameters are available to control the way junctions and endpoints
' are affected by the widening.
'
'Unlike other subs, this actually returns a success/failure result.  GDI+ is prone to unpredictable failures, so the caller may need to take
' precautions if the request fails.  (On failure, this sub will automatically try to restore the original path if it can.)
Friend Function ConvertPath_InflateLikeBalloon(ByVal inflateWidth As Single, Optional ByVal useLineJoin As PD_2D_LineJoin = P2_LJ_Round, Optional ByVal useLineCap As PD_2D_LineCap = P2_LC_Round, Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    
    If (m_PathHandle <> 0) Then
        
        Dim gpResult As GP_Result
        
        'Widening a path requires a GDI+ pen.  All of the pen's properties (width, dashing, transforms, miters, etc)
        ' are considered when applying the widening transform.
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenWidth inflateWidth
        cPen.SetPenLineJoin useLineJoin
        cPen.SetPenLineCap useLineCap
        
        'Because GDI+ widening transforms are prone to failure, make a backup of the current path (as we may need to restore it).
        Dim pathBackup As Long
        gpResult = GdipClonePath(m_PathHandle, pathBackup)
        If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipClonePath", gpResult
        
        'Apply the widen command
        gpResult = GdipWidenPath(m_PathHandle, cPen.GetHandle, 0&, allowableError)
        If (gpResult = GP_OK) Then
            
            'Re-clone the path in its current state
            gpResult = GdipDeletePath(pathBackup)
            If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipDeletePath", gpResult
            gpResult = GdipClonePath(m_PathHandle, pathBackup)
            If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipClonePath", gpResult
            
            'Convert the path to an outline-only representation, and if that fails, restore the original path
            gpResult = GdipWindingModeOutline(m_PathHandle, 0, allowableError)
            ConvertPath_InflateLikeBalloon = (gpResult = GP_OK)
            If (Not ConvertPath_InflateLikeBalloon) Then
                InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipWindingModeOutline", gpResult
                gpResult = GdipClonePath(pathBackup, m_PathHandle)
                If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipClonePath", gpResult
            End If
            
            'Delete the backup path before exiting
            gpResult = GdipDeletePath(pathBackup)
            If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipDeletePath", gpResult
            
        Else
            InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipWidenPath", gpResult
            gpResult = GdipClonePath(pathBackup, m_PathHandle)
            If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipClonePath", gpResult
            gpResult = GdipDeletePath(pathBackup)
            If (gpResult <> GP_OK) Then InternalError "ConvertPath_InflateLikeBalloon", "GDI+ failure: GdipDeletePath", gpResult
            ConvertPath_InflateLikeBalloon = False
        End If
        
    Else
        InternalError "ConvertPath_InflateLikeBalloon", "null path"
    End If
    
End Function

Friend Function CloneExistingPath(ByRef srcPath As pd2DPath) As Boolean
    
    If (Not srcPath Is Nothing) Then
        
        If (m_PathHandle <> 0) Then Me.ReleasePath
        CloneExistingPath = (GdipClonePath(srcPath.GetHandle, m_PathHandle) = GP_OK)
        CloneExistingPath = CloneExistingPath And (m_PathHandle <> 0)
        If CloneExistingPath Then
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyPathCountChange True
        Else
            InternalError "CloneExistingPath", "GDI+ failure"
        End If
    
    Else
        InternalError "CloneExistingPath", "null source path"
        CloneExistingPath = False
    End If
    
End Function

Friend Function DoesPointTouchPathOutlineF(ByVal x As Single, ByVal y As Single, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) And (Not srcPen Is Nothing) Then
        Dim tmpResult As GP_Result, tmpLong As Long
        tmpResult = GdipIsOutlineVisiblePathPoint(m_PathHandle, x, y, srcPen.GetHandle, 0&, tmpLong)
        If (tmpResult = GP_OK) Then
            DoesPointTouchPathOutlineF = (tmpLong <> 0)
        Else
            InternalError "DoesPointTouchPathOutlineF", "GDI+ failure", tmpResult
        End If
    Else
        InternalError "DoesPointTouchPathOutlineF", "null hPath or hPen"
    End If
    
End Function

Friend Function DoesPointTouchPathOutlineL(ByVal x As Long, ByVal y As Long, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) And (Not srcPen Is Nothing) Then
        Dim tmpResult As GP_Result, tmpLong As Long
        tmpResult = GdipIsOutlineVisiblePathPointI(m_PathHandle, x, y, srcPen.GetHandle, 0&, tmpLong)
        If (tmpResult = GP_OK) Then
            DoesPointTouchPathOutlineL = (tmpLong <> 0)
        Else
            InternalError "DoesPointTouchPathOutlineL", "GDI+ failure", tmpResult
        End If
    Else
        InternalError "DoesPointTouchPathOutlineL", "null hPath or hPen"
    End If
    
End Function

Friend Function GetNumPathPoints() As Long
    If (m_PathHandle <> 0) Then
        If (GdipGetPointCount(m_PathHandle, GetNumPathPoints) <> GP_OK) Then InternalError "GetNumPathPoints", "GDI+ failure"
    End If
End Function

'Retrieve this path's boundaries.  Optionally, you can also receive the path's boundaries as if they were drawn with
' a particular pen.  (All pen settings are considered, including tricky things like corner mitering.)
Friend Function GetPathBoundariesF(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RectF
    
    If (m_PathHandle <> 0) Then
    
        Dim hPen As Long, hTransform As Long
        If (Not testPen Is Nothing) Then hPen = testPen.GetHandle Else hPen = 0
        If (Not testTransform Is Nothing) Then
            hTransform = testTransform.GetHandle
        Else
            
            'If the caller specified a pen but *not* a transform, we need to create a temporary transform for them.
            If (hPen <> 0) Then
                Dim tmpTransform As pd2DTransform
                Set tmpTransform = New pd2DTransform
                tmpTransform.Reset
                hTransform = tmpTransform.GetHandle
            Else
                hTransform = 0
            End If
            
        End If
        
        Dim tmpReturn As GP_Result
        tmpReturn = GdipGetPathWorldBounds(m_PathHandle, GetPathBoundariesF, hTransform, hPen)
        If (tmpReturn <> GP_OK) Then InternalError "GetPathBoundariesF", "GDI+ failure"
        
    Else
        InternalError "GetPathBoundariesF", "null hPath"
    End If
    
End Function

Friend Function GetPathBoundariesL(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RectL
    
    If (m_PathHandle <> 0) Then
    
        Dim hPen As Long, hTransform As Long
        If (Not testPen Is Nothing) Then hPen = testPen.GetHandle Else hPen = 0
        If (Not testTransform Is Nothing) Then
            hTransform = testTransform.GetHandle
        Else
            
            'If the caller specified a pen but *not* a transform, we need to create a temporary transform for them.
            If (hPen <> 0) Then
                Dim tmpTransform As pd2DTransform
                Set tmpTransform = New pd2DTransform
                tmpTransform.Reset
                hTransform = tmpTransform.GetHandle
            Else
                hTransform = 0
            End If
            
        End If
        
        Dim tmpReturn As GP_Result
        tmpReturn = GdipGetPathWorldBoundsI(m_PathHandle, GetPathBoundariesL, hTransform, hPen)
        If (tmpReturn <> GP_OK) Then InternalError "GetPathBoundariesL", "GDI+ faulre"
        
    Else
        InternalError "GetPathBoundariesL", "null hPath"
    End If
End Function

'Retrieve the current path as a list of points (representing connected line segments).  YOU MUST READ AND
' UNDERSTAND THIS COMMENT BLOCK TO KNOW HOW TO INTERPRET THIS DATA.
'
'First, note that a temporary path must be created to make this operation lossless, because the source path
' needs to be flattened before we can retrieve straight-line coordinates - for example, an ellipse must be
' flattened into discrete line segments because this function (by design) only returns straight lines.
'
'Note that three destination arrays are required: one for the actual points, and two others (which share a
' size potentially much smaller than the points list) that tell you the indices of each unique subpath in
' the point list, and whether that subpath is open (false) or closed (true).  You can thus use these
' second and third lists to successfully traverse complex shapes, which may be composed of one or more
' separate open and/or closed paths.
Friend Function GetPathPoints(ByRef dstPoints() As PointFloat, ByRef dstNumPoints As Long, ByRef dstPathStartMarkers() As Long, ByRef dstIsPathClosed() As Boolean, ByRef dstNumPathMarkers As Long) As Boolean
    
    Const FUNC_NAME As String = "GetPathPoints"
    
    'Sanity check
    If (m_PathHandle = 0) Then Exit Function
    
    'We need to flatten the path before retrieving it, but this is a destructive operation.
    ' Clone the current path into a temporary path, flatten *that*, then use the flattened
    ' path for all subsequent retrieval calls.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    If (Not tmpPath.CloneExistingPath(Me)) Then
        InternalError FUNC_NAME, "bad clone"
        Exit Function
    End If
    
    If (Not tmpPath.FlattenPath()) Then
        InternalError FUNC_NAME, "bad flatten"
        Exit Function
    End If
    
    'Start by just pulling the path data from the path in GDI+ format.
    ' (We need to translate this into a more portable format, but one problem at a time.)
    Dim numGdipPoints As Long
    If (GdipGetPointCount(tmpPath.GetHandle, numGdipPoints) <> GP_OK) Then
        InternalError FUNC_NAME, "bad GdipPointCount call"
        Exit Function
    End If
    
    'Size two arrays to match: one for points, and another for point *types*.
    Dim srcGdipPoints() As PointFloat, srcGdipTypes() As Byte
    ReDim srcGdipPoints(0 To numGdipPoints - 1) As PointFloat
    ReDim srcGdipTypes(0 To numGdipPoints - 1) As Byte
    
    'Retrieve both arrays and fail if either retrieval fails (since we need both to interpret the data).
    If (GdipGetPathPoints(tmpPath.GetHandle, VarPtr(srcGdipPoints(0)), numGdipPoints) <> GP_OK) Then
        InternalError FUNC_NAME, "bad GdipGetPathPoints call"
        Exit Function
    End If
    
    If (GdipGetPathTypes(tmpPath.GetHandle, VarPtr(srcGdipTypes(0)), numGdipPoints) <> GP_OK) Then
        InternalError FUNC_NAME, "bad GdipGetPathTypes call"
        Exit Function
    End If
    
    'We now have everything we need from the temporary path.  Free it.
    Set tmpPath = Nothing
    
    'We now need to convert the original point list to a PhotoDemon-specific one.
    ' (This function is designed for internal use and it only ever wants line positions.  To simplify
    ' traversal, we don't want to hand off a huge list of GDI+ specific entries - instead, we want a
    ' simple array of point coordinates, and a second simple array of markers that tell us the indices
    ' where each sub-path begins.)
    '
    'We know that both our output lists will be the same size or smaller than the lists we retrieved
    ' from GDI+, thankfully.  (We'll size these arrays precisely after we fill them.)
    ReDim dstPoints(0 To numGdipPoints - 1) As PointFloat
    ReDim dstPathStartMarkers(0 To numGdipPoints - 1) As Long
    ReDim dstIsPathClosed(0 To numGdipPoints - 1) As Boolean
    
    dstNumPoints = 0
    dstNumPathMarkers = 0
    
    'Iterate the original GDI+ list and populate the output lists as we go
    Dim i As Long, ptType As GP_PathPointType
    For i = 0 To numGdipPoints - 1
        
        'Retrieve the GDI+ "type" of this point.  From the possible list of GDI+ entries, we only want
        ' to line endpoints to the destination list.
        ptType = srcGdipTypes(i)
        
        'Points that start a new subpath have the special "start" marker (basically, no line or curve
        ' indicator which means that by default it must be a "start of path" marker).
        If ((ptType And gp_PathPointTypePathTypeMask) = 0) Then
            dstPathStartMarkers(dstNumPathMarkers) = dstNumPoints
            dstPoints(dstNumPoints) = srcGdipPoints(i)
            dstNumPoints = dstNumPoints + 1
            dstNumPathMarkers = dstNumPathMarkers + 1
        
        Else
            
            'Only add line points, not curve points (they should have been flattened in the last step!)
            If ((ptType And gp_PathPointTypeLine) <> 0) Then
                dstPoints(dstNumPoints) = srcGdipPoints(i)
                dstNumPoints = dstNumPoints + 1
            End If
            
            'Look for "this subpath must be closed "markers.  When we see them, add a corresponding note
            ' in the "isPathClosed" array.
            If ((ptType And gp_PathPointTypeCloseSubpath) <> 0) Then
                dstIsPathClosed(dstNumPathMarkers - 1) = True
            End If
            
        End If
            
    Next i
    
    'Resize the destination arrays to match their final size(s)
    If (UBound(dstPoints) <> dstNumPoints - 1) Then ReDim Preserve dstPoints(0 To dstNumPoints - 1) As PointFloat
    If (UBound(dstPathStartMarkers) <> dstNumPathMarkers - 1) Then ReDim Preserve dstPathStartMarkers(0 To dstNumPathMarkers - 1) As Long
    If (UBound(dstIsPathClosed) <> dstNumPathMarkers - 1) Then ReDim Preserve dstIsPathClosed(0 To dstNumPathMarkers - 1) As Boolean
    
    GetPathPoints = True
    
End Function

Friend Function IsPointInsidePathF(ByVal x As Single, ByVal y As Single) As Boolean
    If (m_PathHandle <> 0) Then
        Dim tmpReturn As GP_Result, tmpResult As Long
        tmpReturn = GdipIsVisiblePathPoint(m_PathHandle, x, y, 0&, tmpResult)
        If (tmpReturn = GP_OK) Then
            IsPointInsidePathF = (tmpResult <> 0)
        Else
            InternalError "IsPointInsidePathF", "GDI+ failure"
        End If
    Else
        InternalError "IsPointInsidePathF", "null hPath"
    End If
End Function

Friend Function IsPointInsidePathL(ByVal x As Long, ByVal y As Long) As Boolean
    If (m_PathHandle <> 0) Then
        Dim tmpReturn As GP_Result, tmpResult As Long
        tmpReturn = GdipIsVisiblePathPointI(m_PathHandle, x, y, 0&, tmpResult)
        If (tmpReturn = GP_OK) Then
            IsPointInsidePathL = (tmpResult <> 0)
        Else
            InternalError "IsPointInsidePathL", "GDI+ failure"
        End If
    Else
        InternalError "IsPointInsidePathL", "null hPath"
    End If
End Function

'As a convenience, this class exposes some generic "transform" operations that don't require the user to supply their own
' pd2DTransform object.  Instead, we use an internal instance that we reset between calls.
Private Sub ResetInternalTransform()
    If (m_Transform Is Nothing) Then Set m_Transform = New pd2DTransform
    m_Transform.Reset
End Sub

Private Sub Class_Initialize()
    
    'Set your preferred default fill rule here
    DEFAULT_FILL_RULE = P2_FR_Winding
    
    Me.ResetPath
    
End Sub

Private Sub Class_Terminate()
    Me.ReleasePath
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DPath", errFunction, errDescription, errNum
End Sub
