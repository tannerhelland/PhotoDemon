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
'Copyright 2014-2020 by Tanner Helland
'Created: 17/October/14 (though assembled from various parts written much earlier)
'Last updated: 18/June/16
'Last update: continued work on bringing this class into the pd2D fold
'Dependencies: pd2DTransform, for applying affine transformations to a path.
'
'This class is a VB6-friendly wrapper for the GDI+ GraphicsPath object.  It is not intended as a comprehensive interface;
' instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ path handle is created at class initialization, and released at class termination.  The path handle is
' persistent by design, so no functions are provided for recreating or deleting it.  If you need to start over, use
' the ResetPath() function.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

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
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToIntArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToLongArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As GP_Result

Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal pathFillMode As GP_FillMode, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result

'Unlike GDI+ and DirectX, pd2DPath follows the Cairo convention and defaults to winding mode.
' You can see a visualization of what this means at this MSDN page:
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd368110(v=vs.85).aspx
' If you want the old behavior, you can change it here.
Private Const DEFAULT_FILL_RULE As Long = P2_FR_Winding

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
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
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
Friend Sub AddCircle(ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single)
    If (m_PathHandle = 0) Then CreatePath
    Me.AddEllipse_Absolute centerX - circleRadius, centerY - circleRadius, centerX + circleRadius, centerY + circleRadius
End Sub

Friend Sub AddEllipse_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, x1, y1, x2 - x1, y2 - y1
End Sub

Friend Sub AddEllipse_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, x1, y1, ellipseWidth, ellipseHeight
End Sub

Friend Sub AddEllipse_RectF(ByRef srcRect As RectF)
    If (m_PathHandle = 0) Then CreatePath
    With srcRect
        GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, .Left, .Top, .Width, .Height
    End With
End Sub

Friend Sub AddLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddLine m_PathHandle, x1, y1, x2, y2
End Sub

Friend Sub AddLineInt(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddLineI m_PathHandle, x1, y1, x2, y2
End Sub

Friend Sub AddLines(ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddLines m_PathHandle, ptrToPtFArray, numOfPoints
End Sub

Friend Sub AddLinesInt(ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddLinesI m_PathHandle, ptrToPtLArray, numOfPoints
End Sub

Friend Sub AddArc(ByRef ellipseBoundRect As RectF, ByVal startAngle As Single, ByVal sweepAngle As Single)
    If (m_PathHandle = 0) Then CreatePath
    With ellipseBoundRect
        GDI_Plus.GDIPlus_PathAddArc m_PathHandle, .Left, .Top, .Width, .Height, startAngle, sweepAngle
    End With
End Sub

Friend Sub AddArc_Absolute(ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x, y, arcWidth, arcHeight, startAngle, sweepAngle
End Sub

Friend Sub AddArcCircular(ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle
End Sub

Friend Sub AddPolygon(ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'We have a few different options for adding this shape, based on the passed parameters.
    If autoCloseShape Then
        If useCurveAlgorithm Then
            GDI_Plus.GDIPlus_PathAddClosedCurve m_PathHandle, ptrToPtFArray, numOfPoints, curvatureTension
        Else
            GDI_Plus.GDIPlus_PathAddPolygon m_PathHandle, ptrToPtFArray, numOfPoints
        End If
    Else
        If useCurveAlgorithm Then
            GDI_Plus.GDIPlus_PathAddCurve m_PathHandle, ptrToPtFArray, numOfPoints, curvatureTension
        Else
            GDI_Plus.GDIPlus_PathAddLines m_PathHandle, ptrToPtFArray, numOfPoints
        End If
    End If
    
End Sub

Friend Sub AddPolygonInt(ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'We have a few different options for adding this shape, based on the passed parameters.
    If autoCloseShape Then
        If useCurveAlgorithm Then
            GDI_Plus.GDIPlus_PathAddClosedCurveI m_PathHandle, ptrToPtLArray, numOfPoints, curvatureTension
        Else
            GDI_Plus.GDIPlus_PathAddPolygonI m_PathHandle, ptrToPtLArray, numOfPoints
        End If
    Else
        If useCurveAlgorithm Then
            GDI_Plus.GDIPlus_PathAddCurveI m_PathHandle, ptrToPtLArray, numOfPoints, curvatureTension
        Else
            GDI_Plus.GDIPlus_PathAddLinesI m_PathHandle, ptrToPtLArray, numOfPoints
        End If
    End If
    
End Sub

'Technically this should be called "AddPolygon_RegularConvex", but for brevity, we use just "regular".  Regular polygons
' (https://en.wikipedia.org/wiki/Regular_polygon) are the ones you learned in grade school: square, pentagon, hexagon, etc.
' Regular polygons are defined by a center point, radius, and number of sides.  This function also supports curvature.
Friend Sub AddPolygon_Regular(ByVal numOfSides As Long, ByVal pRadius As Single, Optional ByVal centerX As Single = 0#, Optional ByVal centerY As Single = 0#, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    
    If (numOfSides > 0) Then
        
        If (m_PathHandle = 0) Then CreatePath
        
        'Regardless of backend, we start by establishing the polygon's points using basic geometry.
        Dim polyPoints() As PointFloat
        ReDim polyPoints(0 To numOfSides - 1) As PointFloat
        
        Dim i As Long
        For i = 0 To numOfSides - 1
            polyPoints(i).x = centerX + pRadius * Cos(PI_DOUBLE * CSng(i) / CSng(numOfSides))
            polyPoints(i).y = centerY + pRadius * Sin(PI_DOUBLE * CSng(i) / CSng(numOfSides))
        Next i
        
        'I don't know if there is a performance difference between adding closed curves with tension = 0 vs
        ' adding plain polygons, but just in case there is, we split handling of these two types.
        If useCurveAlgorithm Then
            GDI_Plus.GDIPlus_PathAddClosedCurve m_PathHandle, VarPtr(polyPoints(0)), numOfSides, curvatureTension
        Else
            GDI_Plus.GDIPlus_PathAddPolygon m_PathHandle, VarPtr(polyPoints(0)), numOfSides
        End If
        
    End If
    
End Sub

Friend Sub AddRectangle_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, x1, y1, x2 - x1, y2 - y1
End Sub

Friend Sub AddRectangle_AbsoluteI(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddRectangleI m_PathHandle, x1, y1, x2 - x1, y2 - y1
End Sub

Friend Sub AddRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, x1, y1, rectWidth, rectHeight
End Sub

Friend Sub AddRectangle_RectF(ByRef srcRect As RectF)
    If (m_PathHandle = 0) Then CreatePath
    With srcRect
        GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, .Left, .Top, .Width, .Height
    End With
End Sub

Friend Sub AddRoundedRectangle_RectF(ByRef srcRect As RectF, ByVal cornerRadius As Single)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If (xCurvature > srcRect.Width) Then xCurvature = srcRect.Width
    If (yCurvature > srcRect.Height) Then yCurvature = srcRect.Height
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, srcRect.Left + srcRect.Width - xCurvature, srcRect.Top, xCurvature, yCurvature, 270#, 90#
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, srcRect.Left + srcRect.Width - xCurvature, srcRect.Top + srcRect.Height - yCurvature, xCurvature, yCurvature, 0#, 90#
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, srcRect.Left, srcRect.Top + srcRect.Height - yCurvature, xCurvature, yCurvature, 90#, 90#
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, srcRect.Left, srcRect.Top, xCurvature, yCurvature, 180#, 90#
    GDI_Plus.GDIPlus_PathCloseFigure m_PathHandle
    
End Sub

Friend Sub AddRoundedRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If (xCurvature > rectWidth) Then xCurvature = rectWidth
    If (yCurvature > rectHeight) Then yCurvature = rectHeight
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1 + rectWidth - xCurvature, y1, xCurvature, yCurvature, 270, 90
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1 + rectWidth - xCurvature, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 0, 90
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 90, 90
    GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1, y1, xCurvature, yCurvature, 180, 90
    GDI_Plus.GDIPlus_PathCloseFigure m_PathHandle
    
End Sub

Friend Sub AddBezierCurve(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathAddBezier m_PathHandle, x1, y1, x2, y2, x3, y3, x4, y4
End Sub

'Convenience wrapper for triangles
Friend Sub AddTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)
    
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
    
    Me.AddPolygon 3, VarPtr(tmpPoints(0)), True, False
    
End Sub

Friend Sub AddPath(ByRef srcGraphicsPath As pd2DPath, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0, Optional ByVal connectToLastPointOfThisPath As Boolean = False)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Make sure the glyph exists
    If (Not srcGraphicsPath Is Nothing) Then
        
        'If no offsets are specified, copy the path as-is
        If (offsetX = 0) And (offsetY = 0) Then
            GDI_Plus.GDIPlus_PathAddPath m_PathHandle, srcGraphicsPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
            
        'If offsets were specified, we need to clone the path, translate it, then add it
        Else
        
            'Clone the path
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.CloneExistingPath srcGraphicsPath
            
            'Translate the path
            tmpPath.TranslatePath offsetX, offsetY
            
            'Add it
            GDI_Plus.GDIPlus_PathAddPath m_PathHandle, tmpPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
            
        End If
        
    Else
        InternalError "AddPath", "null source"
    End If
    
End Sub

Friend Sub StartNewFigure()
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathStartFigure m_PathHandle
End Sub

Friend Sub CloseCurrentFigure()
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathCloseFigure m_PathHandle
End Sub

'This is a little silly, but this function creates a squiggly line inside the target rectangle.  It provides a nice shorthand
' method for previewing something like a pen against an arbitrary path.
Friend Sub CreateSamplePathForRect(ByRef srcRect As RectF, Optional ByVal hPadding As Single = 0, Optional ByVal vPadding As Single = 0)
    
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
    Me.AddPolygon 9, VarPtr(samplePoints(0)), False, True, 0.5

End Sub

'Get/Set the path's current fill rule.  (Setting a new fill rule is non-destructive, FYI.)
Friend Function GetFillRule() As PD_2D_FillRule
    If (m_PathHandle = 0) Then
        GetFillRule = DEFAULT_FILL_RULE
    Else
        GetFillRule = GDI_Plus.GDIPlus_PathGetFillRule(m_PathHandle)
    End If
End Function

Friend Sub SetFillRule(ByVal newRule As PD_2D_FillRule)
    If (m_PathHandle = 0) Then CreatePath
    GDI_Plus.GDIPlus_PathSetFillRule m_PathHandle, newRule
End Sub

'Translate the entire path by some amount in the x and/or y positions
Friend Function TranslatePath(ByVal xAmount As Single, ByVal yAmount As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    m_Transform.ApplyTranslation xAmount, yAmount
    
    TranslatePath = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    
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
        ApplyTransformation = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, srcTransform.GetHandle)
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
    RotatePathAroundItsCenter = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    
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
    MirrorPathAroundItsCenter = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    
End Function

'Use this function to strip any interior paths and leave just the outline of the shape.  Winding mode rules are used,
' so discrete subpaths will be "outlined" individually.
'
'Also, this function requires an allowable error parameter because any curves in the image will be flattened
' (e.g. converted to a series of straight lines that approximate the original curve).
Friend Function ConvertPath_OutlineOnly(Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    If (m_PathHandle <> 0) Then
        ConvertPath_OutlineOnly = GDI_Plus.GDIPlus_PathWindingModeOutline(m_PathHandle, 0&, allowableError)
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
            ConvertPath_PenTrace = GDI_Plus.GDIPlus_PathWiden(m_PathHandle, srcPen.GetHandle, 0&, allowableError)
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
    
        'Widening a path requires a GDI+ pen.  All of the pen's properties (width, dashing, transforms, miters, etc)
        ' are considered when applying the widening transform.
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, inflateWidth, , , useLineJoin, useLineCap

        'Because GDI+ widening transforms are prone to failure, make a backup of the current path (as we may need to restore it).
        Dim pathBackup As Long
        pathBackup = GDI_Plus.GDIPlus_PathClone(m_PathHandle)
            
        'Apply the widen command
        If GDI_Plus.GDIPlus_PathWiden(m_PathHandle, cPen.GetHandle, 0&, allowableError) Then
            
            'Re-clone the path in its current state
            GDI_Plus.GDIPlus_PathDelete pathBackup
            pathBackup = GDI_Plus.GDIPlus_PathClone(m_PathHandle)
            
            'Convert the path to an outline-only representation
            If GDI_Plus.GDIPlus_PathWindingModeOutline(m_PathHandle, 0, allowableError) Then
                ConvertPath_InflateLikeBalloon = True
            Else
                m_PathHandle = GDI_Plus.GDIPlus_PathClone(pathBackup)
                ConvertPath_InflateLikeBalloon = False
            End If
            
            GDI_Plus.GDIPlus_PathDelete pathBackup
            
        Else
            m_PathHandle = GDI_Plus.GDIPlus_PathClone(pathBackup)
            GDI_Plus.GDIPlus_PathDelete pathBackup
            ConvertPath_InflateLikeBalloon = False
        End If
        
    Else
        InternalError "ConvertPath_InflateLikeBalloon", "null path"
    End If
    
End Function

Friend Function CloneExistingPath(ByRef srcPath As pd2DPath) As Boolean
    
    If (Not srcPath Is Nothing) Then
        
        If (m_PathHandle <> 0) Then Me.ReleasePath
        m_PathHandle = GDI_Plus.GDIPlus_PathClone(srcPath.GetHandle)
        CloneExistingPath = (m_PathHandle <> 0)
        
        If (CloneExistingPath And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyPathCountChange True
    
    Else
        InternalError "CloneExistingPath", "null source path"
        CloneExistingPath = False
    End If
    
End Function

Friend Function DoesPointTouchPathOutlineF(ByVal x As Single, ByVal y As Single, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) And (Not srcPen Is Nothing) Then
        DoesPointTouchPathOutlineF = GDI_Plus.GDIPlus_PathDoesPointTouchOutlineF(m_PathHandle, x, y, srcPen.GetHandle)
    Else
        InternalError "DoesPointTouchPathOutlineF", "null hPath or hPen"
    End If
    
End Function

Friend Function DoesPointTouchPathOutlineL(ByVal x As Long, ByVal y As Long, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) And (Not srcPen Is Nothing) Then
        DoesPointTouchPathOutlineL = GDI_Plus.GDIPlus_PathDoesPointTouchOutlineL(m_PathHandle, x, y, srcPen.GetHandle)
    Else
        InternalError "DoesPointTouchPathOutlineL", "null hPath or hPen"
    End If
    
End Function

'Retrieve this path's boundaries.  Optionally, you can also receive the path's boundaries as if they were drawn with
' a particular pen.  (All pen settings are considered, including tricky things like corner mitering.)
Friend Function GetPathBoundariesF(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RectF
    
    If (m_PathHandle <> 0) Then
    
        Dim penHandle As Long, transformHandle As Long
        If (Not testPen Is Nothing) Then penHandle = testPen.GetHandle Else penHandle = 0
        If (Not testTransform Is Nothing) Then
            transformHandle = testTransform.GetHandle
        Else
            
            'If the caller specified a pen but *not* a transform, we need to create a temporary transform for them.
            If (penHandle <> 0) Then
                
                Dim tmpTransform As pd2DTransform
                Set tmpTransform = New pd2DTransform
                tmpTransform.Reset
                
                transformHandle = tmpTransform.GetHandle
                
            Else
                transformHandle = 0
            End If
            
        End If
        
        GetPathBoundariesF = GDI_Plus.GDIPlus_PathGetPathBoundsF(m_PathHandle, transformHandle, penHandle)
    
    Else
        InternalError "GetPathBoundariesF", "null hPath"
    End If
    
End Function

Friend Function GetPathBoundariesL(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RectL
        
    If (m_PathHandle <> 0) Then
    
        Dim penHandle As Long, transformHandle As Long
        If (Not testPen Is Nothing) Then penHandle = testPen.GetHandle Else penHandle = 0
        If (Not testTransform Is Nothing) Then transformHandle = testTransform.GetHandle Else transformHandle = 0
        
        GetPathBoundariesL = GDI_Plus.GDIPlus_PathGetPathBoundsL(m_PathHandle, transformHandle, penHandle)
        
    End If

End Function

Friend Function IsPointInsidePathF(ByVal x As Single, ByVal y As Single) As Boolean
    
    If (m_PathHandle <> 0) Then
        IsPointInsidePathF = GDI_Plus.GDIPlus_PathIsPointInsideF(m_PathHandle, x, y)
    Else
        InternalError "IsPointInsidePathF", "null hPath"
    End If
    
End Function

Friend Function IsPointInsidePathL(ByVal x As Long, ByVal y As Long) As Boolean
    
    If (m_PathHandle <> 0) Then
        IsPointInsidePathL = GDI_Plus.GDIPlus_PathIsPointInsideL(m_PathHandle, x, y)
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
