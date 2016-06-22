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
'Copyright 2014-2016 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Unlike GDI+ and DirectX, pd2DPath follows the Cairo convention and defaults to winding mode.  You can see a visualization
' of what this means at this MSDN page: https://msdn.microsoft.com/en-us/library/windows/desktop/dd368110(v=vs.85).aspx
' If you want the old behavior, you can change it here.
Private Const DEFAULT_FILL_RULE As Long = P2_FR_Winding

'Allowable error (in pixels) when flattening a path.  This value is poorly explained on MSDN, but lower values equal a more
' precise approximation.  See http://msdn.microsoft.com/en-us/library/ms535572%28v=vs.85%29.aspx
' The default value of 0.25 is identical to the default value in GdiPlusEnums.h.
Private Const DEFAULT_APPROXIMATION_ERROR As Single = 0.25

'The backend used to manage this object.  At present, this is immutable; once set, it cannot be migrated.
Private m_PathBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'The handle to our GDI+ path object.  This handle is automatically created with the class, and released when the
' class is terminated.  The caller doesn't have to worry about it at all.
Private m_PathHandle As Long

'As a convenience, this class exposes some basic "transform" functions that don't require the caller to supply their own
' pd2DTransform object.  To apply said transformations, we use an internal instance of that class.
Private m_Transform As pd2DTransform

Public Function GetHandle() As Long
    GetHandle = m_PathHandle
End Function

Public Function GetPathBackend() As PD_2D_RENDERING_BACKEND
    GetPathBackend = m_PathBackend
End Function

Public Function HasPath() As Boolean
    HasPath = CBool(m_PathHandle <> 0)
End Function

'Create an actual transform handle using the current backend and the current transform settings.
' NOTE: by design, this function is not exposed externally, because the caller *never* needs to call this directly.
'       If GetTransformHandle is called and transform doesn't yet exist, it will be auto-created (using an
'       identity matrix).  Similarly, if a matrix operation is applied to this class but the base matrix doesn't
'       exist, it will also be auto-created.
Private Function CreatePath() As Boolean

    If (m_PathHandle <> 0) Then Me.ReleasePath
    
    'Call the backend-specific creation function
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            m_PathHandle = GDI_Plus.GDIPlus_PathCreate(DEFAULT_FILL_RULE)
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreatePath = CBool(m_PathHandle <> 0)
    If (CreatePath And m_DebugMode) Then Drawing2D.DEBUG_NotifyPathCountChange m_PathBackend, True
    
End Function

'Free the current path handle.
Public Function ReleasePath() As Boolean
    
    If (m_PathHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleasePath = GDI_Plus.GDIPlus_PathDelete(m_PathHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleasePath Then
            m_PathHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyPathCountChange m_PathBackend, False
        End If
        
    Else
        ReleasePath = False
    End If
    
End Function

'Reset the path to a blank state.  Note that this may also reset the fill mode to a backend-specific value; for consistency,
' we override this with PD's currently specified default fill rule (which is WINDING, although the coder can change it using
' the constant at the top of this class).
Public Sub ResetPath()

    If (m_PathHandle <> 0) Then
        
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GDI_Plus.GDIPlus_PathReset m_PathHandle
        End Select
        
        Me.SetFillRule DEFAULT_FILL_RULE
        
    End If
    
End Sub

'Circles are just a subset of ellipses, so we don't include any circle-specific backend code here
Public Sub AddCircle(ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single)
    If (m_PathHandle = 0) Then CreatePath
    Me.AddEllipse_Absolute centerX - circleRadius, centerY - circleRadius, centerX + circleRadius, centerY + circleRadius
End Sub

Public Sub AddEllipse_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, x1, y1, x2 - x1, y2 - y1
    End Select
End Sub

Public Sub AddEllipse_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, x1, y1, ellipseWidth, ellipseHeight
    End Select
End Sub

Friend Sub AddEllipse_RectF(ByRef srcRect As RECTF)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            With srcRect
                GDI_Plus.GDIPlus_PathAddEllipse m_PathHandle, .Left, .Top, .Width, .Height
            End With
    End Select
End Sub

Public Sub AddLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddLine m_PathHandle, x1, y1, x2, y2
    End Select
End Sub

Friend Sub AddArc(ByRef ellipseBoundRect As RECTF, ByVal startAngle As Single, ByVal sweepAngle As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            With ellipseBoundRect
                GDI_Plus.GDIPlus_PathAddArc m_PathHandle, .Left, .Top, .Width, .Height, startAngle, sweepAngle
            End With
    End Select
End Sub

Public Sub AddArcCircular(ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddArc m_PathHandle, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle
    End Select
End Sub

Public Sub AddPolygon(ByVal numOfPoints As Long, ByVal ptrToFloatArray As Long, ByVal autoCloseShape As Boolean, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    
    If (m_PathHandle = 0) Then CreatePath
    
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
    
            'We have a few different options for adding this shape, based on the passed parameters.
            If autoCloseShape Then
                If useCurveAlgorithm Then
                    GDI_Plus.GDIPlus_PathAddClosedCurve m_PathHandle, ptrToFloatArray, numOfPoints, curvatureTension
                Else
                    GDI_Plus.GDIPlus_PathAddPolygon m_PathHandle, ptrToFloatArray, numOfPoints
                End If
            Else
                If useCurveAlgorithm Then
                    GDI_Plus.GDIPlus_PathAddCurve m_PathHandle, ptrToFloatArray, numOfPoints, curvatureTension
                Else
                    GDI_Plus.GDIPlus_PathAddLines m_PathHandle, ptrToFloatArray, numOfPoints
                End If
            End If
            
    End Select
    
End Sub

Public Sub AddRectangle_Absolute(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, x1, y1, x2 - x1, y2 - y1
    End Select
End Sub

Public Sub AddRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, x1, y1, rectWidth, rectHeight
    End Select
End Sub

Friend Sub AddRectangle_RectF(ByRef srcRect As RECTF)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            With srcRect
                GDI_Plus.GDIPlus_PathAddRectangle m_PathHandle, .Left, .Top, .Width, .Height
            End With
    End Select
End Sub

Public Sub AddRoundedRectangle_Relative(ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Validate the radius twice before applying it.  The width and height curvature cannot be less than
    ' 1/2 the width (or height) of the rect.
    Dim xCurvature As Single, yCurvature As Single
    xCurvature = cornerRadius
    yCurvature = cornerRadius
    
    If (xCurvature > rectWidth) Then xCurvature = rectWidth
    If (yCurvature > rectHeight) Then yCurvature = rectHeight
    
    'Add four arcs, which are auto-connected by the path engine, then close the figure
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1 + rectWidth - xCurvature, y1, xCurvature, yCurvature, 270, 90
            GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1 + rectWidth - xCurvature, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 0, 90
            GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1, y1 + rectHeight - yCurvature, xCurvature, yCurvature, 90, 90
            GDI_Plus.GDIPlus_PathAddArc m_PathHandle, x1, y1, xCurvature, yCurvature, 180, 90
            GDI_Plus.GDIPlus_PathCloseFigure m_PathHandle
            
    End Select
    
End Sub

Public Sub AddBezierCurve(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathAddBezier m_PathHandle, x1, y1, x2, y2, x3, y3, x4, y4
    End Select
End Sub

'Convenience wrapper for triangles
Public Sub AddTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Convert the incoming points to a float array and use the generic polygon wrapper to add 'em
    Dim tmpPoints() As POINTFLOAT
    ReDim tmpPoints(0 To 2) As POINTFLOAT
    
    tmpPoints(0).x = x1
    tmpPoints(0).y = y1
    tmpPoints(1).x = x2
    tmpPoints(1).y = y2
    tmpPoints(2).x = x3
    tmpPoints(2).y = y3
    
    Me.AddPolygon 3, VarPtr(tmpPoints(0)), True, False
    
End Sub

Public Sub AddPath(ByRef srcGraphicsPath As pd2DPath, Optional ByVal offsetX As Single = 0, Optional ByVal offsetY As Single = 0, Optional ByVal connectToLastPointOfThisPath As Boolean = False)
    
    If (m_PathHandle = 0) Then CreatePath
    
    'Make sure the glyph exists
    If (Not srcGraphicsPath Is Nothing) Then
        
        'If no offsets are specified, copy the path as-is
        If (offsetX = 0) And (offsetY = 0) Then
            Select Case m_PathBackend
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    GDI_Plus.GDIPlus_PathAddPath m_PathHandle, srcGraphicsPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
            End Select
            
        'If offsets were specified, we need to clone the path, translate it, then add it
        Else
        
            'Clone the path
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.CloneExistingPath srcGraphicsPath
            
            'Translate the path
            tmpPath.TranslatePath offsetX, offsetY
            
            'Add it
            Select Case m_PathBackend
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    GDI_Plus.GDIPlus_PathAddPath m_PathHandle, tmpPath.GetHandle, IIf(connectToLastPointOfThisPath, 1&, 0&)
            End Select
            
        End If
        
    Else
        InternalError "Empty source path", "WARNING!  pd2DPath.AddPath was handed a non-existent path."
        Exit Sub
    End If
    
End Sub

Public Sub StartNewFigure()
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathStartFigure m_PathHandle
    End Select
End Sub

Public Sub CloseCurrentFigure()
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathCloseFigure m_PathHandle
    End Select
End Sub

'This is a little silly, but this function creates a squiggly line inside the target rectangle.  It provides a nice shorthand
' method for previewing something like a pen against an arbitrary path.
Friend Sub CreateSamplePathForRect(ByRef srcRect As RECTF, Optional ByVal hPadding As Single = 0, Optional ByVal vPadding As Single = 0)
    
    If (m_PathHandle = 0) Then CreatePath
    
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
    Me.AddPolygon 9, VarPtr(samplePoints(0)), False, True, 0.5

End Sub

'Get/Set the path's current fill rule.  (Setting a new fill rule is non-destructive, FYI.)
Public Function GetFillRule() As PD_2D_FillRule
    If (m_PathHandle = 0) Then
        GetFillRule = DEFAULT_FILL_RULE
    Else
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GetFillRule = GDI_Plus.GDIPlus_PathGetFillRule(m_PathHandle)
        End Select
    End If
End Function

Public Sub SetFillRule(ByVal newRule As PD_2D_FillRule)
    If (m_PathHandle = 0) Then CreatePath
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            GDI_Plus.GDIPlus_PathSetFillRule m_PathHandle, newRule
    End Select
End Sub

'Translate the entire path by some amount in the x and/or y positions
Public Function TranslatePath(ByVal xAmount As Single, ByVal yAmount As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    m_Transform.ApplyTranslation xAmount, yAmount
    
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            TranslatePath = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    End Select
    
End Function

'Transform the path by some arbitrary pd2DTransform object
Public Function ApplyTransformation(ByRef srcTransform As pd2DTransform) As Boolean
    If (Not srcTransform Is Nothing) Then
        If (m_PathHandle = 0) Then CreatePath
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ApplyTransformation = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
        End Select
    Else
        InternalError "Source transformation is empty", "ApplyTransformation requires a valid source transformation object!"
    End If
End Function

'Rotate the path around its center point.  The center point is calculated automatically.
Public Function RotatePathAroundItsCenter(ByRef rotateAngle As Single) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    
    'Figure out the path's center
    Dim pathRect As RECTF
    pathRect = Me.GetPathBoundariesF()
    
    'Update the transformation matrix with a matching rotation
    m_Transform.ApplyRotation rotateAngle, pathRect.Left + pathRect.Width / 2, pathRect.Top + pathRect.Height / 2, P2_TO_Append
    
    'Apply the transformation
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            RotatePathAroundItsCenter = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    End Select
    
End Function

'Mirror the path around its center point.  The center point is calculated automatically.
Public Function MirrorPathAroundItsCenter(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean) As Boolean
    
    If (m_PathHandle = 0) Then CreatePath
    ResetInternalTransform
    
    'Figure out the path's center
    Dim pathRect As RECTF
    pathRect = Me.GetPathBoundariesF()
    
    'Update the transformation matrix with a matching mirror operation
    m_Transform.ApplyMirror mirrorHorizontal, mirrorVertical
    
    'Mirroring will reflect a path around the 0-axis, so we need to translate the path back into its original position now.
    If mirrorHorizontal Then m_Transform.ApplyTranslation (pathRect.Left + pathRect.Width / 2) * 2, 0, P2_TO_Append
    If mirrorVertical Then m_Transform.ApplyTranslation 0, (pathRect.Top + pathRect.Height / 2) * 2, P2_TO_Append
    
    'Apply the transformation
    Select Case m_PathBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            MirrorPathAroundItsCenter = GDI_Plus.GDIPlus_PathTransform(m_PathHandle, m_Transform.GetHandle)
    End Select
    
End Function

'Use this function to strip any interior paths and leave just the outline of the shape.  Winding mode rules are used,
' so discrete subpaths will be "outlined" individually.
'
'Also, this function requires an allowable error parameter because any curves in the image will be flattened
' (e.g. converted to a series of straight lines that approximate the original curve).
Public Function ConvertPath_OutlineOnly(Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    If (m_PathHandle <> 0) Then
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ConvertPath_OutlineOnly = GDI_Plus.GDIPlus_PathWindingModeOutline(m_PathHandle, 0&, allowableError)
        End Select
    Else
        InternalError "Path doesn't exist", "You can't convert an empty path!"
    End If
End Function

'This (poorly named?) function converts the current path to a new path, as if the path were traced by the supplied pen.
' All of the pen's settings (including things like line ends, joins, dash patterns, and more!) are considered when creating
' the new path, so this transform is potentially very powerful.
'
'Prior to conversion, any curves in the path will be converted to a series of lines.  The "allowable error" value determines
' how closely those lines approximate the original curve; a higher error means fewer lines and a faster transform, while a
' lower error means a more perfect approximation (at some cost to performance).
Public Function ConvertPath_PenTrace(ByRef srcPen As pd2DPen, Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    
    If (m_PathHandle <> 0) Then
        If (Not srcPen Is Nothing) Then
            Select Case m_PathBackend
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    ConvertPath_PenTrace = GDI_Plus.GDIPlus_PathWiden(m_PathHandle, srcPen.GetHandle, 0&, allowableError)
            End Select
        Else
            ConvertPath_PenTrace = False
        End If
    Else
        InternalError "Path doesn't exist", "You can't convert an empty path!"
    End If
    
End Function

'Inflate the current path by some arbitrary amount.  Note that additional parameters are available to control the way junctions and endpoints
' are affected by the widening.
'
'Unlike other subs, this actually returns a success/failure result.  GDI+ is prone to unpredictable failures, so the caller may need to take
' precautions if the request fails.  (On failure, this sub will automatically try to restore the original path if it can.)
Public Function ConvertPath_InflateLikeBalloon(ByVal inflateWidth As Single, Optional ByVal useLineJoin As PD_2D_LineJoin = P2_LJ_Round, Optional ByVal useLineCap As PD_2D_LineCap = P2_LC_Round, Optional ByVal allowableError As Single = DEFAULT_APPROXIMATION_ERROR) As Boolean
    
    If (m_PathHandle <> 0) Then
    
        'Widening a path requires a GDI+ pen.  All of the pen's properties (width, dashing, transforms, miters, etc)
        ' are considered when applying the widening transform.
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, inflateWidth, , , useLineJoin, useLineCap
        
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                
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
        
        End Select
        
    Else
        InternalError "Path doesn't exist", "You can't convert an empty path!"
    End If
    
End Function

Public Function CloneExistingPath(ByRef srcPath As pd2DPath) As Boolean
    
    If (Not srcPath Is Nothing) Then
        
        If (m_PathHandle <> 0) Then Me.ReleasePath
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                m_PathHandle = GDI_Plus.GDIPlus_PathClone(srcPath.GetHandle)
                CloneExistingPath = CBool(m_PathHandle <> 0)
        End Select
    
    Else
        InternalError "Null source path", "WARNING!  pd2DPath.CloneExistingPath was passed a non-existent path."
        CloneExistingPath = False
    End If
    
End Function

Public Function DoesPointTouchPathOutlineF(ByVal x As Single, ByVal y As Single, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) Then
        If (Not srcPen Is Nothing) Then
            Select Case m_PathBackend
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    DoesPointTouchPathOutlineF = GDI_Plus.GDIPlus_PathDoesPointTouchOutlineF(m_PathHandle, x, y, srcPen.GetHandle)
            End Select
        Else
            InternalError "Pen doesn't exist", "DoesPointTouchPathOutlineF is meaningless if you don't supply a pen!"
        End If
    Else
        InternalError "Path doesn't exist", "DoesPointTouchPathOutlineF is meaningless on an empty path!"
    End If
    
End Function

Public Function DoesPointTouchPathOutlineL(ByVal x As Long, ByVal y As Long, ByRef srcPen As pd2DPen) As Boolean
    
    If (m_PathHandle <> 0) Then
        If (Not srcPen Is Nothing) Then
            Select Case m_PathBackend
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    DoesPointTouchPathOutlineL = GDI_Plus.GDIPlus_PathDoesPointTouchOutlineL(m_PathHandle, x, y, srcPen.GetHandle)
            End Select
        Else
            InternalError "Pen doesn't exist", "DoesPointTouchPathOutlineF is meaningless if you don't supply a pen!"
        End If
    Else
        InternalError "Path doesn't exist", "DoesPointTouchPathOutlineL is meaningless on an empty path!"
    End If
    
End Function

'Retrieve this path's boundaries.  Optionally, you can also receive the path's boundaries as if they were drawn with
' a particular pen.  (All pen settings are considered, including tricky things like corner mitering.)
Friend Function GetPathBoundariesF(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RECTF
    
    If (m_PathHandle <> 0) Then
    
        Dim penHandle As Long, transformHandle As Long
        If (Not testPen Is Nothing) Then penHandle = testPen.GetHandle Else penHandle = 0
        If (Not testTransform Is Nothing) Then transformHandle = testTransform.GetHandle Else transformHandle = 0
        
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GetPathBoundariesF = GDI_Plus.GDIPlus_PathGetPathBoundsF(m_PathHandle, transformHandle, penHandle)
        End Select
        
    End If
    
End Function

Friend Function GetPathBoundariesL(Optional ByRef testPen As pd2DPen = Nothing, Optional ByVal testTransform As pd2DTransform = Nothing) As RECTL
        
    If (m_PathHandle <> 0) Then
    
        Dim penHandle As Long, transformHandle As Long
        If (Not testPen Is Nothing) Then penHandle = testPen.GetHandle Else penHandle = 0
        If (Not testTransform Is Nothing) Then transformHandle = testTransform.GetHandle Else transformHandle = 0
        
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GetPathBoundariesL = GDI_Plus.GDIPlus_PathGetPathBoundsL(m_PathHandle, transformHandle, penHandle)
        End Select
        
    End If

End Function

Public Function IsPointInsidePathF(ByVal x As Single, ByVal y As Single) As Boolean
    
    If (m_PathHandle <> 0) Then
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                IsPointInsidePathF = GDI_Plus.GDIPlus_PathIsPointInsideF(m_PathHandle, x, y)
        End Select
    Else
        InternalError "Path doesn't exist", "IsPointInsidePathF is meaningless on an empty path!"
    End If
    
End Function

Public Function IsPointInsidePathL(ByVal x As Long, ByVal y As Long) As Boolean
    
    If (m_PathHandle <> 0) Then
        Select Case m_PathBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                IsPointInsidePathL = GDI_Plus.GDIPlus_PathIsPointInsideL(m_PathHandle, x, y)
        End Select
    Else
        InternalError "Path doesn't exist", "IsPointInsidePathL is meaningless on an empty path!"
    End If
    
End Function

'As a convenience, this class exposes some generic "transform" operations that don't require the user to supply their own
' pd2DTransform object.  Instead, we use an internal instance that we reset between calls.
Private Sub ResetInternalTransform()
    If (m_Transform Is Nothing) Then Set m_Transform = New pd2DTransform
    m_Transform.Reset
End Sub

'You only need to manually initialize a backend if you plan on using a NON-DEFAULT backend.
' (Objects with the default backend are initialized automatically.)
Friend Function InitializePathBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_PathBackend = targetBackend
    InitializePathBackend = True
End Function

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.ResetPath
End Sub

Private Sub Class_Terminate()
    Me.ReleasePath
End Sub


'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DPath"
End Sub
