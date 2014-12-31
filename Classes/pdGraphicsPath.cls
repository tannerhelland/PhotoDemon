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
'Copyright 2013-2014 by Tanner Helland
'Created: 17/October/14 (though assembled from various parts written much earlier)
'Last updated: 17/October/14
'Last update: initial build
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
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal penColor As Long, ByVal penWidth As Single, ByVal penUnit As GpUnit, ByRef hPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As Long

'Path property changes (both destructive and non-destructive)
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal newBrushMode As GDIFillMode) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newSmoothingMode As SmoothingMode) As Long
Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal widenFlatness As Single) As Long
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal penStartCap As LineCap, ByVal penEndCap As LineCap, ByVal penDashCap As DashCap) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal penLineJoin As LineJoin) As Long

'Figure-specific path additions
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Long
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Long
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As Long

'Drawing functions
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As Long

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
    GdipCreatePen1 0, borderWidth, UnitPixel, hPen
    
    'Apply the requested line cap and miter
    GdipSetPenLineCap hPen, useLineCap, useLineCap, DashCapFlat
    GdipSetPenLineJoin hPen, useLineJoin
    
    'Apply the widen command
    GdipWidenPath m_PathHandle, hPen, 0, allowableError
    
    'Release the temporary pen
    GdipDeletePen hPen

End Sub


' END: DESTRUCTIVE PATH EDITS
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
    GdipDeletePath m_PathHandle

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
' BEGIN: DRAWING FUNCTIONS

'Stroke the current path, using the default arrangement for PD's current theme
Public Sub strokePathToDIB_UIStyle(ByRef dstDIB As pdDIB)

    'PD's default behavior for UI-related paths is to stroke a path twice; once with a thickness of 2px
    ' (which actually ends up being more like 1px along the path, and 0.5 px along either side)
    ' using a black pen at 75% opacity, then again at a thickness of 1px at 75% opacity.  This provides
    ' a sharp line regardless of the underlying image.
    
    'Antialiasing is always used.  Note that this is relevant even for shapes like rectangles, in order
    ' to split the 2px line across 3 pixels.
    
    'Obtain a GDI+ handle to the image
    Dim hGraphics As Long
    GdipCreateFromHDC dstDIB.getDIBDC, hGraphics
    
    'Activate AA
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    'Create two pens, each of which will be used to stroke the line
    Dim hBlackPen As Long, hWhitePen As Long
    GdipCreatePen1 fillQuadWithVBRGB(RGB(0, 0, 0), 192), 2, UnitPixel, hBlackPen
    GdipCreatePen1 fillQuadWithVBRGB(RGB(255, 255, 255), 192), 1, UnitPixel, hWhitePen
    
    'Stroke the paths in turn
    GdipDrawPath hGraphics, hBlackPen, m_PathHandle
    GdipDrawPath hGraphics, hWhitePen, m_PathHandle
        
    'Release the two pens and DIB handle
    GdipDeletePen hBlackPen
    GdipDeletePen hWhitePen
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
