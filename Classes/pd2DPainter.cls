VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DPainter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Painting class (interface for using pd2dBrush and pd2dPen on pd2dSurface objects)
'Copyright 2012-2016 by Tanner Helland
'Created: 01/September/12
'Last updated: 16/May/16
'Last update: continue migratation of paint operations out of GDI+ module and into this class
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

'Quickly verify that source and target backends match (usually performed prior to a draw operation).
'IMPORTANT NOTE: as you can see from the m_DebugMode check inside this function, verification is *not* performed
'                when m_DebugMode is FALSE.  This is done for performance reasons.  If you want full verification
'                performed during non-debug-mode as well, you can remove the "If m_DebugMode..." check.
Private Function VerifyDrawBackends(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen) As Boolean
    
    If m_DebugMode Then
        If (Not (dstSurface Is Nothing)) And (Not (srcPen Is Nothing)) Then
            If (dstSurface.GetSurfaceBackend <> srcPen.GetPenBackend) Then
                InternalError "Mismatched backends", "Draw functions require matching surface and pen backends!"
            Else
                VerifyDrawBackends = True
            End If
        Else
            InternalError "Null object", "You can't pass null objects to draw functions and expect them to work!"
        End If
    Else
        VerifyDrawBackends = True
    End If
    
End Function

Private Function VerifyFillBackends(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush) As Boolean
    
    If m_DebugMode Then
        If (Not (dstSurface Is Nothing)) And (Not (srcBrush Is Nothing)) Then
            If (dstSurface.GetSurfaceBackend <> srcBrush.GetBrushBackend) Then
                InternalError "Mismatched backends", "Fill functions require matching surface and brush backends!"
            Else
                VerifyFillBackends = True
            End If
        Else
            InternalError "Null object", "You can't pass null objects to fill functions and expect them to work!"
        End If
    Else
        VerifyFillBackends = True
    End If
    
End Function

'Draw functions.  Given a target pd2dSurface object and a source pd2dPen, apply the pen to the surface in said shape.
Public Function DrawArcF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawArcF = GDI_Plus.GDIPlus_DrawArcF(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
        End Select
    End If
End Function

Public Function DrawArcI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawArcI = GDI_Plus.GDIPlus_DrawArcI(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
        End Select
    End If
End Function

Public Function DrawCircleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    DrawCircleF = DrawEllipseF(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function DrawCircleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    DrawCircleI = DrawEllipseI(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function DrawEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawEllipseF = GDI_Plus.GDIPlus_DrawEllipseF(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Public Function DrawEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    DrawEllipseF_AbsoluteCoords = Me.DrawEllipseF(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function DrawEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RECTF) As Boolean
    DrawEllipseF_FromRectF = Me.DrawEllipseF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function DrawEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawEllipseI = GDI_Plus.GDIPlus_DrawRectI(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Public Function DrawEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    DrawEllipseI_AbsoluteCoords = Me.DrawEllipseI(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function DrawEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RECTL) As Boolean
    DrawEllipseI_FromRectL = Me.DrawEllipseI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Public Function DrawLineF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineF = GDI_Plus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
        End Select
    End If
End Function

Friend Function DrawLineF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPoint1 As POINTFLOAT, ByRef srcPoint2 As POINTFLOAT) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineF_FromPtF = GDI_Plus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, srcPoint1.x, srcPoint1.y, srcPoint2.x, srcPoint2.y)
        End Select
    End If
End Function

Public Function DrawLineI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineI = GDI_Plus.GDIPlus_DrawLineI(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
        End Select
    End If
End Function

Public Function DrawPath(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPath As pd2DPath) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawPath = GDI_Plus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, srcPath.GetHandle)
        End Select
    End If
End Function

Public Function DrawRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawRectangleF = GDI_Plus.GDIPlus_DrawRectF(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Public Function DrawRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    DrawRectangleF_AbsoluteCoords = Me.DrawRectangleF(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function DrawRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RECTF) As Boolean
    DrawRectangleF_FromRectF = Me.DrawRectangleF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function DrawRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawRectangleI = GDI_Plus.GDIPlus_DrawRectI(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Public Function DrawRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    DrawRectangleI_AbsoluteCoords = Me.DrawRectangleI(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function DrawRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RECTL) As Boolean
    DrawRectangleI_FromRectL = Me.DrawRectangleI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

'Fill functions.  Given a target pd2dSurface and a source pd2dBrush, apply the brush to the surface in said shape.

Public Function FillCircleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    FillCircleF = FillEllipseF(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function FillCircleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    FillCircleI = FillEllipseI(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function FillEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillEllipseF = GDI_Plus.GDIPlus_FillEllipseF(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Public Function FillEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    FillEllipseF_AbsoluteCoords = Me.FillEllipseF(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function FillEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RECTF) As Boolean
    FillEllipseF_FromRectF = Me.FillEllipseF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function FillEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillEllipseI = GDI_Plus.GDIPlus_FillRectI(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Public Function FillEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    FillEllipseI_AbsoluteCoords = Me.FillEllipseI(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function FillEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RECTL) As Boolean
    FillEllipseI_FromRectL = Me.FillEllipseI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Public Function FillPath(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcPath As pd2DPath) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillPath = GDI_Plus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, srcPath.GetHandle)
        End Select
    End If
End Function

Public Function FillRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRectangleF = GDI_Plus.GDIPlus_FillRectF(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Public Function FillRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    FillRectangleF_AbsoluteCoords = Me.FillRectangleF(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function FillRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RECTF) As Boolean
    FillRectangleF_FromRectF = Me.FillRectangleF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function FillRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRectangleI = GDI_Plus.GDIPlus_FillRectI(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Public Function FillRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    FillRectangleI_AbsoluteCoords = Me.FillRectangleI(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function FillRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RECTL) As Boolean
    FillRectangleI_FromRectL = Me.FillRectangleI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function


'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DPainter"
End Sub
