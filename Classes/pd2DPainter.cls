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
'Copyright 2012-2018 by Tanner Helland
'Created: 01/September/12
'Last updated: 28/June/16
'Last update: implement additional surface-to-surface rendering functions
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If possible (e.g. paints wihtout stretching), this painter class will drop back to bare AlphaBlend calls
' for image rendering.  This provides a meaningful improvement over GDI+ draw calls.
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Friend Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

'Quickly verify that source and target backends match (usually performed prior to a draw operation).
'IMPORTANT NOTE: as you can see from the m_DebugMode check inside this function, verification is *not* performed
'                when m_DebugMode is FALSE.  This is done for performance reasons.  If you want full verification
'                performed during non-debug-mode as well, you can remove the "If m_DebugMode..." check.
Private Function VerifyDrawBackends(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen) As Boolean
    
    If m_DebugMode Then
        If (Not dstSurface Is Nothing) And (Not srcPen Is Nothing) Then
            VerifyDrawBackends = (dstSurface.GetSurfaceBackend = srcPen.GetPenBackend)
            If (Not VerifyDrawBackends) Then InternalError "Mismatched backends", "Draw functions require matching surface and pen backends!"
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

'Copy functions.  Copying one surface onto another surface does *not* perform any blending.  It performs a wholesale
' replacement of the destination bytes with the source bytes.
Friend Function CopySurfaceI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByRef srcSurface As pd2DSurface) As Boolean
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    CopySurfaceI = GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than GDI.
Friend Function CopySurfaceF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByRef srcSurface As pd2DSurface) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceF = GDI_Plus.GDIPlus_DrawImageF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'These crop functions are identical to the ones above, except they allow the user to control source width/height instead of
' inferring it automatically.
Friend Function CopySurfaceCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    CopySurfaceCroppedI = GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
End Function

Friend Function CopySurfaceCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceCroppedF = GDI_Plus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, cropWidth, cropHeight)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'You might think we could just wrap StretchBlt here, but StretchBlt is inconsistent in its handling of alpha channels.
' GDI+ is actually pretty comparable speed-wise in nearest-neighbor mode, so this isn't a huge penalty.
Friend Function CopySurfaceResizedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedI = GDI_Plus.GDIPlus_DrawImageRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Friend Function CopySurfaceResizedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedF = GDI_Plus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Friend Function CopySurfaceResizedCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedCroppedF = GDI_Plus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Friend Function CopySurfaceResizedCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long) As Boolean
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedCroppedI = GDI_Plus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight)
    GDI_Plus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'Draw functions.  Given a target pd2dSurface object and a source pd2dPen, apply the pen to the surface in said shape.
Friend Function DrawArcF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawArcF = GDI_Plus.GDIPlus_DrawArcF(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
        End Select
    End If
End Function

Friend Function DrawArcI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawArcI = GDI_Plus.GDIPlus_DrawArcI(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
        End Select
    End If
End Function

Friend Function DrawCircleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    DrawCircleF = DrawEllipseF(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Friend Function DrawCircleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    DrawCircleI = DrawEllipseI(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Friend Function DrawEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawEllipseF = GDI_Plus.GDIPlus_DrawEllipseF(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Friend Function DrawEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    DrawEllipseF_AbsoluteCoords = Me.DrawEllipseF(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function DrawEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF) As Boolean
    DrawEllipseF_FromRectF = Me.DrawEllipseF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Friend Function DrawEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawEllipseI = GDI_Plus.GDIPlus_DrawRectI(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Friend Function DrawEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    DrawEllipseI_AbsoluteCoords = Me.DrawEllipseI(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function DrawEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectL) As Boolean
    DrawEllipseI_FromRectL = Me.DrawEllipseI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

'Drawing entire surfaces onto each other is significantly more convoluted than drawing shapes, because GDI+ Graphics
' objects do not support direct access to bits (which is actually forgivable, because a Graphics object may not have
' a raster object selected into it).  Instead, we must generate - on-the-fly - a GDI+ Image object as our
' "source image".  The surface class helps with this.
'
'Also, note that wherever possible we try to bypass GDI+ and just use GDI, which is totally sufficient for 24-bpp
' targets and/or integer-only coordinates.
Friend Function DrawSurfaceI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Because this function doesn't require stretching, we can drop back to AlphaBlend for improved performance.
    ' (This is only possible because pd2D operates in the premultiplied alpha space; if it didn't, we'd be forced
    ' to use slower GDI+ calls everywhere.)
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (srcSurface.GetSurfaceAlphaSupport Or (customOpacity <> 100)) Then
        DrawSurfaceI = AlphaBlendWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, srcWidth, srcHeight, srcSurface.GetSurfaceAlphaSupport, customOpacity * 2.55)
    Else
        DrawSurfaceI = GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
    End If
    
End Function

Private Function AlphaBlendWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcIs32bpp As Boolean = True, Optional ByVal blendOpacity As Long = 255) As Boolean

    Dim abParams As Long
    
    'Use the image's current alpha channel, and blend it with the supplied customAlpha value
    If srcIs32bpp Then
        abParams = blendOpacity * &H10000 Or &H1000000
        
        'If the source is a pdDIB object, we could actually test for premultiplication here (and in fact,
        ' pdDIB provides its own AlphaBlend wrapper that handles this for us).
    
    'Ignore alpha channel, and only use the supplied customAlpha value.
    Else
        
        ' (My memory is fuzzy after so many years, but I seem to recall old versions of Windows sometimes failing
        '  to AlphaBlend if the alpha value was exactly 255 - as a failsafe, let's use 254 as necessary.
        '  TODO: test this on XP, Win 7, Win 10 to confirm behavior.)
        If (blendOpacity = 255) Then blendOpacity = 254
        abParams = (blendOpacity * &H10000)
    End If
    
    AlphaBlend hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, srcWidth, srcHeight, abParams
    
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than GDI.
Friend Function DrawSurfaceF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    'Custom opacity requires a totally different (and far more complicated) GDI+ function
    If (customOpacity <> 100#) Then
        DrawSurfaceF = GDI_Plus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, srcWidth, srcHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceF = GDI_Plus.GDIPlus_DrawImageF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY)
    End If
    
End Function

Friend Function DrawSurfaceCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Because this function doesn't require stretching, we can drop back to AlphaBlend for improved performance.
    ' (This is only possible because pd2D operates in the premultiplied alpha space; if it didn't, we'd be forced
    ' to use slower GDI+ calls everywhere.)
    If (srcSurface.GetSurfaceAlphaSupport Or (customOpacity <> 100)) Then
        DrawSurfaceCroppedI = AlphaBlendWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, srcX, srcY, cropWidth, cropHeight, srcSurface.GetSurfaceAlphaSupport, customOpacity * 2.55)
    Else
        DrawSurfaceCroppedI = GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, srcX, srcY, vbSrcCopy)
    End If
    
End Function

Friend Function DrawSurfaceCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal cropWidth As Single, ByVal cropHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceCroppedF = GDI_Plus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, cropWidth, cropHeight, srcX, srcY, cropWidth, cropHeight, customOpacity * 0.01)
End Function

Friend Function DrawSurfaceResizedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (customOpacity <> 100#) Then
        DrawSurfaceResizedI = GDI_Plus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceResizedI = GDI_Plus.GDIPlus_DrawImageRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    End If
    
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than GDI.
Friend Function DrawSurfaceResizedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (customOpacity <> 100#) Then
        DrawSurfaceResizedF = GDI_Plus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceResizedF = GDI_Plus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    End If
    
End Function

Friend Function DrawSurfaceResizedCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceResizedCroppedF = GDI_Plus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
End Function

Friend Function DrawSurfaceResizedCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceResizedCroppedI = GDI_Plus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
End Function

Friend Function DrawSurfaceRotatedF(ByRef dstSurface As pd2DSurface, ByVal dstCenterX As Single, ByVal dstCenterY As Single, ByVal rotateAngle As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Create a transform that describes the rotation
    Dim cTransform As pd2DTransform: Set cTransform = New pd2DTransform
    cTransform.ApplyTranslation -1 * (srcX + (srcX + srcWidth)) / 2, -1 * (srcY + (srcY + srcHeight)) / 2
    cTransform.ApplyRotation rotateAngle
    cTransform.ApplyTranslation dstCenterX, dstCenterY
    
    'Translate the corner points of the image to match.  (Note that the order of points is important; GDI+ requires points
    ' in top-left, top-right, bottom-left order, with the fourth point being optional.)
    Dim imgCorners() As PointFloat
    ReDim imgCorners(0 To 3) As PointFloat
    imgCorners(0).x = srcX
    imgCorners(0).y = srcY
    imgCorners(1).x = srcX + srcWidth
    imgCorners(1).y = srcY
    imgCorners(2).x = srcX
    imgCorners(2).y = srcY + srcHeight
    
    cTransform.ApplyTransformToPointFs VarPtr(imgCorners(0)), 3
    
    'Draw the image, using the new corner points as the destination!
    DrawSurfaceRotatedF = GDI_Plus.GDIPlus_DrawImagePointsRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, imgCorners, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
    
End Function

Friend Function DrawSurfaceTransformedF(ByRef dstSurface As pd2DSurface, ByRef srcSurface As pd2DSurface, ByRef srcTransform As pd2DTransform, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Translate the corner points of the image to match.  (Note that the order of points is important; GDI+ requires points
    ' in top-left, top-right, bottom-left order, with the fourth point being optional.)
    Dim imgCorners() As PointFloat
    ReDim imgCorners(0 To 3) As PointFloat
    imgCorners(0).x = srcX
    imgCorners(0).y = srcY
    imgCorners(1).x = srcX + srcWidth
    imgCorners(1).y = srcY
    imgCorners(2).x = srcX
    imgCorners(2).y = srcY + srcHeight
    
    srcTransform.ApplyTransformToPointFs VarPtr(imgCorners(0)), 3
    
    'Draw the image, using the new corner points as the destination!
    DrawSurfaceTransformedF = GDI_Plus.GDIPlus_DrawImagePointsRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, imgCorners, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
    
End Function

Friend Function DrawLineF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineF = GDI_Plus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
                If (Not DrawLineF) Then InternalError "LineDrawFailed", "GDIPlus_DrawLineF drew nothing"
        End Select
    End If
End Function

Friend Function DrawLineF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPoint1 As PointFloat, ByRef srcPoint2 As PointFloat) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineF_FromPtF = GDI_Plus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, srcPoint1.x, srcPoint1.y, srcPoint2.x, srcPoint2.y)
        End Select
    End If
End Function

Friend Function DrawLinesF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                If useCurveAlgorithm Then
                    DrawLinesF_FromPtF = GDI_Plus.GDIPlus_DrawCurveF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension)
                Else
                    DrawLinesF_FromPtF = GDI_Plus.GDIPlus_DrawLinesF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints)
                End If
        End Select
    End If
End Function

Friend Function DrawLineI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineI = GDI_Plus.GDIPlus_DrawLineI(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
        End Select
    End If
End Function

Friend Function DrawLineI_FromPtL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPoint1 As PointLong, ByRef srcPoint2 As PointLong) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawLineI_FromPtL = GDI_Plus.GDIPlus_DrawLineI(dstSurface.GetHandle, srcPen.GetHandle, srcPoint1.x, srcPoint1.y, srcPoint2.x, srcPoint2.y)
        End Select
    End If
End Function

Friend Function DrawLinesI_FromPtL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                If useCurveAlgorithm Then
                    DrawLinesI_FromPtL = GDI_Plus.GDIPlus_DrawCurveI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints, curvatureTension)
                Else
                    DrawLinesI_FromPtL = GDI_Plus.GDIPlus_DrawLinesI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints)
                End If
        End Select
    End If
End Function

Friend Function DrawPath(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPath As pd2DPath) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawPath = GDI_Plus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, srcPath.GetHandle)
        End Select
    End If
End Function

'Helper function; the source path is silently cloned and transformed, leaving the original path untouched
Friend Function DrawPath_Transformed(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPath As pd2DPath, ByRef srcTransform As pd2DTransform) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                Dim tmpPath As pd2DPath: Set tmpPath = New pd2DPath
                tmpPath.CloneExistingPath srcPath
                tmpPath.ApplyTransformation srcTransform
                DrawPath_Transformed = GDI_Plus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, tmpPath.GetHandle)
        End Select
    End If
End Function

Friend Function DrawPolygonF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                If useCurveAlgorithm Then
                    DrawPolygonF = GDI_Plus.GDIPlus_DrawClosedCurveF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension)
                Else
                    DrawPolygonF = GDI_Plus.GDIPlus_DrawPolygonF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints)
                End If
        End Select
    End If
End Function

Friend Function DrawPolygonI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5)
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                If useCurveAlgorithm Then
                    DrawPolygonI = GDI_Plus.GDIPlus_DrawClosedCurveI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints, curvatureTension)
                Else
                    DrawPolygonI = GDI_Plus.GDIPlus_DrawPolygonI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints)
                End If
        End Select
    End If
End Function

Friend Function DrawRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawRectangleF = GDI_Plus.GDIPlus_DrawRectF(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Friend Function DrawRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    DrawRectangleF_AbsoluteCoords = Me.DrawRectangleF(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function DrawRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF) As Boolean
    DrawRectangleF_FromRectF = Me.DrawRectangleF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Friend Function DrawRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawRectangleI = GDI_Plus.GDIPlus_DrawRectI(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Friend Function DrawRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    DrawRectangleI_AbsoluteCoords = Me.DrawRectangleI(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function DrawRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectL) As Boolean
    DrawRectangleI_FromRectL = Me.DrawRectangleI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Friend Function DrawRoundRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.SetDebugMode Drawing2D.GetLibraryDebugMode
    tmpPath.AddRoundedRectangle_RectF srcRect, cornerRadius
    
    If VerifyDrawBackends(dstSurface, srcPen) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                DrawRoundRectangleF_FromRectF = GDI_Plus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, tmpPath.GetHandle)
        End Select
    End If
    
End Function

'Fill functions.  Given a target pd2dSurface and a source pd2dBrush, apply the brush to the surface in said shape.

Friend Function FillCircleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    FillCircleF = FillEllipseF(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Friend Function FillCircleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    FillCircleI = FillEllipseI(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Friend Function FillEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillEllipseF = GDI_Plus.GDIPlus_FillEllipseF(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Friend Function FillEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    FillEllipseF_AbsoluteCoords = Me.FillEllipseF(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function FillEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF) As Boolean
    FillEllipseF_FromRectF = Me.FillEllipseF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Friend Function FillEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillEllipseI = GDI_Plus.GDIPlus_FillRectI(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
        End Select
    End If
End Function

Friend Function FillEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    FillEllipseI_AbsoluteCoords = Me.FillEllipseI(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Friend Function FillEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectL) As Boolean
    FillEllipseI_FromRectL = Me.FillEllipseI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Friend Function FillPath(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcPath As pd2DPath) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillPath = GDI_Plus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, srcPath.GetHandle)
        End Select
    End If
End Function

'Helper function; the source path is silently cloned and transformed, leaving the original path untouched
Friend Function FillPath_Transformed(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcPath As pd2DPath, ByRef srcTransform As pd2DTransform) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                Dim tmpPath As pd2DPath: Set tmpPath = New pd2DPath
                tmpPath.CloneExistingPath srcPath
                tmpPath.ApplyTransformation srcTransform
                FillPath_Transformed = GDI_Plus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
        End Select
    End If
End Function

Friend Function FillPolygonF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5, Optional ByVal fillMode As PD_2D_FillRule = P2_FR_Winding)
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                If useCurveAlgorithm Then
                    FillPolygonF_FromPtF = GDI_Plus.GDIPlus_FillClosedCurveF(dstSurface.GetHandle, srcBrush.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension, fillMode)
                Else
                    FillPolygonF_FromPtF = GDI_Plus.GDIPlus_FillPolygonF(dstSurface.GetHandle, srcBrush.GetHandle, ptrToPtFArray, numOfPoints, fillMode)
                End If
        End Select
    End If
End Function

Friend Function FillRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRectangleF = GDI_Plus.GDIPlus_FillRectF(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Friend Function FillRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    FillRectangleF_AbsoluteCoords = Me.FillRectangleF(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function FillRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF) As Boolean
    FillRectangleF_FromRectF = Me.FillRectangleF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Friend Function FillRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRectangleI = GDI_Plus.GDIPlus_FillRectI(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
        End Select
    End If
End Function

Friend Function FillRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    FillRectangleI_AbsoluteCoords = Me.FillRectangleI(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Friend Function FillRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectL) As Boolean
    FillRectangleI_FromRectL = Me.FillRectangleI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Friend Function FillRoundRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal x As Single, ByVal y As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.SetDebugMode Drawing2D.GetLibraryDebugMode
    tmpPath.AddRoundedRectangle_Relative x, y, rectWidth, rectHeight, cornerRadius
    
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRoundRectangleF = GDI_Plus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
        End Select
    End If
    
End Function

Friend Function FillRoundRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.SetDebugMode Drawing2D.GetLibraryDebugMode
    tmpPath.AddRoundedRectangle_RectF srcRect, cornerRadius
    
    If VerifyFillBackends(dstSurface, srcBrush) Then
        Select Case dstSurface.GetSurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                FillRoundRectangleF_FromRectF = GDI_Plus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
        End Select
    End If
    
End Function

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DPainter"
End Sub

Private Sub Class_Initialize()
    m_DebugMode = Drawing2D.GetLibraryDebugMode()
End Sub
